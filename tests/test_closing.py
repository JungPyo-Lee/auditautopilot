"""주석 자료요청 양식, 감사보고서 초안. 가짜 회사·숫자만 쓴다."""

import io
import zipfile

import openpyxl
import pytest

from auditautopilot.cli import main
from auditautopilot.workpaper import adjust, notes, report

from test_adjust import AJE_ROWS, _aje, bs_prior, filled  # noqa: F401
from test_douzone import dz  # noqa: F401
from test_leadsheet import tb_file, template  # noqa: F401
from test_manage import _open, server, store  # noqa: F401


@pytest.fixture
def adjusted(filled, tmp_path):  # noqa: F811
    wb = openpyxl.load_workbook(filled)
    ws = wb["일반사항"]
    for r, (k, v) in enumerate([("회사 명", "가나상사 주식회사"), ("회계법인 명", "다라회계법인"),
                                ("당기 결산일", "2026-12-31"), ("전기 결산일", "2025-12-31"),
                                ("감사종료일", "2027-03-15")], start=5):
        ws.cell(r, 3, k)
        ws.cell(r, 4, __import__("datetime").date.fromisoformat(v) if v[:2] == "20" else v)
    wb.save(filled)
    aje = _aje(tmp_path, filled, AJE_ROWS)
    return adjust.apply(filled, aje, tmp_path / "반영.xlsx").output


def test_figures(adjusted):
    figs, ni, extra = adjust.figures(adjusted)
    by = {f.row.label: f for f in figs}
    assert ni == 190000 and not extra
    assert by["대손충당금(외상매출금)"].current == -60000 and by["대손충당금(외상매출금)"].prior == -50000
    assert by["대손상각비"].adjustment == 10000 and by["외상매입금"].current == 250000


def test_notes_request(adjusted, tmp_path):
    res = notes.request_workbook(adjusted, tmp_path / "주석.xlsx")
    assert "매출채권과 대손충당금" in res.filled and "대손충당금 변동" in res.filled
    assert "유형자산 변동" not in res.filled                          # 유형자산 계정이 없으면 빠진다
    assert "특수관계자 현황·거래" in res.company and "토지 공시지가" not in res.company
    wb = openpyxl.load_workbook(res.output)
    assert wb.sheetnames[0] == "목차" and wb["목차"]["A1"].value.startswith("가나상사 주식회사")
    ar = wb[next(n for n in wb.sheetnames if "매출채권" in n)]
    rows = [r for r in ar.iter_rows(min_row=5, values_only=True) if r[0]]
    assert rows[0][:3] == ("외상매출금", 1000000, 900000) and rows[1][:3] == ("대손충당금(외상매출금)", -60000, -50000)
    mv = wb[next(n for n in wb.sheetnames if "대손충당금 변동" in n)]
    assert [c.value for c in mv[5]][:6] == ["대손충당금(외상매출금)", 50000, None, None, None, 60000]
    kinds = [r[2] for r in wb["목차"].iter_rows(min_row=6, values_only=True) if r[2]]
    assert kinds == sorted(kinds, key=lambda k: k != notes.KIND_FILLED)   # 정산표 채움 먼저


def test_report_docx(adjusted, tmp_path):
    out, extra = report.draft(adjusted, tmp_path / "r.docx")
    assert not extra
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    assert "독립된 감사인의 감사보고서" in xml and "가나상사 주식회사" in xml and "다라회계법인" in xml
    assert "2026년 12월 31일 현재의 재무상태" in xml and "2027년 3월 15일" in xml
    assert "1,440,000" in xml and "190,000" in xml                          # 자산총계, 당기순이익
    q, _ = report.draft(adjusted, tmp_path / "q.docx", "한정")
    qx = zipfile.ZipFile(q).read("word/document.xml").decode()
    assert "한정의견근거" in qx and "미칠 수 있는 영향을 제외하고는" in qx
    with pytest.raises(ValueError, match="감사의견"):
        report.draft(adjusted, tmp_path / "x.docx", "보류")


def test_cli_and_web(adjusted, tmp_path, capsys, server, store):  # noqa: F811
    assert main(["notes", str(adjusted), "-o", str(tmp_path / "n.xlsx")]) == 0
    assert main(["report", str(adjusted), "--opinion", "적정", "-o", str(tmp_path / "r.docx")]) == 0
    assert "회사 작성" in capsys.readouterr().out
    base, _ = server
    cid = store.add_client("C013", "가나상사", "김담당", "kim@example.com")
    eid = store.add_engagement(cid, 2026, "기말감사", "2027-01-11", "2027-01-15")
    assert _open(base + f"/e/{eid}/notes-request.xlsx")[0] == 400          # 수정분개 반영 전
    store.set_setting(f"adjust:{eid}", {"made_at": "x", "path": str(adjusted), "lines": []})
    st, _, x = _open(base + f"/e/{eid}/notes-request.xlsx")
    assert st == 200 and "목차" in openpyxl.load_workbook(io.BytesIO(x)).sheetnames
    st, h, d = _open(base + f"/e/{eid}/report.docx?opinion=%ED%95%9C%EC%A0%95")
    assert st == 200 and "wordprocessingml" in h["Content-Type"]
    assert "한정의견" in zipfile.ZipFile(io.BytesIO(d)).read("word/document.xml").decode()
    page = _open(base + f"/e/{eid}")[2].decode()
    assert "주석 자료요청 양식 내려받기" in page and "재무제표·주석 초안" in page and "감사의견" not in page


def _fill_notes(path):
    """회사가 주석 자료를 채운 것처럼 꾸민다 (엑셀로 저장하지 않아 합계 수식 값은 없다)."""
    wb = openpyxl.load_workbook(path)
    ar = wb[next(n for n in wb.sheetnames if "매출채권과" in n)]
    ar["D5"], ar["H5"] = 700000, "정상"
    rp = wb[next(n for n in wb.sheetnames if "특수관계자" in n)]
    rp["A5"], rp["B5"], rp["C5"] = "가나홀딩스", "지배기업", 120000
    lit = wb[next(n for n in wb.sheetnames if "소송" in n)]
    lit["A5"] = "없음"
    wb.save(path)
    return path


def test_fs_docx(adjusted, tmp_path):
    from auditautopilot.workpaper import fs_doc
    figs, ni, _ = adjust.figures(adjusted)
    _, tot = fs_doc.balance_sheet(figs, ni)
    assert tot["자산총계"] == tot["부채와자본총계"] == (1440000, 1250000)
    cf, info = fs_doc.cash_flow(figs, ni)
    assert info["increase"] == 100000 and info["check"] == 0
    heads, eq = fs_doc.equity_changes(figs, ni)
    assert heads[-1] == "합계" and eq[1][-1] == 190000 and eq[3][-1] == tot["자본총계"][0]

    blank = fs_doc.draft(adjusted, tmp_path / "a.docx")
    assert any("주석 자료가 없어" in n for n in blank.notes)
    req = notes.request_workbook(adjusted, tmp_path / "요청.xlsx").output
    items = fs_doc.read_filled_notes(_fill_notes(req))
    by = {n.title: n for n in items}
    ar = by["매출채권과 대손충당금"]
    assert ar.head[:4] == ["과목", "당기말(정산표)", "전기말", "3개월 이내"] and "6개월~1년" not in ar.head   # 빈 열은 뺀다
    assert ar.rows[-1][:2] == ["합계", 940000]                                    # 수식 값이 없어도 다시 더한다
    assert by["특수관계자 현황·거래"].rows[0][:3] == ["가나홀딩스", "지배기업", 120000]
    assert "해당사항이 없습니다" in by["소송·우발부채"].text
    assert by["보험 가입 내역"].text.startswith("[")
    res = fs_doc.draft(adjusted, tmp_path / "f.docx", req)
    assert res.notes_used == "요청.xlsx" and not res.notes
    xml = zipfile.ZipFile(res.output).read("word/document.xml").decode()
    for word in ("재 무 상 태 표", "손 익 계 산 서", "자 본 변 동 표", "현 금 흐 름 표", "가나홀딩스", "중요한 회계정책"):
        assert word in xml
    assert "1,440,000" in xml and "(60,000)" in xml and "독립된 감사인" not in xml
    assert main(["fs", str(adjusted), "--notes", str(req), "-o", str(tmp_path / "cli.docx")]) == 0


def test_web_notes_upload_and_fs(adjusted, tmp_path, server, store):  # noqa: F811
    from test_manage import _multipart
    base, _ = server
    cid = store.add_client("C014", "가나상사", "김담당", "kim@example.com")
    eid = store.add_engagement(cid, 2026, "기말감사", "2027-01-11", "2027-01-15")
    store.set_setting(f"adjust:{eid}", {"made_at": "x", "path": str(adjusted), "lines": []})
    req = _fill_notes(notes.request_workbook(adjusted, tmp_path / "요청.xlsx").output)
    data, hdr = _multipart([], [("notes", "가나_주석자료.xlsx", req.read_bytes())])
    st, h, _ = _open(base + f"/e/{eid}/notes-filled", data, hdr)
    assert st == 303 and "#closing" in h["Location"]
    assert store.get_setting(f"notes_filled:{eid}")["name"] == "가나_주석자료.xlsx"
    st, h, d = _open(base + f"/e/{eid}/fs.docx")
    assert st == 200 and "wordprocessingml" in h["Content-Type"]
    assert "가나홀딩스" in zipfile.ZipFile(io.BytesIO(d)).read("word/document.xml").decode()
    assert "가나_주석자료.xlsx" in _open(base + f"/e/{eid}")[2].decode()
    data, hdr = _multipart([], [("notes", "엉뚱.xlsx", adjusted.read_bytes())])
    st, h, _ = _open(base + f"/e/{eid}/notes-filled", data, hdr)
    assert st == 303 and "err=" in h["Location"]


def test_tabs_and_progress(adjusted, server, store):  # noqa: F811
    base, _ = server
    cid = store.add_client("C015", "가나상사", "김담당", "kim@example.com")
    eid = store.add_engagement(cid, 2026, "기말감사", "2027-01-11", "2027-01-15")
    page = _open(base + f"/e/{eid}")[2].decode()
    for k in ("mail", "pbc", "files", "questions", "confirm", "lead", "close"):
        assert f"id='tab-{k}'" in page and f"href='#tab-{k}'" in page
    assert page.count("class='no'") == 7 and "진행 0/7" in page
    store.set_setting(f"leadsheet:{eid}", {"made_at": "x", "path": str(adjusted), "lines": []})
    store.set_setting(f"adjust:{eid}", {"made_at": "x", "path": str(adjusted), "lines": []})
    page = _open(base + f"/e/{eid}")[2].decode()
    assert "진행 1/7" in page and "주석 자료를 아직 안 받음" in page
    home = _open(base + "/")[2].decode()
    assert "ministeps" in home and f"/e/{eid}#tab-close" in home
