"""수정분개 → 정산표·수정후BS·IS·현금흐름정산표. 가짜 회사·숫자만 쓴다."""

import io
import shutil
import subprocess
import urllib.parse

import openpyxl
import pytest

from auditautopilot.cli import main
from auditautopilot.workpaper import adjust, leadsheet

from test_douzone import PERIOD, PRIOR, _save, dz  # noqa: F401
from test_leadsheet import tb_file, template  # noqa: F401
from test_manage import _multipart, _open, server, store  # noqa: F401


@pytest.fixture
def bs_prior(tmp_path):
    """시산표 계정과 같은 과목의 전기 재무상태표 (대차 일치)."""
    return _save(tmp_path, "재무상태표_20261031.xlsx", [
        ["과목", PERIOD, None, PRIOR, None],
        [None, "금액", None, "금액", None],
        ["자산", "", "", "", ""],
        ["    Ⅰ. 유동자산", "", "1,450,000", "", "1,250,000"],
        ["        (1)  당좌자산", "", "1,450,000", "", "1,250,000"],
        ["            보통예금", "", "500,000", "", "400,000"],
        ["            외상매출금", "1,000,000", "", "900,000", ""],
        ["            대손충당금", "50,000", "950,000", "50,000", "850,000"],
        ["자산총계", "", "1,450,000", "", "1,250,000"],
        ["부채", "", "", "", ""],
        ["    Ⅰ. 유동부채", "", "250,000", "", "200,000"],
        ["            외상매입금", "", "250,000", "", "200,000"],
        ["부채총계", "", "250,000", "", "200,000"],
        ["자본", "", "", "", ""],
        ["    Ⅲ. 이익잉여금", "", "1,000,000", "", "1,050,000"],
        ["            미처분이익잉여금", "", "1,000,000", "", "1,050,000"],
    ])


@pytest.fixture
def filled(template, tb_file, bs_prior, tmp_path):  # noqa: F811
    out = tmp_path / "out" / "조서_리드시트.xlsx"
    leadsheet.build(template, [tb_file, bs_prior], out)
    return out


def _aje(tmp_path, filled, rows):
    t = adjust.aje_template(filled, tmp_path / "양식.xlsx")
    wb = openpyxl.load_workbook(t)
    ws = wb["수정분개"]
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            ws.cell(5 + i, 1 + j, v)
    p = tmp_path / "수정분개.xlsx"
    wb.save(p)
    return p


AJE_ROWS = [(1, "2026-12-31", "수정", "대손상각비", 10000, None, "대손 추가 설정", "C1000"),
            (1, None, "수정", "대손충당금(외상매출금)", None, 10000)]


def test_template_lists_unique_accounts(filled, tmp_path):
    wb = openpyxl.load_workbook(adjust.aje_template(filled, tmp_path / "t.xlsx"))
    names = [r[0] for r in wb["계정목록"].iter_rows(min_row=2, values_only=True) if r[0]]
    assert "대손충당금(외상매출금)" in names and "외상매출금" in names and "복리후생비" in names
    groups = [r[4] for r in wb["계정목록"].iter_rows(min_row=2, values_only=True) if r[4]]
    assert "판매관리비 > 판매관리비" in groups
    assert [c.value for c in wb["수정분개"][4]][:6] == ["번호", "일자", "구분", "계정과목", "차변", "대변"]


def test_apply_adds_tabs_and_formulas(filled, tmp_path):
    aje = _aje(tmp_path, filled, AJE_ROWS)
    res = adjust.apply(filled, aje, tmp_path / "반영.xlsx")
    assert res.entries == 1 and res.debit == res.credit == 10000 and not res.problems
    assert any("대손상각비" in n and "판매관리비" in n for n in res.notes)
    wb = openpyxl.load_workbook(res.output)
    assert wb.sheetnames[-5:] == list(adjust.NEW_TABS)
    assert wb["AJE"]["E5"].number_format == "#,##0" and wb["AJE"]["E5"].value == 10000   # 새 서식 번호가 맞는지
    assert wb["AJE"]["A4"].font.b and wb["정산표"]["A1"].font.sz == 13
    ws = wb["정산표"]
    rows = {ws.cell(r, 3).value: r for r in range(5, ws.max_row + 1) if ws.cell(r, 3).value}
    r = rows["대손충당금(외상매출금)"]
    assert ws[f"E{r}"].value.startswith("=N(TB!H") and "SUMIFS('AJE'!$F:$F" in ws[f"G{r}"].value
    assert ws.cell(rows["대손상각비"], 4).value == 0                                   # 새 계정은 수정전 0
    assert ws[f"N{r}"].value == f"=F{r}-G{r}"                                          # 자산 묶음 부호
    lead = wb["C1000"]
    i_formulas = [lead.cell(x, 9).value for x in range(12, 16)]
    assert any(f and "'정산표'!N" in f for f in i_formulas) and all(f and f.startswith("=H") for f in i_formulas)
    cf = wb["현금흐름정산표"]
    labels = [cf.cell(x, 1).value for x in range(1, cf.max_row + 1)]
    assert "  가산: 대손상각비" in labels and any(x and "기말의 현금" in x for x in labels)
    # 다시 반영하면 탭을 새로 만든다 (두 벌 생기지 않음)
    again = adjust.apply(res.output, aje, tmp_path / "반영2.xlsx")
    assert openpyxl.load_workbook(again.output).sheetnames.count("정산표") == 1
    assert any("새로 만들었습니다" in n for n in again.notes)


def test_problems(filled, tmp_path):
    aje = _aje(tmp_path, filled, [(1, None, "수정", "외상매출금", 5000, None),
                                  (1, None, "수정", "이상한계정", None, 4000)])
    res = adjust.apply(filled, aje, tmp_path / "x.xlsx")
    assert any("이상한계정" in p and "새 계정 묶음" in p for p in res.problems)
    assert any("차변 5,000원 ≠ 대변 4,000원" in p for p in res.problems)
    with pytest.raises(ValueError, match="덮어쓰지"):
        adjust.apply(filled, aje, filled)


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice 없음 (수식 계산 확인용)")
def test_formulas_balance_after_recalc(filled, tmp_path):
    aje = _aje(tmp_path, filled, AJE_ROWS)
    out = adjust.apply(filled, aje, tmp_path / "반영.xlsx").output
    prof = tmp_path / "lo"
    prof.mkdir()
    (prof / "user").mkdir()
    (prof / "user" / "registrymodifications.xcu").write_text(
        '<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry" '
        'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse">'
        '<value>0</value></prop></item></oor:items>', encoding="utf-8")
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--convert-to", "xlsx",
                    "--outdir", str(tmp_path / "calc"), str(out)], check=True, capture_output=True, timeout=180)
    wb = openpyxl.load_workbook(tmp_path / "calc" / out.name, data_only=True)

    def find(sheet, label, col=2, label_col=1):
        ws = wb[sheet]
        r = next(r for r in range(1, ws.max_row + 1) if (ws.cell(r, label_col).value or "").strip() == label)
        return ws.cell(r, col).value
    assert find("수정후IS", "Ⅹ. 당기순이익(손실)") == 190000
    assert find("수정후BS", "자산총계") == 1440000 and find("수정후BS", "부채와자본총계") == 1440000
    assert find("현금흐름정산표", "현금의 증가(감소)") == 100000
    assert find("현금흐름정산표", "검증: 기말 현금(수정후BS) - 기말의 현금 (0이어야 함)") == 0
    ws = wb["정산표"]
    checks = [ws.cell(r, 4).value for r in range(1, ws.max_row + 1)
              if (ws.cell(r, 3).value or "").endswith(("차변-대변", "(TB에 없는 과목)"))]
    assert checks == [0, 0, 0, 0, 0]
    assert wb["C1000"]["I13"].value == -60000                          # 대손충당금 Audited = G/L -50,000 - 10,000


def test_cli_and_web(filled, tmp_path, capsys, server, store):  # noqa: F811
    t = tmp_path / "cli양식.xlsx"
    assert main(["adjust", str(filled), "--template", "-o", str(t)]) == 0
    aje = _aje(tmp_path, filled, AJE_ROWS)
    assert main(["adjust", str(filled), str(aje), "-o", str(tmp_path / "cli.xlsx")]) == 0
    assert "정산표에 반영" in capsys.readouterr().out

    base, _ = server
    cid = store.add_client("C012", "가나상사", "김담당", "kim@example.com")
    eid = store.add_engagement(cid, 2026, "기말감사", "2027-01-11", "2027-01-15")
    st, h, _ = _open(base + f"/e/{eid}/aje-template.xlsx")
    assert st == 400                                                    # 조서를 먼저 채워야 한다
    store.set_setting(f"leadsheet:{eid}", {"made_at": "x", "path": str(filled), "lines": []})
    st, h, x = _open(base + f"/e/{eid}/aje-template.xlsx")
    assert st == 200 and x[:2] == b"PK"
    data, hdr = _multipart([], [("aje", "수정분개.xlsx", aje.read_bytes())])
    st, h, _ = _open(base + f"/e/{eid}/adjust", data, hdr)
    assert st == 303 and "#adjust" in h["Location"]
    rec = store.get_setting(f"adjust:{eid}")
    assert rec["path"].endswith("조서_리드시트_수정분개반영.xlsx")
    st, _, x = _open(base + f"/e/{eid}/adjusted.xlsx")
    assert st == 200 and "정산표" in openpyxl.load_workbook(io.BytesIO(x)).sheetnames
    page = _open(base + f"/e/{eid}")[2].decode()
    assert "수정분개 양식 내려받기" in page and "수정분개반영.xlsx" in page


def test_web_swapped_or_wrong_uploads(filled, tmp_path, server, store):  # noqa: F811
    from auditautopilot.manage.web import xlsx_sheet_names
    aje = _aje(tmp_path, filled, AJE_ROWS)
    assert "TB" in xlsx_sheet_names(filled.read_bytes()) and xlsx_sheet_names(b"not zip") == []
    base, _ = server
    cid = store.add_client("C013", "다라상사", "박담당", "park@example.com")
    eid = store.add_engagement(cid, 2026, "기말감사", "2027-01-11", "2027-01-15")
    store.set_setting(f"leadsheet:{eid}", {"made_at": "x", "path": str(filled), "lines": []})
    # 수정분개 양식을 조서 칸에, 조서를 수정분개 칸에 바꿔 올려도 반영된다
    data, hdr = _multipart([], [("aje", "조서.xlsx", filled.read_bytes()),
                                ("workpaper", "그린_수정분개양식.xlsx", aje.read_bytes())])
    st, h, _ = _open(base + f"/e/{eid}/adjust", data, hdr)
    assert st == 303 and store.get_setting(f"adjust:{eid}")["path"].endswith("조서_수정분개반영.xlsx")
    # 수정분개 양식만 조서 칸에 올려도 마지막 조서로 반영한다
    store.set_setting(f"adjust:{eid}", None)
    data, hdr = _multipart([], [("workpaper", "그린_수정분개양식.xlsx", aje.read_bytes())])
    st, h, _ = _open(base + f"/e/{eid}/adjust", data, hdr)
    assert st == 303 and store.get_setting(f"adjust:{eid}")["path"].endswith("조서_리드시트_수정분개반영.xlsx")
    # 조서만 올리면 수정분개 양식을 고르라고 알려 준다
    data, hdr = _multipart([], [("aje", "조서.xlsx", filled.read_bytes())])
    st, h, _ = _open(base + f"/e/{eid}/adjust", data, hdr)
    assert st == 303 and "수정분개 양식을 골라" in urllib.parse.unquote(h["Location"])


def test_open_workbook_repairs_bad_style_index(filled, tmp_path):
    """서식 번호가 서식 목록보다 큰 파일도 열린다 (다른 프로그램이 저장한 파일)."""
    import re
    import zipfile
    from auditautopilot.workpaper.xlsx_xml import open_workbook
    bad = tmp_path / "bad.xlsx"
    with zipfile.ZipFile(filled) as src, zipfile.ZipFile(bad, "w") as dst:
        for i in src.infolist():
            b = src.read(i.filename)
            if i.filename == "xl/worksheets/sheet3.xml":
                b = re.sub(rb'<c r="([A-Z]+\d+)"( s="\d+")?', rb'<c r="\1" s="999"', b, count=1)
            dst.writestr(i, b)
    with pytest.raises(IndexError):
        openpyxl.load_workbook(bad)
    assert "TB" in open_workbook(bad).sheetnames
    res = adjust.apply(bad, _aje(tmp_path, filled, AJE_ROWS), tmp_path / "반영.xlsx")
    assert res.entries == 1
