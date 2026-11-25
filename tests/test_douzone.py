"""더존 재무상태표·손익계산서(과목별)·거래처원장 불러오기. 가짜 숫자만 쓴다."""

import openpyxl
import pytest

from auditautopilot.cli import _months_from_period, main
from auditautopilot.importers import books, douzone

PERIOD = "제 3(당)기[2026/01/01 ~ 2026/10/31]"
PRIOR = "제 2(전)기[2025/01/01 ~ 2025/12/31]"


def _save(tmp_path, name, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    p = tmp_path / name
    wb.save(p)
    return p


@pytest.fixture
def dz(tmp_path):
    bs = _save(tmp_path, "재무상태표과목별_20261125.xlsx", [
        ["과목", PERIOD, None, PRIOR, None],
        [None, "금액", None, "금액", None],
        ["자산", "", "", "", ""],
        ["    Ⅰ. 유동자산", "", "1,450,000", "", "1,300,000"],
        ["        (1)  당좌자산", "", "1,450,000", "", "1,300,000"],
        ["            보통예금", "", "500,000", "", "400,000"],
        ["            외상매출금", "1,000,000", "", "950,000", ""],
        ["            대손충당금", "50,000", "950,000", "50,000", "900,000"],
        ["    Ⅱ. 비유동자산", "", "300,000", "", "400,000"],
        ["            건물", "1,000,000", "", "1,000,000", ""],
        ["            감가상각누계액", "700,000", "300,000", "600,000", "400,000"],
        ["자산총계", "", "1,750,000", "", "1,700,000"],
        ["부채", "", "", "", ""],
        ["    Ⅰ. 유동부채", "", "250,000", "", "200,000"],
        ["            외상매입금", "", "250,000", "", "200,000"],
        ["부채총계", "", "250,000", "", "200,000"],
        ["(당기순이익)", "", "", "", ""],
        ["            당기 :      100,000 원", "", "", "", ""],
    ])
    pl = _save(tmp_path, "손익계산서과목별_20261125.xlsx", [
        ["과 목", "제 3 (당)기  [2026/01/01 ~ 2026/10/31]", None, PRIOR, None],
        [None, "금 액", None, "금 액", None],
        ["Ⅰ. 매출액", "", 5000000, "", 5500000],
        ["           제품매출", 5000000, "", 5500000, ""],
        ["Ⅱ. 매출원가", "", 0, "", 3000000],
        ["           제품매출원가", "", 0, "", 3000000],
        ["           기초 제품 재고액", 100, "", 90, ""],
        ["Ⅷ. 법인세차감전이익", "", 900000, "", 1000000],
    ])
    ar = _save(tmp_path, "거래처원장_20261125 (1).xlsx", [
        ["거래처", "전기(월)이월", "차변", "대변", "잔액"],
        ["가나상사", 500000, 3000000, 2700000, 800000],
        ["다라상사", 450000, 1000000, 1200000, 250000],
        ["마바상사", 0, 0, 50000, -50000],
        ["합계", "950,000", "4,000,000", "3,950,000", "1,000,000"],
    ])
    ap = _save(tmp_path, "거래처원장_20261125 (2).xlsx", [
        ["거래처", "전기(월)이월", "차변", "대변", "잔액"],
        ["사아물산", 200000, 100000, 150000, "250000"],
        ["합계", "200,000", "100,000", "150,000", "250,000"],
    ])
    names_only = _save(tmp_path, "거래처원장_20261125.xlsx", [["거래처"], ["가나상사"], ["다라상사"]])
    journal = _save(tmp_path, "분개장_20261125.xlsx", [
        ["전표일자", "전표번호", "구분", "Code", "계정과목", "차변", "대변", "거래처", "Code", "적요"],
        ["2026-01-05", "00001", "차변", "108", "외상매출금", 300000, 0, "가나상사", "101", "매출"],
        ["2026-01-05", "00001", "대변", "404", "제품매출", 0, 300000, "가나상사", "101", "매출"],
        ["2026-11-02", "00002", "차변", "103", "보통예금", 100, 0, "", "", "기간 뒤"],
        ["2026-11-02", "00002", "대변", "108", "외상매출금", 0, 100, "", "", "기간 뒤"],
        ["합계 (건수 : 2 )", "", "", "", "", "300,100", "300,100", "", "", ""],
    ])
    return {"bs": bs, "pl": pl, "ar": ar, "ap": ap, "names": names_only, "journal": journal}


def test_detect_kind(dz):
    kinds = {k: douzone.detect_kind(douzone._rows_of(p)) for k, p in dz.items()}
    assert kinds == {"bs": "bs", "pl": "pl", "ar": "ledger", "ap": "ledger",
                     "names": "partner_list", "journal": "journal"}


def test_read_bs_contra_and_headers(dz):
    kind, rows, periods = douzone.load_fs(dz["bs"])
    by = {r["name"]: r for r in rows}
    assert kind == "bs" and periods["current"] == PERIOD
    assert by["외상매출금"]["current"] == 1_000_000 and by["외상매출금"]["is_account"]
    assert by["대손충당금(외상매출금)"]["current"] == 50_000
    assert by["감가상각누계액(건물)"]["prior"] == 600_000
    assert by["보통예금"]["current"] == 500_000
    assert not by["Ⅰ.유동자산"]["is_account"] and not by["자산총계"]["is_account"]
    assert "자산" not in by and not any("당기:" in n for n in by)


def test_read_pl_subtotals(dz):
    _, rows, _ = douzone.load_fs(dz["pl"])
    by = {r["name"]: r for r in rows}
    assert by["제품매출"]["is_account"] and by["제품매출"]["current"] == 5_000_000
    assert not by["제품매출원가"]["is_account"]
    assert not by["Ⅰ.매출액"]["is_account"] and by["Ⅰ.매출액"]["current"] == 5_000_000


def test_books_matches_ledgers_and_codes(dz):
    b = books.load(list(dz.values()))
    assert b.source == "더존"
    assert {r["account_name"] for r in b.partners} == {"외상매출금", "외상매입금"}
    ar = books.partners_of(b, "외상매출금")
    assert len(ar) == 3 and ar[0]["account_code"] == "108"
    assert books.partners_of(b, "108") == ar
    assert len(b.journal) == 2  # 11월 전표와 합계 줄은 빠진다
    assert any("이름만" in n for n in b.notes)
    assert any("뒤의 2줄" in n for n in b.notes)
    assert next(r for r in b.bs if r["name"] == "외상매출금")["code"] == "108"


def test_unmatched_ledger_uses_file_name(tmp_path, dz):
    odd = _save(tmp_path, "미수금_20261125.xlsx", [
        ["거래처", "전기(월)이월", "차변", "대변", "잔액"], ["가", 0, 7, 0, 7], ["합계", 0, 7, 0, 7]])
    b = books.load([dz["bs"], odd])
    assert b.partners[0]["account_name"] == "미수금"
    assert any("못 찾아" in n for n in b.notes)


def test_period_helpers():
    assert _months_from_period(PERIOD) == 10
    assert books.period_end(PERIOD).isoformat() == "2026-10-31"
    assert books.period_end("제 14(당)기 2026년1월1일 ~ 2026년9월30일").isoformat() == "2026-09-30"


def test_cli_with_douzone_files(dz, tmp_path, capsys):
    assert main(["materiality", str(dz["bs"]), str(dz["pl"]), "--benchmark", "매출액", "--overall-pct", "1"]) == 0
    out = capsys.readouterr().out
    assert "기준 기간(개월): 10" in out and "연환산 기준 금액: 6,000,000" in out

    an = tmp_path / "an.xlsx"
    assert main(["analytics", str(dz["bs"]), str(dz["pl"]), "-o", str(an), "--threshold-amount", "10000"]) == 0
    assert "질문서 초안" in openpyxl.load_workbook(an).sheetnames

    sp = tmp_path / "s.xlsx"
    assert main(["sample", str(dz["bs"]), str(dz["ar"]), "--account", "외상매출금", "--method", "random",
                 "--n", "2", "--seed", "1", "-o", str(sp)]) == 0
    ws = openpyxl.load_workbook(sp)["모집단"]
    assert ws.max_row == 4

    imp = tmp_path / "imp.xlsx"
    assert main(["import", "douzone", *map(str, dz.values()), "-o", str(imp)]) == 0
    assert set(openpyxl.load_workbook(imp).sheetnames) == {"재무상태표", "손익계산서", "분개장", "거래처별잔액"}
