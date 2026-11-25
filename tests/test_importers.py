from datetime import date

from auditautopilot.importers import douzone, semusarang
from auditautopilot.importers.common import mask_reg_no, normalize_code, parse_account_label, to_number


def test_parse_account_label():
    assert parse_account_label("  [0103]   보    통    예    금") == ("103", "보통예금")
    assert parse_account_label("   1.유  동  자  산") == (None, "1.유동자산")


def test_normalize_code_and_numbers():
    assert normalize_code("0253") == normalize_code("253") == "253"
    assert to_number("(1,200)") == -1200
    assert to_number(None) == 0


def test_mask_reg_no_hides_resident_numbers_only():
    assert mask_reg_no("900101-1234567") == "900101-1******"
    assert mask_reg_no("123-45-67890") == "123-45-67890"


def test_semusarang_bundle(semusarang_file):
    b = semusarang.load(semusarang_file)
    assert b.periods["current"].startswith("제 2(당)기")
    accounts = [r for r in b.tb if r["level"] == "account"]
    assert [r["code"] for r in accounts] == ["103", "108", "251", "404"]
    assert {r["code"]: r["balance"] for r in accounts}["251"] == -300000
    assert sum(r["level"] == "total" for r in b.tb) == 1

    bs = {r["name"]: r for r in b.bs}
    assert bs["보통예금"]["current"] == 1000000      # 재무상태표 계정은 집계금액 열
    pl = {r["name"]: r for r in b.pl}
    assert pl["제품매출"]["prior"] == 9000000        # 손익계산서 계정은 금액 열

    assert len(b.journal) == 4
    assert sum(r["debit"] for r in b.journal) == sum(r["credit"] for r in b.journal)
    assert b.journal[0]["account_code"] == "108"
    assert b.journal[0]["partner_code"] == "101"
    assert b.journal[0]["date"] == date(2026, 1, 5)

    p = {r["partner"]: r for r in b.partners}
    assert p["개인고객"]["reg_no"] == "900101-1******"
    assert {r["account_code"] for r in b.partners} == {"108"}


def test_douzone_journal(douzone_journal_file):
    j = douzone.load_journal(douzone_journal_file)
    assert [r["account_code"] for r in j] == ["253", "103"]
    assert [r["partner_code"] for r in j] == ["101", "101"]
    assert sum(r["debit"] for r in j) == sum(r["credit"] for r in j) == 1000
    assert j[0]["date"] == date(2026, 3, 2)
