import pandas as pd

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
    accounts = b.tb[b.tb["level"] == "account"]
    assert list(accounts["code"]) == ["103", "108", "251", "404"]
    assert accounts.set_index("code").loc["251", "balance"] == -300000
    assert (b.tb["level"] == "total").sum() == 1

    bs = b.bs.set_index("name")
    assert bs.loc["보통예금", "current"] == 1000000      # 재무상태표 계정은 집계금액 열
    pl = b.pl.set_index("name")
    assert pl.loc["제품매출", "prior"] == 9000000        # 손익계산서 계정은 금액 열

    assert len(b.journal) == 4
    assert b.journal["debit"].sum() == b.journal["credit"].sum()
    assert b.journal.loc[0, "account_code"] == "108"
    assert b.journal.loc[0, "partner_code"] == "101"
    assert pd.api.types.is_datetime64_any_dtype(b.journal["date"])

    p = b.partners.set_index("partner")
    assert p.loc["개인고객", "reg_no"] == "900101-1******"
    assert set(b.partners["account_code"]) == {"108"}


def test_douzone_journal(douzone_journal_file):
    j = douzone.load_journal(douzone_journal_file)
    assert list(j["account_code"]) == ["253", "103"]
    assert list(j["partner_code"]) == ["101", "101"]
    assert j["debit"].sum() == j["credit"].sum() == 1000
