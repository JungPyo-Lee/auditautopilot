"""장부로 질문서 초안 만들기. 가짜 회사·거래처·숫자만 쓴다."""

import io
import urllib.parse
from datetime import date

import openpyxl
import pytest

from auditautopilot.importers.books import Books
from auditautopilot.manage import auto_questions as aq
from auditautopilot.manage import questions
from test_douzone import dz  # noqa: F401
from test_manage import _multipart, _open, server, store  # noqa: F401

PERIOD = "제 3(당)기[2026/01/01 ~ 2026/12/31]"
PRIOR = "제 2(전)기[2025/01/01 ~ 2025/12/31]"


def _acc(name, cur, pri, is_account=True):
    return {"code": None, "name": name, "is_account": is_account, "current": cur, "prior": pri}


def _line(d, no, acc, dr=0.0, cr=0.0, partner="", memo="", side=""):
    return {"date": d, "voucher_no": no, "side": side or ("차변" if dr else "대변"), "account_code": None,
            "account_name": acc, "debit": dr, "credit": cr, "partner_code": None, "partner": partner, "memo": memo}


@pytest.fixture
def bk():
    bs = [_acc("Ⅰ.유동자산", 0, 0, False), _acc("(1)당좌자산", 0, 0, False),
          _acc("보통예금", 5_000_000, 1_000_000), _acc("외상매출금", 3_000_000, 3_000_000),
          _acc("대손충당금(외상매출금)", 30_000, 30_000), _acc("받을어음", 200_000, 900_000),
          _acc("(2)재고자산", 0, 0, False), _acc("상품", 2_000_000, 500_000),
          _acc("Ⅰ.유동부채", 0, 0, False), _acc("외상매입금", 1_100_000, 1_000_000)]
    pl = [_acc("Ⅰ.매출액", 10_000_000, 8_000_000, False), _acc("상품매출", 10_000_000, 8_000_000),
          _acc("Ⅱ.매출원가", 7_000_000, 5_000_000, False), _acc("상품매출원가", 7_000_000, 5_000_000, False),
          _acc("기초상품재고액", 500_000, 400_000), _acc("기말상품재고액", 2_000_000, 500_000),
          _acc("Ⅳ.판매비와관리비", 0, 0, False), _acc("급여", 900_000, 850_000)]
    partners = [
        {"account_name": "외상매출금", "partner": "가나상사", "opening": 400_000, "debit": 0, "credit": 0, "balance": 400_000},
        {"account_name": "외상매출금", "partner": "다라상사", "opening": 100_000, "debit": 50_000, "credit": 150_000,
         "balance": 0},
        {"account_name": "외상매출금", "partner": "마바상사", "opening": 0, "debit": 2_600_000, "credit": 0,
         "balance": 2_600_000}]
    j = []
    for i in range(1, 13):           # 평범한 매출 12건 (평일)
        d = date(2026, i, 10 if date(2026, i, 10).weekday() < 5 else 12)
        j += [_line(d, "00001", "외상매출금", dr=100_000 * i, partner="가나상사"),
              _line(d, "00001", "상품매출", cr=100_000 * i, partner="가나상사", memo="상품 매출")]
        j += [_line(d, "00002", "상품", dr=50_000 * i, partner="사아물산"),
              _line(d, "00002", "외상매입금", cr=50_000 * i, partner="사아물산", memo="상품 매입")]
    sat = date(2026, 3, 14)          # 토요일, 매출 상대계정이 가수금
    j += [_line(sat, "00003", "가수금", dr=2_000_000, memo="대표이사 가수금 매출 대체"),
          _line(sat, "00003", "상품매출", cr=2_000_000)]
    j += [_line(date(2026, 5, 6), "00004", "상품매출", dr=300_000, memo="매출 취소"),
          _line(date(2026, 5, 6), "00004", "외상매출금", cr=300_000)]
    j += [_line(date(2026, 6, 3), "00005", "보통예금", dr=1_500_000, memo="이자"),
          _line(date(2026, 6, 3), "00005", "이자수익", cr=1_500_000)]      # 금액만 큰 평범한 거래
    return Books(bs=bs, pl=pl, journal=j, partners=partners, periods={"current": PERIOD, "prior": PRIOR})


def test_rules(bk):
    plan = aq.build(bk, 1_000_000, seed=7)
    keys = {d.key for d in plan.drafts}
    # 1. 증감: 보통예금 +4,000,000, 상품 +1,500,000, 매출 +2,000,000 / 외상매입금 +100,000 은 아님
    assert {"증감:BS:보통예금", "증감:BS:상품", "증감:PL:상품매출"} <= keys
    assert "증감:PL:기말상품재고액" not in keys        # 재고는 재무상태표와 원가율 규칙에서 본다
    assert "증감:BS:외상매입금" not in keys and "증감:PL:급여" not in keys
    # 2. 대손: 외상매출금 전기=당기, 가나상사 거래 없이 그대로, 대손충당금 전기와 같음
    assert {"대손:계정:외상매출금", "대손:거래처:외상매출금", "대손:충당금:외상매출금"} <= keys
    ar = next(d for d in plan.drafts if d.key == "대손:거래처:외상매출금")
    assert "가나상사" in ar.text and "다라상사" not in ar.text and ar.amount == 400_000
    # 3. 원가율: 62.5% -> 70.0%
    cr = next(d for d in plan.drafts if d.key == "원가율")
    assert "62.5%" in cr.text and "70.0%" in cr.text and "+7.5%p" in cr.text
    # 4. 특이분개: 토요일 가수금 매출, 매출 차변 / 금액만 큰 이자 입금은 고르지 않는다
    je = [d for d in plan.drafts if d.rule == "특이분개"]
    assert [d.key for d in je] == ["JE:2026-03-14:00003", "JE:2026-05-06:00004"]
    assert "주말 전표" in je[0].text and "매출 상대계정이 가수금" in je[0].text
    assert "매출 계정 차변" in je[1].text
    # 5. 표본: 매출·매입 5건씩, 시드 기록, 같은 시드면 같은 표본
    s = [d for d in plan.drafts if d.rule.startswith("표본")]
    assert sum(d.rule == "표본(매출)" for d in s) == 5 and sum(d.rule == "표본(매입)" for d in s) == 5
    assert plan.seed == 7 and all("시드 7" in d.text for d in s)
    assert [d.key for d in s] == [d.key for d in aq.build(bk, 1_000_000, seed=7).drafts if d.rule.startswith("표본")]
    assert not any("12개월" in n for n in plan.notes)


def test_interim_annualize_and_pre_closing(bk):
    bk.periods["current"] = "제 3(당)기[2026/01/01 ~ 2026/09/30]"
    for r in bk.pl:
        if r["name"] in ("Ⅱ.매출원가", "상품매출원가"):
            r["current"] = 0
    plan = aq.build(bk, 1_000_000, seed=1)
    keys = {d.key for d in plan.drafts}
    sales = next(d for d in plan.drafts if d.key == "증감:PL:상품매출")
    assert "12개월 환산 13,333,333원" in sales.text
    assert "증감:PL:기말상품재고액" not in keys and "증감:BS:상품" not in keys and "원가율" not in keys
    assert any("결산 전" in n for n in plan.notes) and any("9개월" in n for n in plan.notes)


def test_months_and_refusals():
    assert aq.months_of("제 38(당)기[2026/01/01 ~ 2026/10/31]") == 10
    assert aq.months_of("2026년1월1일 ~ 2026년6월30일") == 6
    with pytest.raises(ValueError, match="중요성"):
        aq.build(Books(), 0)
    plan = aq.build(Books(), 1)
    assert not plan.drafts and len(plan.notes) >= 4


def test_apply_dedupes_and_request_sheet(store, bk):  # noqa: F811
    cid = store.add_client("C010", "가나상사", "김담당", "kim@example.com")
    eid = store.add_engagement(cid, 2026, "기말감사", "2027-01-11", "2027-01-15")
    plan = aq.build(bk, 1_000_000, seed=3)
    added, skipped = aq.apply(store, eid, plan)
    assert added == len(plan.drafts) and skipped == 0
    have = {q["source"] for q in questions.all_(store, eid)}
    again = aq.build(bk, 1_000_000, seed=99, have=have)      # 표본은 다시 뽑지 않는다
    assert not [d for d in again.drafts if d.rule.startswith("표본")]
    assert aq.apply(store, eid, again) == (0, len(again.drafts))
    e = store.engagement(eid)
    wb = openpyxl.load_workbook(io.BytesIO(questions.xlsx_bytes(e, questions.all_(store, eid))))
    ws = wb["요청자료"]
    rows = [r for r in ws.iter_rows(min_row=5, values_only=True)]
    assert len(rows) == 10 and rows[0][1] == "매출 표본" and "세금계산서" in rows[0][6] and "가나상사" in rows[0][4]
    assert wb.worksheets[0].title == "질문서"


def test_web_auto_questions(server, store, dz):  # noqa: F811
    base, _ = server
    cid = store.add_client("C011", "가나상사", "김담당", "kim@example.com")
    eid = store.add_engagement(cid, 2026, "기말감사", "2027-01-11", "2027-01-15")
    files = [("files", p.name, p.read_bytes()) for p in (dz["bs"], dz["pl"], dz["ar"], dz["journal"])]
    data, hdr = _multipart([("materiality", ""), ("seed", "")], files)
    st, h, _ = _open(base + f"/e/{eid}/questions/auto", data, hdr)
    assert st == 303 and "중요성" in urllib.parse.unquote(h["Location"])
    data, hdr = _multipart([("materiality", "100,000"), ("seed", "5")], files)
    st, h, _ = _open(base + f"/e/{eid}/questions/auto", data, hdr)
    assert st == 303 and "#questions" in h["Location"]
    srcs = {q["source"] for q in questions.all_(store, eid)}
    assert "증감:PL:제품매출" in srcs and "샘플:매출:2026-01-05:00001" in srcs
    page = _open(base + f"/e/{eid}")[2].decode()
    assert "장부로 증감분석 질문 초안 만들기" in page and "표본 시드 5" in page
