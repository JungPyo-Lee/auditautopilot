from datetime import date

import pytest

from auditautopilot import analytics, materiality, sampling


def test_materiality_basic_and_rounding():
    r = materiality.compute("매출액", 10_000_000_000, overall_pct=1, performance_pct=75,
                            trivial_pct=5, rounding_unit=1000)
    assert r.overall == 100_000_000
    assert r.performance == 75_000_000
    assert r.trivial == 5_000_000


def test_materiality_annualizes_interim():
    r = materiality.compute("매출액", 9_000_000, overall_pct=10, months=9, rounding_unit=1)
    assert r.annualized_amount == 12_000_000
    assert r.overall == 1_200_000


@pytest.mark.parametrize("kwargs", [{"overall_pct": 0}, {"overall_pct": 5, "months": 13}])
def test_materiality_rejects_bad_input(kwargs):
    with pytest.raises(ValueError):
        materiality.compute("매출액", 1000, **kwargs)


def test_find_benchmark_amount():
    fs = [{"name": "Ⅰ.매출액", "current": 100.0, "prior": 0},
          {"name": "Ⅷ.법인세차감전이익", "current": 10.0, "prior": 0}]
    assert materiality.find_benchmark_amount(fs, "세전이익") == 10.0
    assert materiality.find_benchmark_amount(fs, "자산총계") is None


def _fs():
    rows = [("103", "보통예금", 1_000.0, 950.0), ("108", "외상매출금", 5_000.0, 2_000.0),
            ("120", "미수금", 900.0, 0.0), ("404", "제품매출", 9_000.0, 12_000.0)]
    return [{"code": c, "name": n, "is_account": True, "current": cur, "prior": pri}
            for c, n, cur, pri in rows]


def test_compare_periods_flags():
    v = analytics.compare_periods(_fs(), analytics.VarianceRule(threshold_amount=800, threshold_pct=10))
    assert {r["code"]: r["flag"] for r in v} == {"103": False, "108": True, "120": True, "404": True}
    by = {r["code"]: r for r in v}
    assert by["120"]["reason"].startswith("전기 잔액 없음")
    assert by["120"]["pct"] is None
    assert by["108"]["pct"] == pytest.approx(150)


def test_compare_periods_skips_subtotals():
    fs = _fs() + [{"code": None, "name": "Ⅰ.유동자산", "is_account": False, "current": 1e9, "prior": 0}]
    v = analytics.compare_periods(fs, analytics.VarianceRule(800))
    assert len(v) == 4


def test_compare_periods_annualizes():
    rule = analytics.VarianceRule(threshold_amount=800, threshold_pct=10, annualize_factor=12 / 9)
    v = {r["code"]: r for r in analytics.compare_periods(_fs(), rule, annualize=True)}
    assert v["404"]["current_compared"] == pytest.approx(12_000)
    assert not v["404"]["flag"]


def test_draft_questions_sorted_by_size():
    v = analytics.compare_periods(_fs(), analytics.VarianceRule(800, 10))
    q = analytics.draft_questions(v)
    assert [r["계정과목"] for r in q] == ["외상매출금", "제품매출", "미수금"]
    assert q[0]["요청번호"] == "질문-01"
    assert "변동 사유" in q[0]["질문"]
    assert "새로 발생" in q[2]["질문"]


def test_monthly_trend():
    j = [
        {"date": date(2026, 1, 5), "account_code": "108", "account_name": "외상매출금", "debit": 100.0, "credit": 0.0},
        {"date": date(2026, 1, 20), "account_code": "108", "account_name": "외상매출금", "debit": 50.0, "credit": 0.0},
        {"date": date(2026, 2, 1), "account_code": "108", "account_name": "외상매출금", "debit": 0.0, "credit": 30.0},
        {"date": None, "account_code": "108", "account_name": "외상매출금", "debit": 999.0, "credit": 0.0},
    ]
    t = analytics.monthly_trend(j)
    assert t == [{"account_code": "108", "account_name": "외상매출금", "2026-01": 150.0, "2026-02": -30.0}]


def _population():
    return [{"id": i, "amount": a} for i, a in enumerate([5, 10, 1000, 20, 30, -40, 50, 0, 60, 70])]


def _ids(result):
    return [r["id"] for r in result.selected]


def test_random_sample_is_reproducible():
    a = sampling.sample(_population(), "amount", "random", n=4, seed=42)
    b = sampling.sample(_population(), "amount", "random", n=4, seed=42)
    assert _ids(a) == _ids(b)
    assert len(a.selected) == 4
    assert 7 not in _ids(a)  # 0원 항목은 제외
    assert a.population_count == 9


def test_mus_always_picks_items_above_interval():
    r = sampling.sample(_population(), "amount", "mus", n=5, seed=1)
    interval = r.params["표본간격"]
    big = {p["id"] for p in _population() if abs(p["amount"]) >= interval}
    assert big and big <= set(_ids(r))


def test_key_plus_random():
    r = sampling.sample(_population(), "amount", "random", n=2, threshold=60, seed=3)
    by_reason: dict[str, set] = {}
    for row in r.selected:
        by_reason.setdefault(row["_reason"], set()).add(row["id"])
    assert by_reason["기준금액 이상 전수"] == {2, 8, 9}
    assert len(by_reason["random"]) == 2
    assert not by_reason["random"] & {2, 8, 9}


def test_systematic_sample_count():
    r = sampling.sample(_population(), "amount", "systematic", n=3, seed=5)
    assert len(r.selected) == 3
    assert "간격(건)" in r.params


def test_summary_rows():
    r = sampling.sample(_population(), "amount", "mus", n=3, seed=2)
    items = {row["항목"]: row["값"] for row in r.summary()}
    assert items["시드값"] == 2
    assert items["모집단 건수"] == 9
