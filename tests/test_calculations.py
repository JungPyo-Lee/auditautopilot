import pandas as pd
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
    fs = pd.DataFrame({"name": ["Ⅰ.매출액", "Ⅷ.법인세차감전이익"], "current": [100.0, 10.0], "prior": [0, 0]})
    assert materiality.find_benchmark_amount(fs, "세전이익") == 10.0
    assert materiality.find_benchmark_amount(fs, "자산총계") is None


def _fs():
    return pd.DataFrame({
        "code": ["103", "108", "120", "404"],
        "name": ["보통예금", "외상매출금", "미수금", "제품매출"],
        "is_account": [True, True, True, True],
        "current": [1_000.0, 5_000.0, 900.0, 9_000.0],
        "prior": [950.0, 2_000.0, 0.0, 12_000.0],
    })


def test_compare_periods_flags():
    v = analytics.compare_periods(_fs(), analytics.VarianceRule(threshold_amount=800, threshold_pct=10))
    flags = dict(zip(v["code"], v["flag"]))
    assert flags == {"103": False, "108": True, "120": True, "404": True}
    assert v.set_index("code").loc["120", "reason"].startswith("전기 잔액 없음")


def test_compare_periods_annualizes():
    rule = analytics.VarianceRule(threshold_amount=800, threshold_pct=10, annualize_factor=12 / 9)
    v = analytics.compare_periods(_fs(), rule, annualize=True).set_index("code")
    assert v.loc["404", "current_compared"] == pytest.approx(12_000)
    assert not v.loc["404", "flag"]


def test_draft_questions_sorted_by_size():
    v = analytics.compare_periods(_fs(), analytics.VarianceRule(800, 10))
    q = analytics.draft_questions(v)
    assert list(q["계정과목"]) == ["외상매출금", "제품매출", "미수금"]
    assert q.loc[0, "요청번호"] == "질문-01"
    assert "변동 사유" in q.loc[0, "질문"]


def test_monthly_trend():
    j = pd.DataFrame({
        "date": pd.to_datetime(["2026-01-05", "2026-01-20", "2026-02-01"]),
        "account_code": ["108", "108", "108"], "account_name": ["외상매출금"] * 3,
        "debit": [100.0, 50.0, 0.0], "credit": [0.0, 0.0, 30.0],
    })
    t = analytics.monthly_trend(j)
    assert t.loc[0, "2026-01"] == 150 and t.loc[0, "2026-02"] == -30


def _population():
    return pd.DataFrame({"id": range(10), "amount": [5, 10, 1000, 20, 30, -40, 50, 0, 60, 70]})


def test_random_sample_is_reproducible():
    a = sampling.sample(_population(), "amount", "random", n=4, seed=42)
    b = sampling.sample(_population(), "amount", "random", n=4, seed=42)
    assert list(a.selected["id"]) == list(b.selected["id"])
    assert len(a.selected) == 4
    assert 7 not in set(a.selected["id"])  # 0원 항목은 제외


def test_mus_always_picks_items_above_interval():
    r = sampling.sample(_population(), "amount", "mus", n=5, seed=1)
    interval = r.params["표본간격"]
    big = _population()[_population()["amount"].abs() >= interval]["id"]
    assert set(big) <= set(r.selected["id"])


def test_key_plus_random():
    r = sampling.sample(_population(), "amount", "random", n=2, threshold=60, seed=3)
    reasons = r.selected.groupby("_reason")["id"].apply(set).to_dict()
    assert reasons["기준금액 이상 전수"] == {2, 8, 9}
    assert len(reasons["random"]) == 2


def test_systematic_sample_count():
    r = sampling.sample(_population(), "amount", "systematic", n=3, seed=5)
    assert len(r.selected) == 3
    assert "간격(건)" in r.params
