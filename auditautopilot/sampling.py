"""표본추출.

모든 방법은 시드값을 기록해 같은 표본을 다시 만들 수 있게 한다 (감사기준서 530의 문서화 대비).
모집단은 dict 목록이고, 결과 행에는 _row(모집단 순번), _abs(금액 절대값), _reason(추출 사유)이 붙는다.
"""

from __future__ import annotations

import bisect
import random
import secrets
from dataclasses import dataclass, field
from itertools import accumulate

METHODS = ("key", "random", "systematic", "mus")


@dataclass
class SampleResult:
    method: str
    seed: int | None
    population_count: int
    population_amount: float
    selected: list[dict]
    params: dict = field(default_factory=dict)

    def summary(self) -> list[dict]:
        items = [
            ("추출 방법", self.method),
            ("시드값", self.seed if self.seed is not None else "해당 없음"),
            ("모집단 건수", self.population_count),
            ("모집단 금액(절대값 합계)", self.population_amount),
            ("표본 건수", len(self.selected)),
            ("표본 금액(절대값 합계)", sum(r["_abs"] for r in self.selected)),
        ]
        items += [(f"설정: {k}", v) for k, v in self.params.items()]
        return [{"항목": k, "값": v} for k, v in items]


def _prepare(rows: list[dict], amount_col: str) -> list[dict]:
    out = []
    for i, r in enumerate(rows, start=1):
        r = dict(r)
        r.setdefault("_row", i)
        r["_abs"] = abs(r[amount_col])
        out.append(r)
    return out


def _seed(seed: int | None) -> int:
    return seed if seed is not None else secrets.randbelow(2**31)


def _total(pop: list[dict]) -> float:
    return float(sum(r["_abs"] for r in pop))


def _tag(rows: list[dict], reason: str) -> list[dict]:
    return [{**r, "_reason": reason} for r in rows]


def key_items(rows: list[dict], amount_col: str, threshold: float) -> SampleResult:
    """금액(절대값)이 기준 이상인 항목 전수."""
    pop = _prepare(rows, amount_col)
    sel = _tag([r for r in pop if r["_abs"] >= threshold], "기준금액 이상 전수")
    return SampleResult("key", None, len(pop), _total(pop), sel, {"기준금액": threshold})


def random_sample(rows: list[dict], amount_col: str, n: int, seed: int | None = None) -> SampleResult:
    pop = _prepare(rows, amount_col)
    s = _seed(seed)
    n = min(n, len(pop))
    idx = sorted(random.Random(s).sample(range(len(pop)), n))
    return SampleResult("random", s, len(pop), _total(pop),
                        _tag([pop[i] for i in idx], "random"), {"표본수": n})


def systematic_sample(rows: list[dict], amount_col: str, n: int, seed: int | None = None) -> SampleResult:
    pop = _prepare(rows, amount_col)
    s = _seed(seed)
    n = min(n, len(pop))
    if n == 0:
        return SampleResult("systematic", s, len(pop), _total(pop), [], {"표본수": 0})
    step = len(pop) / n
    start = random.Random(s).uniform(0, step)
    idx = sorted({int(start + k * step) for k in range(n)})
    return SampleResult("systematic", s, len(pop), _total(pop),
                        _tag([pop[i] for i in idx], "systematic"),
                        {"표본수": n, "간격(건)": round(step, 2)})


def mus_sample(rows: list[dict], amount_col: str, n: int, seed: int | None = None) -> SampleResult:
    """금액단위표본추출(MUS). 표본 간격 이상인 항목은 반드시 뽑힌다."""
    pop = _prepare(rows, amount_col)
    s = _seed(seed)
    total = _total(pop)
    if n <= 0 or total == 0:
        return SampleResult("mus", s, len(pop), total, [], {"표본수": n})
    interval = total / n
    start = random.Random(s).uniform(0, interval)
    cum = list(accumulate(r["_abs"] for r in pop))
    idx = sorted({bisect.bisect_right(cum, start + interval * k) for k in range(n)})
    idx = [i for i in idx if i < len(pop)]
    return SampleResult("mus", s, len(pop), total, _tag([pop[i] for i in idx], "mus"),
                        {"표본수(선택점)": n, "표본간격": round(interval, 2), "시작점": round(start, 2)})


def sample(rows: list[dict], amount_col: str, method: str, n: int = 0,
           threshold: float | None = None, seed: int | None = None,
           exclude_zero: bool = True) -> SampleResult:
    """기준금액 이상 전수 + 나머지에서 선택한 방법으로 추출한다.

    threshold 를 주면 그 이상은 전수로 뽑고, 나머지 모집단에서 method 로 n건을 뽑는다.
    """
    if method not in METHODS:
        raise ValueError(f"지원하는 방법: {', '.join(METHODS)}")
    pop = [r for r in rows if r[amount_col]] if exclude_zero else list(rows)
    pop = _prepare(pop, amount_col)
    if method == "key":
        if threshold is None:
            raise ValueError("key 방법은 기준금액(threshold)이 필요합니다.")
        return key_items(pop, amount_col, threshold)
    fn = {"random": random_sample, "systematic": systematic_sample, "mus": mus_sample}[method]
    if threshold is None:
        return fn(pop, amount_col, n, seed)
    key = key_items(pop, amount_col, threshold)
    rest = [r for r in pop if r["_abs"] < threshold]
    other = fn(rest, amount_col, n, seed)
    params = {"기준금액": threshold, **other.params}
    return SampleResult(f"key+{method}", other.seed, len(pop), _total(pop),
                        key.selected + other.selected, params)
