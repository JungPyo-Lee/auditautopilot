"""표본추출.

모든 방법은 시드값을 기록해 같은 표본을 다시 만들 수 있게 한다 (감사기준서 530의 문서화 대비).
"""

from __future__ import annotations

import secrets
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

METHODS = ("key", "random", "systematic", "mus")


@dataclass
class SampleResult:
    method: str
    seed: int | None
    population_count: int
    population_amount: float
    selected: pd.DataFrame
    params: dict = field(default_factory=dict)

    def summary(self) -> pd.DataFrame:
        rows = [
            ("추출 방법", self.method),
            ("시드값", self.seed if self.seed is not None else "해당 없음"),
            ("모집단 건수", self.population_count),
            ("모집단 금액(절대값 합계)", self.population_amount),
            ("표본 건수", len(self.selected)),
            ("표본 금액(절대값 합계)", float(self.selected["_abs"].sum()) if len(self.selected) else 0.0),
        ]
        rows += [(f"설정: {k}", v) for k, v in self.params.items()]
        return pd.DataFrame(rows, columns=["항목", "값"])


def _prepare(df: pd.DataFrame, amount_col: str) -> pd.DataFrame:
    out = df.reset_index(drop=True).copy()
    out["_abs"] = out[amount_col].abs()
    out["_row"] = range(1, len(out) + 1)
    return out


def _seed(seed: int | None) -> int:
    return seed if seed is not None else secrets.randbelow(2**31)


def key_items(df: pd.DataFrame, amount_col: str, threshold: float) -> SampleResult:
    """금액(절대값)이 기준 이상인 항목 전수."""
    pop = _prepare(df, amount_col)
    sel = pop[pop["_abs"] >= threshold]
    return SampleResult("key", None, len(pop), float(pop["_abs"].sum()), sel,
                        {"기준금액": threshold})


def random_sample(df: pd.DataFrame, amount_col: str, n: int, seed: int | None = None) -> SampleResult:
    pop = _prepare(df, amount_col)
    s = _seed(seed)
    n = min(n, len(pop))
    rng = np.random.default_rng(s)
    idx = np.sort(rng.choice(len(pop), size=n, replace=False)) if n else []
    return SampleResult("random", s, len(pop), float(pop["_abs"].sum()), pop.iloc[idx],
                        {"표본수": n})


def systematic_sample(df: pd.DataFrame, amount_col: str, n: int, seed: int | None = None) -> SampleResult:
    pop = _prepare(df, amount_col)
    s = _seed(seed)
    n = min(n, len(pop))
    if n == 0:
        return SampleResult("systematic", s, len(pop), 0.0, pop.iloc[[]], {"표본수": 0})
    step = len(pop) / n
    start = np.random.default_rng(s).uniform(0, step)
    idx = sorted({int(start + k * step) for k in range(n)})
    return SampleResult("systematic", s, len(pop), float(pop["_abs"].sum()), pop.iloc[idx],
                        {"표본수": n, "간격(건)": round(step, 2)})


def mus_sample(df: pd.DataFrame, amount_col: str, n: int, seed: int | None = None) -> SampleResult:
    """금액단위표본추출(MUS). 표본 간격 이상인 항목은 반드시 뽑힌다."""
    pop = _prepare(df, amount_col)
    s = _seed(seed)
    total = float(pop["_abs"].sum())
    if n <= 0 or total == 0:
        return SampleResult("mus", s, len(pop), total, pop.iloc[[]], {"표본수": n})
    interval = total / n
    start = np.random.default_rng(s).uniform(0, interval)
    points = start + interval * np.arange(n)
    cum = pop["_abs"].cumsum().to_numpy()
    idx = np.unique(np.searchsorted(cum, points, side="right"))
    idx = idx[idx < len(pop)]
    return SampleResult("mus", s, len(pop), total, pop.iloc[idx],
                        {"표본수(선택점)": n, "표본간격": round(interval, 2), "시작점": round(float(start), 2)})


def sample(df: pd.DataFrame, amount_col: str, method: str, n: int = 0,
           threshold: float | None = None, seed: int | None = None,
           exclude_zero: bool = True) -> SampleResult:
    """기준금액 이상 전수 + 나머지에서 선택한 방법으로 추출한다.

    threshold 를 주면 그 이상은 전수로 뽑고, 나머지 모집단에서 method 로 n건을 뽑는다.
    """
    if method not in METHODS:
        raise ValueError(f"지원하는 방법: {', '.join(METHODS)}")
    pop = df[df[amount_col] != 0] if exclude_zero else df
    if method == "key":
        if threshold is None:
            raise ValueError("key 방법은 기준금액(threshold)이 필요합니다.")
        return key_items(pop, amount_col, threshold)
    fn = {"random": random_sample, "systematic": systematic_sample, "mus": mus_sample}[method]
    if threshold is None:
        return fn(pop, amount_col, n, seed)
    key = key_items(pop, amount_col, threshold)
    rest = pop.reset_index(drop=True)[pop[amount_col].abs().reset_index(drop=True) < threshold]
    other = fn(rest, amount_col, n, seed)
    key_sel = key.selected.assign(_reason="기준금액 이상 전수")
    other_sel = other.selected.assign(_reason=method)
    selected = pd.concat([key_sel.drop(columns=["_row"]), other_sel.drop(columns=["_row"])],
                         ignore_index=True)
    params = {"기준금액": threshold, **other.params}
    return SampleResult(f"key+{method}", other.seed, len(pop),
                        float(pop[amount_col].abs().sum()), selected, params)
