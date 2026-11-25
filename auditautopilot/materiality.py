"""중요성 금액 계산.

기준(benchmark)과 비율은 감사인이 판단해 정한다. 여기 있는 기본값은 화면에 처음 채워 넣는
예시일 뿐이며, 회계법인 내부 정책 값으로 바꿔 쓴다.
"""

from __future__ import annotations

import math
from dataclasses import asdict, dataclass

from .importers.common import squash

# 화면 초기값 예시 (퍼센트). 법인 정책에 맞게 바꾼다.
DEFAULT_PERCENTAGES = {
    "세전이익": 5.0,
    "매출액": 1.0,
    "자산총계": 1.0,
    "자본총계": 2.0,
}
DEFAULT_PERFORMANCE_PCT = 75.0
DEFAULT_TRIVIAL_PCT = 5.0

# 재무제표 과목명에서 기준 금액을 찾을 때 쓰는 키워드 (자간 공백 제거 후 비교)
_BENCHMARK_KEYWORDS = {
    "세전이익": ["법인세차감전", "법인세비용차감전"],
    "매출액": ["매출액"],
    "자산총계": ["자산총계"],
    "자본총계": ["자본총계"],
}


@dataclass
class MaterialityResult:
    benchmark: str
    benchmark_amount: float
    annualized_amount: float
    months: int
    overall_pct: float
    overall: float
    performance_pct: float
    performance: float
    trivial_pct: float
    trivial: float
    rounding_unit: int

    def as_rows(self) -> list[dict]:
        items = [
            ("기준", self.benchmark),
            ("기준 금액", self.benchmark_amount),
            ("기준 기간(개월)", self.months),
            ("연환산 기준 금액", self.annualized_amount),
            ("전체중요성 비율(%)", self.overall_pct),
            ("전체중요성", self.overall),
            ("수행중요성 비율(%)", self.performance_pct),
            ("수행중요성", self.performance),
            ("명백히 사소한 금액 비율(%)", self.trivial_pct),
            ("명백히 사소한 금액", self.trivial),
            ("절사 단위", self.rounding_unit),
        ]
        return [{"항목": k, "값": v} for k, v in items]

    def to_dict(self) -> dict:
        return asdict(self)


def _round_down(amount: float, unit: int) -> float:
    if unit <= 1:
        return float(math.floor(amount))
    return float(math.floor(amount / unit) * unit)


def compute(
    benchmark: str,
    benchmark_amount: float,
    overall_pct: float,
    performance_pct: float = DEFAULT_PERFORMANCE_PCT,
    trivial_pct: float = DEFAULT_TRIVIAL_PCT,
    months: int = 12,
    rounding_unit: int = 1000,
) -> MaterialityResult:
    """중요성 금액을 계산한다.

    - 비율은 퍼센트 숫자로 받는다 (5 = 5%).
    - 수행중요성과 명백히 사소한 금액은 전체중요성에 대한 비율이다.
    - months < 12 이면 기준 금액을 12개월로 연환산한다 (중간감사용).
    - 금액은 rounding_unit 단위로 내림한다.
    """
    if not 0 < overall_pct <= 100:
        raise ValueError("전체중요성 비율은 0보다 크고 100 이하여야 합니다.")
    if not 0 < performance_pct <= 100:
        raise ValueError("수행중요성 비율은 0보다 크고 100 이하여야 합니다.")
    if not 0 <= trivial_pct <= 100:
        raise ValueError("명백히 사소한 금액 비율은 0 이상 100 이하여야 합니다.")
    if not 1 <= months <= 12:
        raise ValueError("기준 기간은 1~12개월이어야 합니다.")

    annualized = abs(benchmark_amount) * 12 / months
    overall = _round_down(annualized * overall_pct / 100, rounding_unit)
    performance = _round_down(overall * performance_pct / 100, rounding_unit)
    trivial = _round_down(overall * trivial_pct / 100, rounding_unit)
    return MaterialityResult(
        benchmark=benchmark,
        benchmark_amount=float(benchmark_amount),
        annualized_amount=annualized,
        months=months,
        overall_pct=overall_pct,
        overall=overall,
        performance_pct=performance_pct,
        performance=performance,
        trivial_pct=trivial_pct,
        trivial=trivial,
        rounding_unit=rounding_unit,
    )


def find_benchmark_amount(fs: list[dict], benchmark: str, column: str = "current") -> float | None:
    """재무제표 행(name, current, prior)에서 기준 금액을 찾는다. 없으면 None."""
    keywords = _BENCHMARK_KEYWORDS.get(benchmark, [benchmark])
    for row in fs:
        name = squash(row["name"])
        if any(k in name for k in keywords):
            return float(row[column])
    return None
