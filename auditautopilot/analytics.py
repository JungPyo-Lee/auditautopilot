"""분석적 검토: 전기 대비 증감 분석, 월별 추이, 질문서 초안."""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd


@dataclass
class VarianceRule:
    threshold_amount: float          # 이 금액 이상 변동한 계정만 표시 (보통 수행중요성)
    threshold_pct: float = 10.0      # 그리고 이 비율(%) 이상 변동한 계정
    annualize_factor: float = 1.0    # 손익 과목을 연환산할 배수 (예: 9개월 -> 12/9)


def compare_periods(fs: pd.DataFrame, rule: VarianceRule, accounts_only: bool = True,
                    annualize: bool = False) -> pd.DataFrame:
    """재무제표(code, name, current, prior)의 전기 대비 증감을 계산하고 기준을 넘는 계정을 표시한다.

    annualize=True 이면 당기 금액에 rule.annualize_factor 를 곱해 비교한다 (중간감사 손익).
    """
    df = fs.copy()
    if accounts_only and "is_account" in df.columns:
        df = df[df["is_account"].astype(bool)]
    factor = rule.annualize_factor if annualize else 1.0
    df["current_compared"] = df["current"] * factor
    df["diff"] = df["current_compared"] - df["prior"]
    with np.errstate(divide="ignore", invalid="ignore"):
        df["pct"] = np.where(df["prior"] != 0, df["diff"] / df["prior"].abs() * 100, np.nan)
    big = df["diff"].abs() >= rule.threshold_amount
    pct_ok = df["pct"].isna() | (df["pct"].abs() >= rule.threshold_pct)
    df["flag"] = big & pct_ok
    df["reason"] = np.where(
        ~df["flag"], "",
        np.where(df["prior"] == 0, "전기 잔액 없음, 당기 신규 발생", "증감액·증감률 기준 초과"),
    )
    return df.reset_index(drop=True)


def _fmt_amount(x: float) -> str:
    return f"{x:+,.0f}원"


def draft_questions(variance: pd.DataFrame, prefix: str = "질문") -> pd.DataFrame:
    """표시된 계정마다 회사에 보낼 질문 초안을 만든다. 회계사가 화면에서 고친 뒤 승인한다."""
    flagged = variance[variance["flag"]].copy()
    flagged = flagged.reindex(flagged["diff"].abs().sort_values(ascending=False).index)
    rows = []
    for i, (_, r) in enumerate(flagged.iterrows(), start=1):
        if pd.isna(r["pct"]):
            change = f"{_fmt_amount(r['diff'])} 새로 발생했습니다"
        else:
            change = f"{_fmt_amount(r['diff'])}({r['pct']:+.1f}%) 변동했습니다"
        rows.append({
            "요청번호": f"{prefix}-{i:02d}",
            "계정코드": r.get("code"),
            "계정과목": r["name"],
            "전기": r["prior"],
            "당기": r["current"],
            "증감": r["diff"],
            "질문": f"{r['name']} 금액이 전기 대비 {change}. 변동 사유와 근거 자료를 알려주세요.",
        })
    return pd.DataFrame(rows, columns=["요청번호", "계정코드", "계정과목", "전기", "당기", "증감", "질문"])


def monthly_trend(journal: pd.DataFrame, account_codes: list[str] | None = None) -> pd.DataFrame:
    """분개장에서 계정별 월별 순발생액(차변-대변)을 만든다."""
    df = journal.dropna(subset=["date"]).copy()
    if account_codes:
        df = df[df["account_code"].isin(account_codes)]
    df["month"] = df["date"].dt.to_period("M").astype(str)
    df["net"] = df["debit"] - df["credit"]
    pivot = df.pivot_table(index=["account_code", "account_name"], columns="month",
                           values="net", aggfunc="sum", fill_value=0)
    return pivot.reset_index()
