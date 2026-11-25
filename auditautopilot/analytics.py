"""분석적 검토: 전기 대비 증감 분석, 월별 추이, 질문서 초안."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass


@dataclass
class VarianceRule:
    threshold_amount: float          # 이 금액 이상 변동한 계정만 표시 (보통 수행중요성)
    threshold_pct: float = 10.0      # 그리고 이 비율(%) 이상 변동한 계정
    annualize_factor: float = 1.0    # 손익 과목을 연환산할 배수 (예: 9개월 -> 12/9)


def compare_periods(fs: list[dict], rule: VarianceRule, accounts_only: bool = True,
                    annualize: bool = False) -> list[dict]:
    """재무제표 행(code, name, current, prior)의 전기 대비 증감을 계산하고 기준을 넘는 계정을 표시한다.

    annualize=True 이면 당기 금액에 rule.annualize_factor 를 곱해 비교한다 (중간감사 손익).
    결과 행에는 current_compared, diff, pct(전기 0이면 None), flag, reason 이 더해진다.
    """
    factor = rule.annualize_factor if annualize else 1.0
    out = []
    for row in fs:
        if accounts_only and not row.get("is_account", True):
            continue
        r = dict(row)
        r["current_compared"] = r["current"] * factor
        r["diff"] = r["current_compared"] - r["prior"]
        r["pct"] = r["diff"] / abs(r["prior"]) * 100 if r["prior"] else None
        big = abs(r["diff"]) >= rule.threshold_amount
        pct_ok = r["pct"] is None or abs(r["pct"]) >= rule.threshold_pct
        r["flag"] = big and pct_ok
        if not r["flag"]:
            r["reason"] = ""
        elif not r["prior"]:
            r["reason"] = "전기 잔액 없음, 당기 신규 발생"
        else:
            r["reason"] = "증감액·증감률 기준 초과"
        out.append(r)
    return out


def _fmt_amount(x: float) -> str:
    return f"{x:+,.0f}원"


def draft_questions(variance: list[dict], prefix: str = "질문") -> list[dict]:
    """표시된 계정마다 회사에 보낼 질문 초안을 만든다. 회계사가 화면에서 고친 뒤 승인한다."""
    flagged = sorted((r for r in variance if r["flag"]), key=lambda r: abs(r["diff"]), reverse=True)
    rows = []
    for i, r in enumerate(flagged, start=1):
        if r["pct"] is None:
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
    return rows


def monthly_trend(journal: list[dict], account_codes: list[str] | None = None) -> list[dict]:
    """분개장에서 계정별 월별 순발생액(차변-대변)을 만든다.
    결과: 계정마다 {'account_code', 'account_name', 'YYYY-MM': 금액, ...}"""
    wanted = set(account_codes) if account_codes else None
    totals: dict[tuple, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    months: set[str] = set()
    for j in journal:
        if j["date"] is None or (wanted and j["account_code"] not in wanted):
            continue
        month = j["date"].strftime("%Y-%m")
        months.add(month)
        totals[(j["account_code"] or "", j["account_name"])][month] += j["debit"] - j["credit"]
    ordered = sorted(months)
    out = []
    for (code, name) in sorted(totals):
        row = {"account_code": code, "account_name": name}
        row.update({m: totals[(code, name)].get(m, 0.0) for m in ordered})
        out.append(row)
    return out
