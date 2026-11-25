"""회사 장부로 질문서 초안을 만든다 (분석적 검토 질문 + 특이분개 질문 + 표본 증빙 요청).

규칙 (회계사가 화면에서 고치고 승인한 뒤 보낸다):
1. 증감: 전기 대비 증감액이 중요성 금액 이상인 계정 -> 증감 사유
   (중간감사 손익은 12개월로 환산해 비교)
2. 대손: 매출채권(외상매출금·받을어음) 중 전기와 당기 금액이 같은 계정, 거래 없이 잔액이 그대로인 거래처,
   전기와 같은 금액인 대손충당금 -> 대손충당금 설정 여부
3. 원가율: 매출원가율이 전기 대비 3%p 이상 변동 -> 사유
4. 특이분개: 분개장에서 비경상적인 전표를 점수로 골라 최대 10건 -> 거래 내용과 근거
   (주말·공휴일, 매출 차변, 매출 상대계정이 채권·현금·부가세가 아닌 경우, 결산분개, 수정·취소·가수금 적요,
    마이너스 금액, 중요성 이상 큰 금액, 백만원 단위로 딱 떨어지는 큰 금액, 적요 없는 큰 금액, 기말 직전 큰 금액)
5. 표본: 매출전표·매입전표에서 무작위로 5건씩 (시드값 기록) -> 세금계산서·거래명세서·입출금 증빙 요청

질문마다 source 키를 붙여 같은 질문을 두 번 만들지 않는다.
"""

from __future__ import annotations

import random
import re
import secrets
from dataclasses import dataclass, field
from datetime import date, timedelta

from ..importers.books import Books, period_end
from ..importers.douzone import section_label

COST_RATIO_PP = 3.0        # 매출원가율 변동 기준 (%p)
JE_MAX = 10                # 특이분개 질문 최대 건수
SAMPLE_N = 5               # 매출·매입 표본 건수

RECEIVABLES = ("외상매출금", "받을어음")
SALES_COUNTER = ("외상매출금", "받을어음", "현금", "보통예금", "당좌예금", "제예금", "예수부가세", "부가세예수금",
                 "선수금", "미수금")
PURCHASE_DEBIT = ("상품", "원재료", "부재료", "저장품", "미착품", "외주가공비")
PURCHASE_CREDIT = ("외상매입금", "지급어음")
MEMO_STRONG = ("수정", "정정", "오류", "취소", "역분개", "가지급", "가수금")
MEMO_WEAK = ("조정", "대체", "임시")     # '계정대체'처럼 흔히 쓰는 말이라 점수를 낮게 준다
STALE_IN_CLOSING = ("미처분이익잉여금", "이월이익잉여금", "미처리결손금")

# 양력 고정 공휴일 (설·추석·대체공휴일 등 음력·임시 공휴일은 주말로만 잡힌다)
FIXED_HOLIDAYS = {(1, 1), (3, 1), (5, 5), (6, 6), (8, 15), (10, 3), (10, 9), (12, 25)}


@dataclass
class Draft:
    key: str
    rule: str
    account: str
    amount: float | None
    text: str


@dataclass
class Plan:
    drafts: list[Draft] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    seed: int | None = None

    def count(self, rule: str) -> int:
        return sum(1 for d in self.drafts if d.rule == rule)


def _won(x: float) -> str:
    return f"{x:,.0f}원"


def _period_start(label: str) -> date | None:
    m = re.findall(r"(\d{4})\s*[/.\-년]\s*(\d{1,2})\s*[/.\-월]\s*(\d{1,2})", label or "")
    return date(*(int(x) for x in m[0])) if m else None


def months_of(label: str) -> int | None:
    """'제 38(당)기[2026/01/01 ~ 2026/10/31]' -> 10"""
    start, end = _period_start(label), period_end(label)
    if not start or not end:
        return None
    return (end.year - start.year) * 12 + end.month - start.month + 1


def _headers(pl: list[dict]) -> dict[str, dict]:
    out = {}
    for r in pl:
        if not r.get("is_account"):
            out.setdefault(section_label(r["name"]), r)
    return out


def _sales_cogs(pl: list[dict]) -> tuple[dict | None, dict | None]:
    h = _headers(pl)
    return h.get("매출액") or h.get("매출"), h.get("매출원가")


def pre_closing(pl: list[dict]) -> bool:
    """결산 전 손익계산서: 매출은 있는데 매출원가가 0원 (기말재고를 아직 안 잡음)"""
    sales, cogs = _sales_cogs(pl)
    return bool(sales and sales["current"] and cogs is not None and not cogs["current"])


# ── 1. 증감 ────────────────────────────────────────────
def _sections(rows: list[dict]):
    """(바로 위 구분 줄 이름, 행). '상품매출원가'처럼 안쪽 소계 줄도 구분 줄로 본다."""
    section = ""
    for r in rows:
        if not r.get("is_account"):
            section = section_label(r["name"])
        yield section, r


def variance(books: Books, materiality: float, plan: Plan) -> None:
    months = months_of(books.periods.get("current", ""))
    factor = 12 / months if months and months < 12 else 1.0
    closing = pre_closing(books.pl)
    rows = []
    for section, r in _sections(books.bs):
        if not r.get("is_account"):
            continue
        if closing and (section == "재고자산" or r["name"] in STALE_IN_CLOSING):
            continue  # 결산 전에는 기말재고·당기순이익이 아직 안 잡혀 있다
        rows.append(("BS", r, 1.0))
    for section, r in _sections(books.pl):
        if not r.get("is_account"):
            continue
        if section.endswith("매출원가") and (closing or "재고액" in r["name"]):
            continue  # 기말재고는 결산 때 잡히므로 원가율 규칙에서 본다
        rows.append(("PL", r, factor))
    if not rows:
        plan.notes.append("증감: 재무상태표·손익계산서(전기 금액 포함)가 없어 건너뜀")
        return
    if factor != 1.0:
        plan.notes.append(f"증감: 손익은 당기 {months}개월 금액을 12개월로 환산해 전기와 비교했습니다")
    if closing:
        plan.notes.append("증감: 결산 전(매출원가 0원)이라 매출원가 구성 항목, 재고자산, 미처분이익잉여금은 뺐습니다")
    found = []
    for kind, r, f in rows:
        cur = r["current"] * f
        diff = cur - r["prior"]
        if abs(diff) >= materiality:
            found.append((abs(diff), kind, r, cur, diff, f))
    found.sort(key=lambda x: -x[0])
    for _, kind, r, cur, diff, f in found:
        if r["prior"]:
            change = f"{diff:+,.0f}원({diff / abs(r['prior']) * 100:+.1f}%)"
        else:
            change = f"{diff:+,.0f}원(전기 0원, 당기 신규 발생)"
        basis = (f"당기 {_won(r['current'])}(12개월 환산 {_won(cur)})" if f != 1.0 else f"당기 {_won(r['current'])}")
        verb = "증가" if diff > 0 else "감소"
        plan.drafts.append(Draft(
            f"증감:{kind}:{r['name']}", "증감", r["name"], r["current"],
            f"{r['name']}: 전기 {_won(r['prior'])}, {basis}으로 {change} {verb}했습니다. "
            f"{verb} 사유(거래처·거래 유형·단가 변동 등)를 설명해 주시고 근거 자료를 보내 주세요."))


# ── 2. 매출채권 대손충당금 ──────────────────────────────
def allowance(books: Books, plan: Plan) -> None:
    bs = [r for r in books.bs if r.get("is_account")]
    n0 = len(plan.drafts)
    for r in bs:
        if r["name"] in RECEIVABLES and r["current"] and abs(r["current"] - r["prior"]) < 1:
            plan.drafts.append(Draft(
                f"대손:계정:{r['name']}", "대손", r["name"], r["current"],
                f"{r['name']} 잔액이 전기와 당기 모두 {_won(r['current'])}으로 같습니다. 회수가 멈춘 채권이 있는지, "
                f"대손충당금을 설정했는지(설정했다면 설정 기준과 금액) 알려 주세요."))
    for acc in RECEIVABLES:
        still = [p for p in books.partners if p.get("account_name") == acc and abs(p["balance"]) >= 1
                 and abs(p["balance"] - p["opening"]) < 1 and not p["debit"] and not p["credit"]]
        if not still:
            continue
        still.sort(key=lambda p: -abs(p["balance"]))
        names = ", ".join(f"{p['partner'] or '(거래처 없음)'} {_won(p['balance'])}" for p in still[:10])
        more = f" 외 {len(still) - 10}곳" if len(still) > 10 else ""
        total = sum(p["balance"] for p in still)
        plan.drafts.append(Draft(
            f"대손:거래처:{acc}", "대손", acc, total,
            f"{acc} 중 기초 잔액이 그대로이고 당기 거래가 없는 거래처가 {len(still)}곳({_won(total)})입니다: {names}{more}. "
            f"회수 가능성과 대손충당금 설정 여부(설정 기준·금액)를 알려 주시고, 회수 계획이나 독촉 자료가 있으면 보내 주세요."))
    for r in bs:
        base = re.match(r"대손충당금\((.+)\)$", r["name"])
        if base and base.group(1) in RECEIVABLES and r["current"] and abs(r["current"] - r["prior"]) < 1:
            rec = next((x for x in bs if x["name"] == base.group(1)), None)
            moved = f" ({base.group(1)}는 전기 {_won(rec['prior'])} → 당기 {_won(rec['current'])})" if rec else ""
            plan.drafts.append(Draft(
                f"대손:충당금:{base.group(1)}", "대손", r["name"], r["current"],
                f"{base.group(1)}의 대손충당금이 전기와 같은 {_won(r['current'])}입니다{moved}. "
                f"당기 대손충당금을 다시 평가했는지, 설정 기준(연령분석·개별 평가 등)과 기말 설정 예정 금액을 알려 주세요."))
    if len(plan.drafts) == n0 and not books.bs:
        plan.notes.append("대손: 재무상태표가 없어 건너뜀")


# ── 3. 매출원가율 ─────────────────────────────────────
def cost_ratio(books: Books, plan: Plan, threshold_pp: float = COST_RATIO_PP) -> None:
    sales, cogs = _sales_cogs(books.pl)
    if not sales or cogs is None:
        plan.notes.append("원가율: 손익계산서에서 매출액·매출원가 줄을 못 찾아 건너뜀")
        return
    if pre_closing(books.pl):
        plan.notes.append("원가율: 결산 전이라 당기 매출원가가 0원입니다. 기말재고를 반영한 손익계산서로 다시 만들어 주세요")
        return
    if not sales["current"] or not sales["prior"]:
        plan.notes.append("원가율: 당기 또는 전기 매출액이 0원이라 건너뜀")
        return
    cur = cogs["current"] / sales["current"] * 100
    pri = cogs["prior"] / sales["prior"] * 100
    if abs(cur - pri) < threshold_pp:
        plan.notes.append(f"원가율: 전기 {pri:.1f}% → 당기 {cur:.1f}%로 {threshold_pp:g}%p 미만 변동")
        return
    verb = "상승" if cur > pri else "하락"
    plan.drafts.append(Draft(
        "원가율", "원가율", "매출원가", cogs["current"],
        f"매출원가율이 전기 {pri:.1f}%에서 당기 {cur:.1f}%로 {cur - pri:+.1f}%p {verb}했습니다 "
        f"(매출액 {_won(sales['current'])}, 매출원가 {_won(cogs['current'])}). 판매단가·원재료 단가·제품 구성 변화 등 "
        f"{verb} 사유를 설명해 주시고, 주요 제품별 매출·원가 자료를 보내 주세요."))


# ── 4. 특이분개 (JE 테스트) ─────────────────────────────
def vouchers(journal: list[dict]) -> dict[tuple, list[dict]]:
    """더존 전표번호는 날짜마다 다시 시작하므로 (일자, 번호)로 묶는다."""
    out: dict[tuple, list[dict]] = {}
    for j in journal:
        if j["date"] is not None:
            out.setdefault((j["date"], j["voucher_no"]), []).append(j)
    return out


def _is_sales(name: str) -> bool:
    return name.endswith("매출") and "원가" not in name


def _amount(lines: list[dict]) -> float:
    return max(abs(sum(j["debit"] for j in lines)), abs(sum(j["credit"] for j in lines)))


def je_reasons(lines: list[dict], materiality: float, end: date | None) -> list[tuple[int, str]]:
    d = lines[0]["date"]
    amt = _amount(lines)
    big = amt >= materiality
    out: list[tuple[int, str]] = []
    if d.weekday() >= 5:
        out.append((2, "주말 전표"))
    elif (d.month, d.day) in FIXED_HOLIDAYS:
        out.append((2, "공휴일 전표"))
    names = {j["account_name"] for j in lines}
    if any(_is_sales(j["account_name"]) and j["debit"] for j in lines):
        out.append((3, "매출 계정 차변(매출 취소·감액)"))
    if any(_is_sales(j["account_name"]) and j["credit"] for j in lines):
        others = {n for n in names if not _is_sales(n)}
        odd = sorted(n for n in others if n not in SALES_COUNTER)
        if odd:
            out.append((3, f"매출 상대계정이 {', '.join(odd)}"))
    if any(j["side"].startswith("결") for j in lines):
        out.append((2, "결산분개"))
    memo = " ".join(j["memo"] for j in lines)
    word = next((w for w in MEMO_STRONG if w in memo), None)
    if word:
        out.append((2, f"적요에 '{word}'"))
    else:
        word = next((w for w in MEMO_WEAK if w in memo), None)
        if word:
            out.append((1, f"적요에 '{word}'"))
    if any(j["debit"] < 0 or j["credit"] < 0 for j in lines):
        out.append((1, "마이너스 금액"))
    if big:
        out.append((2, f"중요성 금액({_won(materiality)}) 이상"))
    if amt >= max(materiality / 10, 10_000_000) and amt % 1_000_000 == 0:
        out.append((1, "백만원 단위로 딱 떨어지는 큰 금액"))
    if amt >= materiality / 2 and not memo.strip():
        out.append((1, "적요 없는 큰 금액"))
    if end and big and end - timedelta(days=5) <= d <= end:
        out.append((1, "기간 말 5일 이내 큰 금액"))
    return out


def journal_entries(books: Books, materiality: float, plan: Plan, limit: int = JE_MAX) -> None:
    if not books.journal:
        plan.notes.append("특이분개: 분개장이 없어 건너뜀")
        return
    end = period_end(books.periods.get("current", ""))
    scored = []
    for (d, no), lines in vouchers(books.journal).items():
        reasons = je_reasons(lines, materiality, end)
        score = sum(s for s, _ in reasons)
        # 금액 기준 하나만 걸린 평범한 큰 거래는 빼고, 성격상 특이한 점이 하나라도 있어야 고른다
        if score >= 3 and any(s >= 2 and "중요성" not in r for s, r in reasons):
            scored.append((score, _amount(lines), d, no, lines, reasons))
    scored.sort(key=lambda x: (-x[0], -x[1]))
    plan.notes.append(f"특이분개: 전표 {len(vouchers(books.journal)):,}건 중 기준에 걸린 {len(scored):,}건에서 "
                      f"점수·금액 순으로 {min(limit, len(scored))}건")
    for score, amt, d, no, lines, reasons in scored[:limit]:
        body = "; ".join(
            f"{'차' if j['debit'] else '대'} {j['account_name']} {_won(j['debit'] or j['credit'])}"
            + (f"({j['partner']})" if j["partner"] else "") for j in lines[:6])
        if len(lines) > 6:
            body += f" 외 {len(lines) - 6}줄"
        memo = next((j["memo"] for j in lines if j["memo"]), "")
        plan.drafts.append(Draft(
            f"JE:{d.isoformat()}:{no}", "특이분개", lines[0]["account_name"], amt,
            f"{d.isoformat()} 전표 {no} ({body}{', 적요: ' + memo if memo else ''}) — {', '.join(r for _, r in reasons)}. "
            f"거래 내용과 이렇게 회계처리한 이유를 설명해 주시고, 전표 승인자와 근거 자료(계약서·세금계산서·품의서 등)를 보내 주세요."))


# ── 5. 표본 증빙 요청 ─────────────────────────────────
def sales_vouchers(vs: dict) -> list[tuple]:
    return [(k, sum(j["credit"] for j in v if _is_sales(j["account_name"]))) for k, v in vs.items()
            if any(_is_sales(j["account_name"]) and j["credit"] > 0 for j in v)]


def purchase_vouchers(vs: dict) -> list[tuple]:
    out = []
    for k, v in vs.items():
        inv = sum(j["debit"] for j in v if j["account_name"] in PURCHASE_DEBIT)
        pay = sum(j["credit"] for j in v if j["account_name"] in PURCHASE_CREDIT)
        if inv > 0 or pay > 0:
            out.append((k, inv or pay))
    return out


def samples(books: Books, plan: Plan, n: int = SAMPLE_N, seed: int | None = None,
            skip: tuple[str, ...] = ()) -> None:
    if not books.journal:
        plan.notes.append("표본: 분개장이 없어 건너뜀")
        return
    seed = seed if seed is not None else secrets.randbelow(10**6)
    plan.seed = seed
    vs = vouchers(books.journal)
    for label, pop, ask in (
            ("매출", sales_vouchers(vs), "세금계산서(또는 계산서·영수증), 거래명세서·출고증, 입금 내역(통장 사본)"),
            ("매입", purchase_vouchers(vs), "세금계산서, 거래명세서·입고증(검수 확인), 지급 내역(통장 사본)")):
        if label in skip:
            plan.notes.append(f"표본: {label} 표본은 이미 만들어 둔 것이 있어 건너뜀 (다시 뽑으려면 기존 표본 질문을 지우세요)")
            continue
        pop = [x for x in pop if x[1] > 0]
        if not pop:
            plan.notes.append(f"표본: 분개장에서 {label}전표를 못 찾아 건너뜀")
            continue
        rng = random.Random(f"{seed}:{label}")
        picked = sorted(rng.sample(pop, min(n, len(pop))), key=lambda x: x[0])
        plan.notes.append(f"표본: {label}전표 {len(pop):,}건 중 무작위 {len(picked)}건 (시드 {seed})")
        for (d, no), amt in picked:
            lines = vs[(d, no)]
            partner = next((j["partner"] for j in lines if j["partner"]), "")
            memo = next((j["memo"] for j in lines if j["memo"]), "")
            plan.drafts.append(Draft(
                f"샘플:{label}:{d.isoformat()}:{no}", f"표본({label})", f"{label} 표본", amt,
                f"[증빙 요청] {label}전표 {d.isoformat()} {no}"
                f"{' ' + partner if partner else ''} {_won(amt)}{' (' + memo + ')' if memo else ''}: "
                f"{ask}을 보내 주세요. (표본 시드 {seed})"))


# ── 묶어서 실행 ───────────────────────────────────────
def build(books: Books, materiality: float, seed: int | None = None, have: set[str] | None = None) -> Plan:
    if not materiality or materiality <= 0:
        raise ValueError("중요성 금액을 넣어 주세요.")
    have = have or set()
    plan = Plan()
    variance(books, materiality, plan)
    allowance(books, plan)
    cost_ratio(books, plan)
    journal_entries(books, materiality, plan)
    skip = tuple(lbl for lbl in ("매출", "매입") if any(k.startswith(f"샘플:{lbl}:") for k in have))
    samples(books, plan, seed=seed, skip=skip)
    return plan


def apply(store, eid: int, plan: Plan, today: date | None = None) -> tuple[int, int]:
    """초안을 질문 목록에 '작성' 상태로 넣는다. 반환: (새로 넣은 건수, 이미 있어 건너뛴 건수)"""
    from . import questions

    have = {q["source"] for q in questions.all_(store, eid)}
    added = skipped = 0
    for d in plan.drafts:
        if d.key in have:
            skipped += 1
            continue
        questions.add(store, eid, d.text, d.account, d.amount, d.key, today)
        have.add(d.key)
        added += 1
    return added, skipped
