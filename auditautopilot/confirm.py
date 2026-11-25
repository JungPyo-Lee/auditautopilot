"""채권채무 전자조회 회신현황 불러오기와 대사.

전자조회 사이트에서 내려받은 회신현황 엑셀(기업명·구분·발송금액·회신금액·일치여부 …)을 읽어
- 거래처·계정마다 일치 / 불일치 / 미회신으로 나누고
- 장부(더존 거래처원장)가 있으면 발송금액이 장부 잔액과 같은지 다시 확인하고
- 재무상태표가 있으면 계정별 조회 비율(발송금액 ÷ 계정 잔액)을 낸다.

불일치는 차이조정, 미회신은 대체적 절차(후속 입금·지급, 세금계산서 등) 대상으로 따로 뽑는다.
금융기관 조회는 아직 다루지 않는다.

명령 예)
  python -m auditautopilot confirm 회신현황.xlsx -o 채권채무조회.xlsx
  python -m auditautopilot confirm 회신현황.xlsx 재무상태표.xlsx 거래처원장*.xlsx -o 채권채무조회.xlsx
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from .importers import books as books_mod
from .importers.common import cell, column_positions, find_header_row, squash, to_number
from .importers.douzone import _rows_of

# 머리글 이름이 조금 달라도 읽도록 여러 이름을 받는다
COLUMNS = {
    "partner": ["기업명", "거래처명", "거래처", "회사명"],
    "round": ["조회차수", "차수"],
    "biz_no": ["사업자등록번호", "사업자번호", "등록번호"],
    "contact": ["담당자명", "담당자"],
    "email": ["담당자이메일", "이메일", "담당자메일"],
    "agree": ["일치여부", "일치"],
    "currency": ["통화", "통화코드"],
    "label": ["구분", "계정", "계정과목", "조회항목"],
    "sent": ["발송금액", "조회금액", "장부금액"],
    "replied": ["회신금액"],
    "diff": ["차이금액", "차이"],
    "sent_at": ["최초발송일시", "발송일시", "발송일"],
    "replied_at": ["회신일시", "회신일"],
}
REQUIRED = ("partner", "label", "sent", "replied")

# 영문 구분 → 계정과목 (해외 거래처는 영문으로 나온다)
ENGLISH = {
    "accountsreceivable": "외상매출금", "tradereceivables": "외상매출금", "tradereceivable": "외상매출금",
    "accountspayable": "외상매입금", "tradepayables": "외상매입금", "tradepayable": "외상매입금",
    "otherreceivables": "미수금", "otherreceivable": "미수금",
    "otherpayables": "미지급금", "otherpayable": "미지급금", "nontradepayables": "미지급금",
    "notesreceivable": "받을어음", "notespayable": "지급어음",
    "advancesreceived": "선수금", "advancereceived": "선수금",
    "advancepayments": "선급금", "advancepayment": "선급금", "prepayments": "선급금",
    "loansreceivable": "단기대여금", "loanspayable": "단기차입금", "borrowings": "단기차입금",
    "deposits": "보증금", "depositsreceived": "임대보증금", "guaranteedeposits": "임차보증금",
}

STATUS_OK, STATUS_DIFF, STATUS_NONE = "일치", "불일치", "미회신"


def account_of(label: object) -> str:
    """'외상매출금(커머스)' -> '외상매출금', 'accounts receivable' -> '외상매출금'"""
    text = re.sub(r"[（(].*?[)）]", "", str(label or "")).strip()
    key = re.sub(r"[^a-z]", "", text.lower())
    if key and key in ENGLISH:
        return ENGLISH[key]
    return squash(text)


def norm_name(name: object) -> str:
    """거래처 이름 비교용: (주)·주식회사·공백·전각 괄호를 뺀다."""
    t = str(name or "").replace("（", "(").replace("）", ")")
    t = re.sub(r"\(\s*(주|유|사|재|합)\s*\)|㈜|주식회사|유한회사|\b(co\.?,?\s*ltd\.?|inc\.?|gmbh|corp\.?)\b", "", t, flags=re.I)
    return re.sub(r"[\s.,·\-]", "", t).lower()


def digits(v: object) -> str:
    return re.sub(r"\D", "", str(v or ""))


def _when(v: object):
    if isinstance(v, datetime):
        return v.replace(microsecond=0)
    if isinstance(v, date):
        return v
    s = str(v or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return datetime.strptime(s[:19], fmt)
        except ValueError:
            pass
    return s


def is_status_file(rows: list[tuple]) -> bool:
    try:
        _header(rows)
        return True
    except ValueError:
        return False


def _header(rows: list[tuple]) -> tuple[int, dict[str, int]]:
    for i, row in enumerate(rows[:30]):
        pos = column_positions(row)
        found = {}
        for key, names in COLUMNS.items():
            for n in names:
                if squash(n) in pos:
                    found[key] = pos[squash(n)][0]
                    break
        if all(k in found for k in REQUIRED):
            return i, found
    raise ValueError("전자조회 회신현황 형식이 아닙니다 (기업명·구분·발송금액·회신금액 머리글을 찾지 못함).")


def read_status_rows(rows: list[tuple]) -> list[dict]:
    head, col = _header(rows)
    out = []
    for row in rows[head + 1:]:
        partner = cell(row, col["partner"])
        if partner is None or not str(partner).strip() or squash(partner) in ("합계", "총계"):
            continue
        g = lambda k: cell(row, col.get(k))  # noqa: E731
        sent = to_number(g("sent"))
        raw_reply = g("replied")
        replied_at = _when(g("replied_at"))
        replied = None if (raw_reply is None or str(raw_reply).strip() == "") else to_number(raw_reply)
        agree = str(g("agree") or "").strip().upper()
        if replied is None and not replied_at:
            status = STATUS_NONE
        elif agree in ("Y", "일치", "O") or (agree == "" and replied is not None and abs(replied - sent) < 0.005):
            status = STATUS_OK
        else:
            status = STATUS_DIFF
        diff = g("diff")
        diff = to_number(diff) if diff not in (None, "") else (None if replied is None else replied - sent)
        currency = str(g("currency") or "KRW").strip().upper() or "KRW"
        r = {
            "partner": str(partner).strip(),
            "biz_no": digits(g("biz_no")),
            "contact": str(g("contact") or "").strip(),
            "email": str(g("email") or "").strip(),
            "round": int(to_number(g("round"))) if g("round") not in (None, "") else 1,
            "label": str(g("label") or "").strip(),
            "account": account_of(g("label")),
            "currency": currency,
            "sent": sent,
            "replied": replied,
            "diff": diff if status != STATUS_NONE else None,
            "status": status,
            "sent_at": _when(g("sent_at")),
            "replied_at": replied_at,
            "note": "",
        }
        if status == STATUS_DIFF and replied and sent and (abs(replied / sent) > 10 or abs(sent / replied) > 10):
            r["note"] = "회신 금액 단위(통화·천원 등)가 발송과 다른 것으로 보임 - 회신서 원문 확인"
        out.append(r)
    # 같은 거래처·계정·통화를 여러 차수로 보냈으면 마지막 차수만 쓴다
    latest: dict[tuple, dict] = {}
    for r in out:
        k = (r["biz_no"] or norm_name(r["partner"]), r["label"], r["currency"])
        if k not in latest or r["round"] >= latest[k]["round"]:
            latest[k] = r
    return list(latest.values())


def load_status(path: str | Path) -> list[dict]:
    return read_status_rows(_rows_of(path))


@dataclass
class Result:
    rows: list[dict]
    summary: list[dict]
    notes: list[str] = field(default_factory=list)

    def by_status(self, status: str) -> list[dict]:
        return [r for r in self.rows if r["status"] == status]

    def sheets(self) -> dict[str, list[dict]]:
        def show(rows, extra=None):
            if not rows:
                return [{"거래처": "해당 없음"}]
            out = []
            for r in rows:
                d = {"거래처": r["partner"], "사업자등록번호": r["biz_no"], "계정": r["account"], "조회 구분": r["label"],
                     "통화": r["currency"], "발송금액": r["sent"], "회신금액": r["replied"], "차이(회신-발송)": r["diff"],
                     "결과": r["status"]}
                if "book" in r:
                    d["장부잔액"] = r["book"]
                    d["장부-발송 차이"] = r["book_diff"]
                d.update({"담당자": r["contact"], "이메일": r["email"], "발송일시": r["sent_at"],
                          "회신일시": r["replied_at"], "확인할 점": r["note"]})
                d.update(extra or {})
                out.append(d)
            return out
        return {
            "요약": self.summary,
            "불일치_차이조정": show(self.by_status(STATUS_DIFF),
                              {"차이 원인": "", "조정 내역(시점차이·미착 등)": "", "조정 후 차이": "", "결론": ""}),
            "미회신_대체절차": show(self.by_status(STATUS_NONE),
                              {"대체적 절차(후속 입금·지급, 세금계산서, 계약서 등)": "", "확인 금액": "", "결론": ""}),
            "명세": show(self.rows),
            "참고": [{"내용": n} for n in self.notes] or [{"내용": "-"}],
        }


def reconcile(rows: list[dict], books: books_mod.Books | None = None) -> Result:
    notes: list[str] = []
    rows = sorted(rows, key=lambda r: (r["account"], r["status"] != STATUS_DIFF, -abs(r["sent"])))
    if books and books.partners:
        ledger: dict[tuple[str, str], float] = {}
        accounts = {p.get("account_name") for p in books.partners}
        for p in books.partners:
            for key in filter(None, (digits(p.get("reg_no")), norm_name(p["partner"]))):
                k = (p.get("account_name") or "", key)
                ledger[k] = ledger.get(k, 0.0) + p["balance"]
        for r in rows:
            if r["account"] not in accounts or r["currency"] != "KRW":
                continue
            book = ledger.get((r["account"], r["biz_no"])) if r["biz_no"] else None
            if book is None:
                book = ledger.get((r["account"], norm_name(r["partner"])))
            r["book"] = book
            r["book_diff"] = None if book is None else abs(book) - r["sent"]
            if book is None:
                r["note"] = (r["note"] + " / " if r["note"] else "") + "장부(거래처원장)에서 이 거래처를 못 찾음"
            elif abs(r["book_diff"]) >= 0.5:
                r["note"] = (r["note"] + " / " if r["note"] else "") + "발송금액이 장부 잔액과 다름 - 조회 기준일·금액 확인"
        notes.append(f"거래처원장이 있는 계정({', '.join(sorted(a for a in accounts if a))})은 발송금액을 장부 잔액과 다시 맞춰 봤습니다.")

    bs_balance = {}
    if books and books.bs:
        for b in books.bs:
            if b.get("is_account"):
                bs_balance.setdefault(b["name"], b["current"])
    summary = []
    for acc in sorted({r["account"] for r in rows}):
        group = [r for r in rows if r["account"] == acc]
        krw = [r for r in group if r["currency"] == "KRW"]
        amt = lambda rs, st=None: sum(r["sent"] for r in rs if st is None or r["status"] == st)  # noqa: E731
        cnt = lambda st: sum(1 for r in group if r["status"] == st)  # noqa: E731
        line = {
            "계정": acc,
            "조회 건수": len(group),
            "일치": cnt(STATUS_OK), "불일치": cnt(STATUS_DIFF), "미회신": cnt(STATUS_NONE),
            "발송금액(원화)": amt(krw),
            "일치 금액(원화)": amt(krw, STATUS_OK),
            "불일치 금액(원화)": amt(krw, STATUS_DIFF),
            "미회신 금액(원화)": amt(krw, STATUS_NONE),
            "외화 건수": len(group) - len(krw),
        }
        if acc in bs_balance:
            bal = abs(bs_balance[acc])
            line["재무상태표 잔액"] = bal
            line["조회 비율(원화)"] = round(line["발송금액(원화)"] / bal * 100, 1) if bal else None
        summary.append(line)
    if any(r["currency"] != "KRW" for r in rows):
        notes.append("외화 조회 건은 환율을 곱하지 않아 원화 합계와 조회 비율에서 뺐습니다. 기말 환율로 따로 환산해 주세요.")
    if summary and not bs_balance:
        notes.append("재무상태표 파일을 함께 주면 계정별 조회 비율을 계산합니다.")
    return Result(rows, summary, notes)


def build(paths) -> Result:
    """회신현황 파일과 (있으면) 더존 재무상태표·거래처원장을 함께 받아 대사 결과를 만든다."""
    paths = [Path(p) for p in paths]
    status, others = [], []
    for p in paths:
        if not p.exists():
            raise FileNotFoundError(f"파일이 없습니다: {p}")
        rows = _rows_of(p)
        (status if is_status_file(rows) else others).append((p, rows))
    if not status:
        raise ValueError("전자조회 회신현황 파일이 없습니다 (기업명·구분·발송금액·회신금액 머리글이 있는 파일).")
    data = [r for _, rows in status for r in read_status_rows(rows)]
    books = books_mod.load([p for p, _ in others]) if others else None
    res = reconcile(data, books)
    if books:
        res.notes = books.notes + res.notes
    return res
