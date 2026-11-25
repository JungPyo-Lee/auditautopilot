"""여러 파일을 한 번에 받아 종류를 알아서 맞춰 하나의 장부 묶음으로 만든다.

- 세무사랑 '재무제표 등' 묶음 파일 하나
- 또는 더존 파일 여러 개: 합계잔액시산표, 재무상태표, 손익계산서, 거래처원장(계정마다 한 파일), 분개장

명령 예)
  python -m auditautopilot analytics 재무상태표.xlsx 손익계산서.xlsx -o 분석.xlsx --threshold-amount ...
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import openpyxl

from . import douzone, semusarang
from .common import squash


@dataclass
class Books:
    tb: list[dict] = field(default_factory=list)
    bs: list[dict] = field(default_factory=list)
    pl: list[dict] = field(default_factory=list)
    journal: list[dict] = field(default_factory=list)
    partners: list[dict] = field(default_factory=list)
    periods: dict[str, str] = field(default_factory=dict)
    source: str = ""
    notes: list[str] = field(default_factory=list)


_BS_SECTIONS = {"유동자산", "당좌자산", "재고자산", "비유동자산", "투자자산", "유형자산", "무형자산",
                "기타비유동자산", "유동부채", "비유동부채"}


def period_end(label: str) -> date | None:
    """'제 38(당)기[2026/01/01 ~ 2026/10/31]' 또는 '2026년1월1일 ~ 2026년9월30일' -> 끝 날짜"""
    m = re.findall(r"(\d{4})\s*[/.\-년]\s*(\d{1,2})\s*[/.\-월]\s*(\d{1,2})", label or "")
    if len(m) >= 2:
        y, mo, d = (int(x) for x in m[1])
        return date(y, mo, d)
    return None


def _sheetnames(path: Path) -> list[str]:
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        return wb.sheetnames
    finally:
        wb.close()


def load(paths, ledger_names: dict[str, str] | None = None) -> Books:
    """ledger_names: 거래처원장 파일 이름 -> 계정명. 금액으로 못 찾을 때 사람이 정해 준다."""
    paths = [Path(p) for p in ([paths] if isinstance(paths, (str, Path)) else paths)]
    for p in paths:
        if not p.exists():
            raise FileNotFoundError(f"파일이 없습니다: {p}")
    if len(paths) == 1 and {"TB", "BS"} & set(_sheetnames(paths[0])):
        b = semusarang.load(paths[0])
        return Books(b.tb, b.bs, b.pl, b.journal, b.partners, b.periods, "세무사랑")

    books = Books(source="더존")
    ledgers: list[tuple[Path, list[dict], dict | None]] = []
    for p in paths:
        rows = douzone._rows_of(p)
        kind = douzone.detect_kind(rows)
        if kind in ("bs", "pl"):
            data, periods = douzone.read_fs_rows(rows, kind)
            setattr(books, kind, data)
            books.periods = books.periods or periods
        elif kind == "tb":
            books.tb = douzone.read_tb_rows(rows)
        elif kind == "ledger":
            ledgers.append((p, *douzone.read_partner_ledger_rows(rows)))
        elif kind == "journal":
            books.journal.extend(douzone.read_journal_rows(rows))
        elif kind == "partner_list":
            books.notes.append(f"{p.name}: 거래처 이름만 있는 파일이라 건너뜀 (거래처원장 '잔액' 화면으로 다시 내보내 주세요)")
        else:
            books.notes.append(f"{p.name}: 형식을 알아보지 못해 건너뜀")

    # 계정코드: 더존 재무제표에는 없으므로 분개장에서 과목명으로 찾아 채운다
    if books.journal:
        code_of: dict[str, set[str]] = {}
        for j in books.journal:
            if j["account_code"]:
                code_of.setdefault(j["account_name"], set()).add(j["account_code"])
        for r in books.bs + books.pl:
            codes = code_of.get(r["name"])
            if r["is_account"] and codes and len(codes) == 1:
                r["code"] = next(iter(codes))

    # 분개장이 재무제표 기간보다 길면 기간 끝까지만 쓴다
    end = period_end(books.periods.get("current", ""))
    if end and books.journal:
        before = len(books.journal)
        books.journal = [j for j in books.journal if j["date"] is None or j["date"] <= end]
        if len(books.journal) < before:
            books.notes.append(f"분개장에서 재무제표 기간({end}) 뒤의 {before - len(books.journal):,}줄은 뺐습니다")

    # 거래처원장이 어느 계정인지: 시산표 잔액과 먼저 맞추고, 없으면 재무상태표와 맞춘다
    tb_accounts, section = [], ""
    for r in books.tb:
        if r["level"] == "group":
            section = r["name"]
        elif r["level"] == "account" and section in _BS_SECTIONS:
            tb_accounts.append({"name": r["name"], "code": r["code"], "is_account": True,
                                "current": r["balance"], "prior": 0.0})
    for p, rows, total in ledgers:
        name = (ledger_names or {}).get(p.name) or (ledger_names or {}).get(str(p))
        stem = squash(re.sub(r"[_ ]?\d{8}.*$", "", p.stem))
        named = name is not None
        if name is None and stem in {r["name"] for r in tb_accounts + books.bs if r.get("is_account")}:
            name, named = stem, True      # 파일 이름이 계정명이면 그것을 따른다
        if name is None and total is not None:
            for cands in (tb_accounts, books.bs):
                if cands and name is None:
                    name = douzone.match_ledger_account(total, cands)
        if name is None:
            name = re.sub(r"[_ ]?\d{8}.*$", "", p.stem) or p.stem
            books.notes.append(f"{p.name}: 합계 잔액과 같은 시산표·재무상태표 계정을 못 찾아 '{name}'(으)로 둠")
        elif total is not None and not named:
            books.notes.append(f"{p.name}: 합계 잔액 {total['balance']:,.0f}원 = '{name}'")
        code = next((r["code"] for r in tb_accounts + books.bs if r["name"] == name and r.get("code")), None)
        for r in rows:
            r["account_name"] = name
            r["account_code"] = code
        books.partners.extend(rows)
    return books


def partners_of(books: Books, account: str) -> list[dict]:
    """계정코드(108) 또는 계정명(외상매출금)으로 거래처 잔액을 고른다."""
    key = account.strip()
    code = str(int(key)).zfill(3) if key.isdigit() else None
    return [r for r in books.partners
            if (code and r.get("account_code") == code) or (r.get("account_name") or "") == key]


def journal_of(books: Books, account: str) -> list[dict]:
    key = account.strip()
    code = str(int(key)).zfill(3) if key.isdigit() else None
    return [r for r in books.journal if (code and r["account_code"] == code) or r["account_name"] == key]
