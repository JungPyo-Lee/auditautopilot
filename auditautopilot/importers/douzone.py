"""더존 엑셀 내보내기 파일 불러오기.

2026-11 받은 샘플 기준.
- 분개장: 단일 시트, 머리글 1줄.
  전표일자 | 전표번호 | 구분 | Code | 계정과목 | 차변 | 대변 | 거래처 | Code | 적요
- 재무상태표·손익계산서(과목별): 과목 | 당기 금액 | 당기 집계 | 전기 금액 | 전기 집계. 머리글 2줄.
  계정코드가 없어서 과목명으로 계정을 구분한다. 금액은 '1,234' 같은 글자일 때도 있다.
  차감계정(대손충당금, 감가상각누계액)은 왼쪽 칸에 자기 금액, 오른쪽 칸에 차감 후 금액이 있다.
- 거래처원장(잔액): 거래처 | 전기(월)이월 | 차변 | 대변 | 잔액, 마지막 줄 '합계'.
  어느 계정의 원장인지 파일에 없으므로 합계 잔액을 재무상태표 계정 금액과 맞춰 찾는다.

합계잔액시산표는 샘플을 받으면 추가한다.
"""

from __future__ import annotations

import re
from pathlib import Path

import openpyxl

from .common import (
    cell, column_positions, find_header_row, normalize_code, parse_date, squash, to_number,
)

_SECTION_RE = re.compile(r"^\s*([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]+\s*\.|\(\s*\d+\s*\))")
_TOP_LABELS = {"자산", "부채", "자본"}


def read_journal_rows(rows: list[tuple]) -> list[dict]:
    header = find_header_row(rows, ["전표일자", "전표번호", "차변", "대변"])
    pos = column_positions(rows[header])
    codes = sorted(pos.get("Code", []) + pos.get("코드", []))
    account_col = pos["계정과목"][0]
    partner_col = (pos.get("거래처") or pos.get("거래처명"))[0]
    acc_code_col = max((p for p in codes if p < account_col), default=None)
    partner_code_col = min((p for p in codes if p > partner_col), default=None)
    memo_col = pos.get("적요", [None])[0]
    records = []
    for row in rows[header + 1:]:
        raw_date = cell(row, pos["전표일자"][0])
        when = parse_date(raw_date)
        if when is None:
            continue  # 빈 줄, 마지막 '합계 (건수 : ...)' 줄
        records.append({
            "date": when,
            "voucher_no": str(cell(row, pos["전표번호"][0]) or "").strip(),
            "side": str(cell(row, pos["구분"][0]) or "").strip(),
            "account_code": normalize_code(cell(row, acc_code_col)),
            "account_name": squash(cell(row, account_col)),
            "debit": to_number(cell(row, pos["차변"][0])),
            "credit": to_number(cell(row, pos["대변"][0])),
            "partner_code": normalize_code(cell(row, partner_code_col)),
            "partner": str(cell(row, partner_col) or "").strip(),
            "memo": str(cell(row, memo_col) or "").strip(),
        })
    return records


def load_journal(path: str | Path) -> list[dict]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    return read_journal_rows(rows)


def _blank(v: object) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def _rows_of(path: str | Path) -> list[tuple]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        return [tuple(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()


def detect_kind(rows: list[tuple]) -> str | None:
    """첫 몇 줄을 보고 더존 파일 종류를 맞힌다: journal, bs, pl, ledger, partner_list."""
    head = [squash(v) for r in rows[:5] for v in r if v is not None]
    names = [squash(cell(r, 0)) for r in rows[:120]]
    if {"전표일자", "전표번호"} <= set(head):
        return "journal"
    if "거래처" in head and "잔액" in head and any(h.startswith("전기") for h in head):
        return "ledger"
    if head and head[0] in ("과목",):
        if any(n.endswith("자산총계") for n in names) or "자산" in names:
            return "bs"
        if any("매출액" in n or "당기순이익" in n for n in names):
            return "pl"
    if head[:1] == ["거래처"] and all(len(r) <= 1 or all(_blank(v) for v in r[1:]) for r in rows[:20]):
        return "partner_list"
    return None


def read_fs_rows(rows: list[tuple], kind: str = "bs") -> tuple[list[dict], dict[str, str]]:
    """재무상태표·손익계산서(과목별)를 표준 행으로 바꾼다.

    결과 행: code(None), name, is_account, current, prior, level(들여쓰기 칸 수)
    차감계정은 '대손충당금(외상매출금)'처럼 앞 계정 이름을 붙여 구분한다.
    """
    top = next(i for i, r in enumerate(rows) if squash(cell(r, 0)) == "과목")
    periods: dict[str, str] = {}
    labels = [str(v).strip() for v in rows[top][1:] if not _blank(v)]
    if labels:
        periods["current"] = labels[0]
    if len(labels) > 1:
        periods["prior"] = labels[1]
    start = top + 1
    if start < len(rows) and _blank(cell(rows[start], 0)):
        start += 1  # '금액' 머리글 둘째 줄
    out: list[dict] = []
    last_account = None
    for row in rows[start:]:
        raw = cell(row, 0)
        if _blank(raw):
            continue
        nums = [cell(row, i) for i in (1, 2, 3, 4)]
        if all(_blank(v) for v in nums):
            continue  # '자산', '(당기순이익)', '당기 : ... 원' 같은 줄
        text = str(raw)
        name = squash(text)
        level = len(text) - len(text.lstrip(" "))
        header = bool(_SECTION_RE.match(text)) or name in _TOP_LABELS or name.endswith(("총계", "합계")) \
            or level == 0
        cur_d, cur_s, pri_d, pri_s = nums
        if kind == "pl" and not header and _blank(cur_d) and _blank(pri_d):
            header = True  # 손익계산서의 '상품매출원가'처럼 오른쪽 칸만 있는 줄은 소계
        current = to_number(cur_s if _blank(cur_d) else cur_d)
        prior = to_number(pri_s if _blank(pri_d) else pri_d)
        contra = kind == "bs" and not header and not _blank(cur_d) and not _blank(cur_s) and last_account
        if contra:
            name = f"{name}({last_account})"
        out.append({"code": None, "name": name, "is_account": not header,
                    "current": current, "prior": prior, "level": level})
        if not header and not contra:
            last_account = name
    return out, periods


def load_fs(path: str | Path) -> tuple[str, list[dict], dict[str, str]]:
    """재무상태표 또는 손익계산서 파일. 반환: (kind, 행, 기간)"""
    rows = _rows_of(path)
    kind = detect_kind(rows)
    if kind not in ("bs", "pl"):
        raise ValueError(f"{Path(path).name}: 더존 재무상태표·손익계산서(과목별) 형식이 아닙니다.")
    data, periods = read_fs_rows(rows, kind)
    return kind, data, periods


def read_partner_ledger_rows(rows: list[tuple]) -> tuple[list[dict], dict | None]:
    """거래처원장(잔액). 반환: (행, 합계 줄 {'opening', 'balance'} 또는 None)"""
    header = find_header_row(rows, ["거래처", "차변", "대변", "잔액"])
    pos = column_positions(rows[header])
    opening_col = next((p[0] for k, p in pos.items() if k.startswith("전기")), None)
    out, total = [], None
    for row in rows[header + 1:]:
        partner = cell(row, pos["거래처"][0])
        if _blank(partner):
            continue
        if squash(partner) in ("합계", "총계"):
            total = {"opening": to_number(cell(row, opening_col)), "balance": to_number(cell(row, pos["잔액"][0]))}
            continue
        out.append({
            "account_code": None,
            "account_name": None,
            "partner_code": None,
            "partner": str(partner).strip(),
            "reg_no": "",
            "opening": to_number(cell(row, opening_col)),
            "debit": to_number(cell(row, pos["차변"][0])),
            "credit": to_number(cell(row, pos["대변"][0])),
            "balance": to_number(cell(row, pos["잔액"][0])),
        })
    return out, total


def match_ledger_account(total: dict, bs: list[dict]) -> str | None:
    """원장 합계 잔액과 금액이 같은 재무상태표 계정 이름.
    여러 계정이 맞으면 전기이월 합계와 전기 금액까지 같은 계정으로 좁힌다. 하나로 정해질 때만 돌려준다."""
    def same(a, b):
        return abs(abs(a) - abs(b)) < 0.5
    hits = [r for r in bs if r.get("is_account") and same(r["current"], total["balance"])]
    if len(hits) > 1:
        hits = [r for r in hits if same(r["prior"], total["opening"])]
    return hits[0]["name"] if len(hits) == 1 else None
