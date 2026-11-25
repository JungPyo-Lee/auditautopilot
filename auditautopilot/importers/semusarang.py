"""세무사랑 '재무제표 등' 엑셀 묶음 파일 불러오기.

확인한 형식 (2026-11 받은 샘플 기준):
- TB 시트: 차변(잔액, 합계) | 계정과목 | 대변(합계, 잔액). 과목명에 '[0103] 보 통 예 금'처럼 코드가 붙는다.
- BS, PL 시트: 과목 | 당기 금액 | 당기 집계금액 | 전기 금액 | 전기 집계금액.
  재무상태표는 계정 금액도 집계금액 열에, 손익계산서는 계정 금액이 금액 열에 있다.
- 분개장 시트: 머리글 2줄. 둘째 줄이 실제 열 이름이다.
- 숫자 이름 시트(예: '118'): 그 계정의 거래처원장. 코드, 거래처, 등록번호, ... 전기이월, 차변, 대변, 잔액.

각 표는 dict 목록이며 키는 common.py의 *_COLUMNS 를 따른다.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from .common import (
    cell, column_positions, find_header_row, mask_reg_no, normalize_code,
    parse_account_label, parse_date, squash, to_number,
)

Rows = list[dict]


@dataclass
class SemusarangBundle:
    tb: Rows
    bs: Rows
    pl: Rows
    journal: Rows
    partners: Rows
    periods: dict[str, str] = field(default_factory=dict)


def _rows(ws) -> list[tuple]:
    return [tuple(r) for r in ws.iter_rows(values_only=True)]


def read_tb(rows: list[tuple]) -> Rows:
    header = find_header_row(rows, ["잔액", "합계", "계정과목"])
    pos = column_positions(rows[header])
    d_bal, c_bal = pos["잔액"][0], pos["잔액"][-1]
    d_tot, c_tot = pos["합계"][0], pos["합계"][-1]
    name_col = pos["계정과목"][0]
    records = []
    for row in rows[header + 1:]:
        label = cell(row, name_col)
        if label is None or not str(label).strip():
            continue
        code, name = parse_account_label(label)
        level = "account" if code else ("total" if name == "합계" else "group")
        db, cb = to_number(cell(row, d_bal)), to_number(cell(row, c_bal))
        records.append({
            "code": code, "name": name, "level": level,
            "debit_balance": db, "debit_total": to_number(cell(row, d_tot)),
            "credit_total": to_number(cell(row, c_tot)), "credit_balance": cb,
            "balance": db - cb,
        })
    return records


def read_fs(rows: list[tuple]) -> tuple[Rows, dict[str, str]]:
    header = find_header_row(rows, ["과목", "금액"])
    periods: dict[str, str] = {}
    if header > 0:
        labels = [v for v in rows[header - 1][1:] if v]
        if labels:
            periods["current"] = str(labels[0]).strip()
        if len(labels) > 1:
            periods["prior"] = str(labels[1]).strip()
    records = []
    for row in rows[header + 1:]:
        label = cell(row, 0)
        if label is None or not str(label).strip():
            continue
        code, name = parse_account_label(label)
        # 재무상태표는 계정 금액도 '집계금액' 열에, 손익계산서는 '금액' 열에 들어 있다.
        current = to_number(cell(row, 1)) or to_number(cell(row, 2))
        prior = to_number(cell(row, 3)) or to_number(cell(row, 4))
        records.append({"code": code, "name": name, "is_account": bool(code),
                        "current": current, "prior": prior})
    return records, periods


def _nearest_before(positions: list[int], limit: int) -> int | None:
    before = [p for p in positions if p < limit]
    return before[-1] if before else None


def read_journal(rows: list[tuple]) -> Rows:
    header = find_header_row(rows, ["전표번호", "차변", "대변", "거래처명"])
    pos = column_positions(rows[header])
    codes = pos.get("코드", [])
    account_col = pos["계정과목"][0]
    partner_col = pos["거래처명"][0]
    date_col = (pos.get("일자") or pos.get("전표일자"))[0]
    memo_col = pos.get("적요", [None])[0]
    acc_code_col = _nearest_before(codes, account_col)
    partner_code_col = _nearest_before(codes, partner_col)
    records = []
    for row in rows[header + 1:]:
        if cell(row, date_col) is None:
            continue
        records.append({
            "date": parse_date(cell(row, date_col)),
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


def read_partner_ledger(rows: list[tuple], account_code: str) -> Rows:
    header = find_header_row(rows, ["코드", "거래처", "잔액"])
    pos = column_positions(rows[header])
    records = []
    for row in rows[header + 1:]:
        code = cell(row, pos["코드"][0])
        if code is None or not str(code).strip():
            continue
        records.append({
            "account_code": normalize_code(account_code),
            "partner_code": normalize_code(code),
            "partner": str(cell(row, pos["거래처"][0]) or "").strip(),
            "reg_no": mask_reg_no(cell(row, pos.get("등록번호", [None])[0])),
            "opening": to_number(cell(row, pos.get("전기이월", [None])[0])),
            "debit": to_number(cell(row, pos["차변"][0])),
            "credit": to_number(cell(row, pos["대변"][0])),
            "balance": to_number(cell(row, pos["잔액"][0])),
        })
    return records


def load(path: str | Path) -> SemusarangBundle:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        names = wb.sheetnames
        tb = read_tb(_rows(wb["TB"])) if "TB" in names else []
        periods: dict[str, str] = {}
        bs: Rows = []
        pl: Rows = []
        if "BS" in names:
            bs, periods = read_fs(_rows(wb["BS"]))
        if "PL" in names:
            pl, pl_periods = read_fs(_rows(wb["PL"]))
            periods = periods or pl_periods
        journal = read_journal(_rows(wb["분개장"])) if "분개장" in names else []
        partners: Rows = []
        for n in names:
            if n.strip().isdigit():
                partners.extend(read_partner_ledger(_rows(wb[n]), n))
    finally:
        wb.close()
    return SemusarangBundle(tb=tb, bs=bs, pl=pl, journal=journal, partners=partners, periods=periods)
