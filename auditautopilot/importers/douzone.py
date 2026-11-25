"""더존 엑셀 내보내기 파일 불러오기.

지금은 분개장만 지원한다 (2026-11 받은 샘플 기준).
- 분개장: 단일 시트, 머리글 1줄.
  전표일자 | 전표번호 | 구분 | Code | 계정과목 | 차변 | 대변 | 거래처 | Code | 적요

합계잔액시산표, 재무상태표·손익계산서, 거래처원장은 샘플 파일을 받은 뒤 추가한다.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pandas as pd

from .common import (
    JOURNAL_COLUMNS, cell, column_positions, find_header_row, normalize_code,
    squash, to_number,
)


def read_journal_rows(rows: list[tuple]) -> pd.DataFrame:
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
        date = cell(row, pos["전표일자"][0])
        if date is None or not str(date).strip():
            continue
        records.append({
            "date": date,
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
    df = pd.DataFrame(records, columns=JOURNAL_COLUMNS)
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    return df


def load_journal(path: str | Path) -> pd.DataFrame:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    return read_journal_rows(rows)
