"""회계 프로그램 내보내기 파일을 읽을 때 공통으로 쓰는 도구.

모든 불러오기 함수는 아래 표준 열 이름을 키로 하는 dict 목록(list[dict])을 돌려준다.
윈도우 앱 제어 정책에 막히지 않도록 pandas·numpy 같은 컴파일된 패키지는 쓰지 않는다.
"""

from __future__ import annotations

import re
from datetime import date, datetime
from typing import Iterable, Sequence

# 표준 열 이름
TB_COLUMNS = [
    "code", "name", "level",
    "debit_balance", "debit_total", "credit_total", "credit_balance", "balance",
]
FS_COLUMNS = ["code", "name", "is_account", "current", "prior"]
JOURNAL_COLUMNS = [
    "date", "voucher_no", "side", "account_code", "account_name",
    "debit", "credit", "partner_code", "partner", "memo",
]
PARTNER_COLUMNS = [
    "account_code", "partner_code", "partner", "reg_no",
    "opening", "debit", "credit", "balance",
]

_CODE_RE = re.compile(r"^\s*\[(\d{3,5})\]\s*(.+?)\s*$")
_RRN_RE = re.compile(r"^(\d{6})-?([1-8])\d{6}$")


def squash(value: object) -> str:
    """'보    통    예    금' 처럼 자간을 띄운 과목명을 '보통예금'으로 만든다."""
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value))


def normalize_code(code: object) -> str | None:
    """계정코드를 비교 가능한 형태로 맞춘다. '0103' -> '103', '253' -> '253'."""
    if code is None:
        return None
    text = str(code).strip()
    if not text:
        return None
    if text.isdigit():
        return str(int(text)).zfill(3)
    return text


def parse_account_label(label: object) -> tuple[str | None, str]:
    """'[0103]   보 통 예 금' -> ('103', '보통예금'). 코드가 없으면 (None, 과목명)."""
    match = _CODE_RE.match(str(label or ""))
    if match:
        return normalize_code(match.group(1)), squash(match.group(2))
    return None, squash(label)


def to_number(value: object) -> float:
    """빈 칸, 쉼표, 괄호 음수를 숫자로 바꾼다."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text or text == "-":
        return 0.0
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    try:
        number = float(text)
    except ValueError:
        return 0.0
    return -number if negative else number


def mask_reg_no(value: object) -> str:
    """주민등록번호 형태는 뒷자리를 가린다. 사업자등록번호는 그대로 둔다."""
    if value is None:
        return ""
    text = str(value).strip()
    match = _RRN_RE.match(text)
    if match:
        return f"{match.group(1)}-{match.group(2)}******"
    return text


def find_header_row(rows: Sequence[Sequence[object]], required: Iterable[str]) -> int:
    """필수 머리글이 모두 들어 있는 첫 행 번호(0부터)를 찾는다."""
    needed = {squash(r) for r in required}
    for i, row in enumerate(rows):
        present = {squash(v) for v in row if v is not None}
        if needed <= present:
            return i
    raise ValueError(f"머리글 행을 찾지 못했습니다: {sorted(needed)}")


def column_positions(header: Sequence[object]) -> dict[str, list[int]]:
    """머리글 이름 -> 열 위치 목록. 같은 이름이 여러 번 나오면 순서대로 모두 담는다."""
    positions: dict[str, list[int]] = {}
    for i, value in enumerate(header):
        if value is None:
            continue
        positions.setdefault(squash(value), []).append(i)
    return positions


def cell(row: Sequence[object], index: int | None) -> object:
    if index is None or index >= len(row):
        return None
    return row[index]


def parse_date(value: object) -> date | None:
    """'2026-01-05', '2026.01.05', '20260105', datetime 값을 date로 바꾼다. 못 읽으면 None."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None
