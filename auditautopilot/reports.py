"""결과를 엑셀 파일로 저장한다."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

NUMBER_FORMAT = "#,##0;(#,##0);-"


def write_excel(path: str | Path, sheets: dict[str, list[dict]],
                headers: dict[str, dict[str, str]] | None = None) -> Path:
    """시트 이름 -> dict 목록을 저장한다.

    - headers[시트] 로 키를 한글 머리글로 바꿀 수 있다. 지정하지 않은 키는 그대로 쓴다.
    - '_'로 시작하는 내부용 키는 저장하지 않는다. 단 _reason 은 '추출사유'로 저장한다.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name[:31])
        keys: list[str] = []
        for r in rows:
            for k in r:
                if k not in keys and (not str(k).startswith("_") or k == "_reason"):
                    keys.append(k)
        rename = {"_reason": "추출사유", **((headers or {}).get(name, {}))}
        ws.append([rename.get(k, k) for k in keys])
        numeric = {k: True for k in keys}
        for r in rows:
            values = []
            for k in keys:
                v = r.get(k)
                if v is not None and not isinstance(v, (int, float)) or isinstance(v, bool):
                    numeric[k] = False
                values.append(v)
            ws.append(values)
        for i, k in enumerate(keys, start=1):
            letter = get_column_letter(i)
            sample = [str(rename.get(k, k))] + [str(r.get(k, "")) for r in rows[:200]]
            ws.column_dimensions[letter].width = min(max(max(len(s) for s in sample) * 1.2 + 2, 8), 60)
            if numeric[k] and rows:
                for c in ws[letter][1:]:
                    c.number_format = NUMBER_FORMAT
            elif rows and isinstance(rows[0].get(k), date):
                for c in ws[letter][1:]:
                    c.number_format = "yyyy-mm-dd"
        ws.freeze_panes = "A2"
    if not wb.sheetnames:
        wb.create_sheet("빈 결과")
    wb.save(path)
    return path
