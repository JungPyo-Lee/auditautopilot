"""결과를 엑셀 파일로 저장한다."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter

NUMBER_FORMAT = "#,##0;(#,##0);-"


def write_excel(path: str | Path, sheets: dict[str, pd.DataFrame]) -> Path:
    """여러 DataFrame을 시트별로 저장하고 숫자 서식과 열 너비를 맞춘다.
    '_'로 시작하는 내부용 열은 빼고 저장한다."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            out = df[[c for c in df.columns if not str(c).startswith("_")]]
            out.to_excel(writer, sheet_name=name[:31], index=False)
            ws = writer.sheets[name[:31]]
            for col_idx, col in enumerate(out.columns, start=1):
                letter = get_column_letter(col_idx)
                width = max([len(str(col))] + [len(str(v)) for v in out[col].head(200)]) + 2
                ws.column_dimensions[letter].width = min(max(width * 1.2, 8), 60)
                if pd.api.types.is_numeric_dtype(out[col]):
                    for cell in ws[letter][1:]:
                        cell.number_format = NUMBER_FORMAT
            ws.freeze_panes = "A2"
    return path
