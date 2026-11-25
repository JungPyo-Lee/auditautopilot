"""입증감사조서 파일 정리.

- 깨진 이름 정의(#REF!) 정리: 파일이 느려지는 주된 원인 중 하나.
- 절차서 탭(A, B, DER 처럼 숫자가 없는 대문자 탭) 숨김. 인쇄용 사본에서는 다시 보인다.
- 담당자를 고르면 총괄 시트 기준으로 내 인덱스 탭만 남긴다.
- 맨 앞에 '목차' 시트를 만들어 리드시트·세부조서로 바로 이동하게 한다.

원본 파일은 바꾸지 않고 새 파일로 저장한다.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from .xlsx_xml import XlsxPackage, build_sheet_xml, sheet_link

FRONT_SHEETS = ("일반사항", "총괄", "공시점검", "TB", "BS", "IS")
INDEX_SHEET = "목차"


@dataclass
class Assignment:
    index: str
    account: str
    assignee: str


@dataclass
class SheetRole:
    name: str
    kind: str          # front | procedure | lead | detail | other
    index: str | None


@dataclass
class PrepareReport:
    hidden: list[str] = field(default_factory=list)
    visible: list[str] = field(default_factory=list)
    names: dict[str, int] = field(default_factory=dict)
    index_sheet_added: bool = False
    assignee: str | None = None

    def lines(self) -> list[str]:
        out = [f"보이는 탭 {len(self.visible)}개, 숨긴 탭 {len(self.hidden)}개"]
        if self.assignee:
            out.append(f"담당자: {self.assignee}")
        if self.names:
            out.append(
                f"이름 정의 {self.names['before']:,}개 중 깨진 이름 {self.names['removed_ref']:,}개"
                f"{', 미사용 숨김 이름 %s개' % format(self.names['removed_unused'], ',') if self.names['removed_unused'] else ''}"
                f" 정리, 남은 이름 {self.names['after']:,}개"
            )
        if self.index_sheet_added:
            out.append("맨 앞에 '목차' 시트 추가")
        return out


def _text(v: object) -> str:
    return "" if v is None else str(v).strip()


def read_assignments(path: str | Path) -> list[Assignment]:
    """총괄 시트의 '인덱스 | 계정과목 | 담당자' 표(왼쪽·오른쪽 두 묶음)를 읽는다.
    담당자 칸이 수식이어도 엑셀이 저장해 둔 값을 읽는다."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if "총괄" not in wb.sheetnames:
            return []
        rows = [tuple(r) for r in wb["총괄"].iter_rows(values_only=True)]
    finally:
        wb.close()
    header_row = None
    blocks: list[tuple[int, int, int]] = []
    for i, row in enumerate(rows):
        cells = [_text(v) for v in row]
        starts = [j for j, v in enumerate(cells) if v == "인덱스"]
        if starts:
            header_row = i
            for j in starts:
                assignee_col = next((k for k in range(j + 1, min(j + 4, len(cells)))
                                     if cells[k] == "담당자"), j + 2)
                blocks.append((j, j + 1, assignee_col))
            break
    if header_row is None:
        return []
    result = []
    for row in rows[header_row + 1:]:
        for idx_col, acc_col, who_col in blocks:
            index = _text(row[idx_col]) if idx_col < len(row) else ""
            if not index:
                continue
            result.append(Assignment(
                index=index,
                account=_text(row[acc_col]) if acc_col < len(row) else "",
                assignee=_text(row[who_col]) if who_col < len(row) else "",
            ))
    return result


def classify_sheets(names: list[str], indices: list[str]) -> list[SheetRole]:
    """탭 이름으로 역할을 나눈다.

    - front: 일반사항, 총괄, TB 등 앞쪽 공통 탭
    - procedure: 숫자가 없는 탭 (A, B, DER, BB,DD) = 입증감사절차서 (인쇄용)
    - lead: 1000으로 끝나는 탭 (A1000) = 리드시트
    - detail: 그 밖에 숫자가 있는 탭 (A1100, E100)
    인덱스는 총괄 시트의 인덱스 중 탭 이름의 가장 긴 앞부분과 맞는 것.
    """
    by_length = sorted({i for i in indices if i}, key=len, reverse=True)
    roles = []
    for name in names:
        if name in FRONT_SHEETS or name == INDEX_SHEET:
            roles.append(SheetRole(name, "front", None))
            continue
        index = next((i for i in by_length if name.startswith(i)), None)
        has_digit = any(ch.isdigit() for ch in name)
        if not has_digit:
            kind = "procedure"
        elif name.endswith("1000"):
            kind = "lead"
        else:
            kind = "detail"
        if index is None:
            kind = "other" if has_digit else kind
        roles.append(SheetRole(name, kind, index))
    return roles


def _index_sheet_xml(roles: list[SheetRole], assignments: list[Assignment],
                     visible: set[str], company: str, assignee: str | None) -> bytes:
    title = f"목차 - {company}" if company else "목차"
    rows: list[list[str | None]] = [
        [title],
        [f"담당자: {assignee}" if assignee else "전체 보기", None, None,
         "탭 이름을 누르면 해당 시트로 이동합니다. 숨긴 절차서 탭은 인쇄용 사본에서 볼 수 있습니다."],
        [],
        ["인덱스", "계정과목", "담당자", "리드시트", "세부조서"],
    ]
    links: dict[tuple[int, int], str] = {}
    front = [r.name for r in roles if r.kind == "front" and r.name in visible]
    rows.append(["공통", "", "", *front[:2]])
    r_front = len(rows) - 1
    for c, name in enumerate(front[:2], start=3):
        links[(r_front, c)] = sheet_link(name)
    for extra in [front[i:i + 2] for i in range(2, len(front), 2)]:
        rows.append(["", "", "", *extra])
        for c, name in enumerate(extra, start=3):
            links[(len(rows) - 1, c)] = sheet_link(name)

    seen: set[str] = set()
    for a in assignments:
        if a.index in seen:
            continue
        seen.add(a.index)
        mine = [r for r in roles if r.index == a.index and r.name in visible and r.kind in ("lead", "detail")]
        if not mine:
            continue
        leads = [r.name for r in mine if r.kind == "lead"]
        details = [r.name for r in mine if r.kind == "detail"]
        first = [a.index, a.account, a.assignee, leads[0] if leads else "", details[0] if details else ""]
        rows.append(first)
        r0 = len(rows) - 1
        if leads:
            links[(r0, 3)] = sheet_link(leads[0])
        if details:
            links[(r0, 4)] = sheet_link(details[0])
        extra_leads, extra_details = leads[1:], details[1:]
        for k in range(max(len(extra_leads), len(extra_details))):
            lead = extra_leads[k] if k < len(extra_leads) else ""
            detail = extra_details[k] if k < len(extra_details) else ""
            rows.append(["", "", "", lead, detail])
            rk = len(rows) - 1
            if lead:
                links[(rk, 3)] = sheet_link(lead)
            if detail:
                links[(rk, 4)] = sheet_link(detail)
    return build_sheet_xml(rows, links, widths=[10, 34, 12, 14, 14])


def _company_name(path: Path) -> str:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if "일반사항" not in wb.sheetnames:
            return ""
        for row in wb["일반사항"].iter_rows(values_only=True):
            cells = [_text(v) for v in row]
            if "회사 명" in cells or "회사명" in cells:
                key = cells.index("회사 명") if "회사 명" in cells else cells.index("회사명")
                rest = [v for v in cells[key + 1:] if v]
                return rest[0] if rest else ""
    finally:
        wb.close()
    return ""


def prepare_workpaper(
    src: str | Path,
    dst: str | Path,
    assignee: str | None = None,
    hide_procedures: bool = True,
    add_index: bool = True,
    clean_names: bool = True,
    aggressive_names: bool = False,
) -> PrepareReport:
    src, dst = Path(src), Path(dst)
    if src.resolve() == dst.resolve():
        raise ValueError("원본과 다른 파일 이름으로 저장해야 합니다.")
    pkg = XlsxPackage(src)
    assignments = read_assignments(src)
    sheet_names = [s.name for s in pkg.sheets()]
    roles = classify_sheets(sheet_names, [a.index for a in assignments])

    my_indices: set[str] | None = None
    if assignee:
        my_indices = {a.index for a in assignments if a.assignee == assignee}
        if not my_indices:
            people = sorted({a.assignee for a in assignments if a.assignee})
            raise ValueError(f"총괄 시트에 '{assignee}' 담당 인덱스가 없습니다. 담당자 목록: {', '.join(people)}")

    report = PrepareReport(assignee=assignee)
    visible: set[str] = set()
    for role in roles:
        show = True
        if role.kind == "procedure" and hide_procedures:
            show = False
        if my_indices is not None and role.kind in ("procedure", "lead", "detail") and role.index not in my_indices:
            show = False
        pkg.set_state(role.name, "visible" if show else "hidden")
        (report.visible if show else report.hidden).append(role.name)
        if show:
            visible.add(role.name)

    if clean_names:
        report.names = pkg.clean_defined_names(aggressive=aggressive_names)

    if add_index and INDEX_SHEET not in sheet_names:
        listed = [a for a in assignments if my_indices is None or a.index in my_indices]
        xml = _index_sheet_xml(roles, listed, visible, _company_name(src), assignee)
        pkg.add_sheet_first(INDEX_SHEET, xml)
        report.index_sheet_added = True
        report.visible.insert(0, INDEX_SHEET)
    else:
        pkg.set_active_sheet(pkg.first_visible_index())
    pkg.save(dst)
    return report


def make_print_copy(src: str | Path, dst: str | Path, assignee: str | None = None) -> PrepareReport:
    """인쇄용 사본: 절차서 탭까지 모두 보이게 한다. 담당자를 주면 그 사람 인덱스만 남긴다."""
    return prepare_workpaper(src, dst, assignee=assignee, hide_procedures=False,
                             add_index=False, clean_names=True)
