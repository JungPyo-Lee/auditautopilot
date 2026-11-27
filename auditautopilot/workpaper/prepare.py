"""입증감사조서 파일 정리.

- 깨진 이름 정의(#REF!) 정리: 파일이 느려지는 주된 원인 중 하나.
- 절차서 탭(A, B, DER 처럼 숫자가 없는 대문자 탭) 숨김. 인쇄용 사본에서는 다시 보인다.
- 담당자를 고르면 총괄 시트(assign 탭이 있으면 그 V 표시) 기준으로 내 인덱스 탭만 남긴다.
- 맨 앞에 '목차' 시트를 만들어 리드시트·세부조서로 바로 이동하게 한다.

원본 파일은 바꾸지 않고 새 파일로 저장한다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from .xlsx_xml import XlsxPackage, build_sheet_xml, sheet_link

FRONT_SHEETS = ("일반사항", "총괄", "공시점검", "TB", "BS", "IS")
# 수정분개 반영으로 생기는 탭: 누구에게나 보인다
REPORT_SHEETS = ("AJE", "정산표", "수정후BS", "수정후IS", "정산표-CF", "보고서 BS", "보고서IS", "자본변동표",
                 "보고서현금흐름표")
INDEX_SHEET = "목차"
# assign 탭에서 줄여 쓰는 이름 -> 총괄 계정과목에 들어 있는 말
MATRIX_ALIAS = {"판관비": "판매비와관리비", "영업외손익": "기타이익,기타비용", "법인세등": "법인세"}


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


def _index_of(name: str, by_length: list[str]) -> str | None:
    """탭 이름이 인덱스로 시작하고 그 뒤가 비었거나 숫자·쉼표면 그 인덱스 (PL이 P 인덱스로 잡히지 않게)"""
    for i in by_length:
        rest = name[len(i):]
        if not name.startswith(i):
            continue
        if not rest or rest[0] == "," or (rest[0].isdigit() and not i[-1].isdigit()):
            return i        # E100 인덱스는 E1000 탭을 갖지 않는다 (E의 리드시트)
    return None


def classify_sheets(names: list[str], indices: list[str]) -> list[SheetRole]:
    """탭 이름으로 역할을 나눈다.

    - front: 일반사항, 총괄, TB 등 공통 탭. 첫 인덱스 탭보다 앞에 있는 탭과 정산표·보고서 탭도 공통이다.
    - procedure: 숫자가 없는 인덱스 탭 (A, B, DER, BB,DD) = 입증감사절차서 (인쇄용)
    - lead: 1000으로 끝나는 탭 (A1000) = 리드시트
    - detail: 그 밖에 숫자가 있는 탭 (A1100, E100), 그리고 인덱스 탭 사이에 끼운 보조 탭 (예: C1100 뒤의
      '금융평잔')은 바로 앞 인덱스의 세부조서로 본다.
    인덱스는 총괄 시트의 인덱스 중 탭 이름의 가장 긴 앞부분과 맞는 것.
    """
    by_length = sorted({i for i in indices if i}, key=len, reverse=True)
    roles = []
    current: str | None = None
    for name in names:
        if name in FRONT_SHEETS or name in REPORT_SHEETS or name == INDEX_SHEET:
            roles.append(SheetRole(name, "front", None))
            continue
        index = _index_of(name, by_length)
        has_digit = any(ch.isdigit() for ch in name)
        if index is None and current is not None:
            m = re.fullmatch(r"([A-Z]+(?:,[A-Z]+)*)\d*", name)
            if m:
                index = m.group(1)    # 총괄 표에 없는 인덱스 탭 (예: DER): 따로 묶어 담당자 조서에서는 숨긴다
        if index is None:
            if current is None:
                roles.append(SheetRole(name, "front", None))
            else:
                roles.append(SheetRole(name, "detail", current))
            continue
        current = index
        if not has_digit:
            kind = "procedure"
        elif name.endswith("1000"):
            kind = "lead"
        else:
            kind = "detail"
        roles.append(SheetRole(name, kind, index))
    return roles


def people(assignments: list[Assignment]) -> list[str]:
    """담당자 이름 (한 인덱스에 '홍길동, 김감사'처럼 여럿 적어도 된다)"""
    out: list[str] = []
    for a in assignments:
        for n in split_names(a.assignee):
            if n not in out:
                out.append(n)
    return out


def split_names(text: str) -> list[str]:
    return [n.strip() for n in re.split(r"[,/·\s]+", text or "") if n.strip()]


# 미리 배정 카테고리: 매출·매입 사이클(채권·채무·매출·원가·판관비)과 그 밖의 계정
CATEGORIES = ("매출및매입", "기타계정")
# 묶음 -> 총괄 계정과목에 들어 있으면 그 묶음 (위에서부터 먼저 맞는 것: '매출원가'는 원가, '매출채권'은 채권)
SALES_PURCHASE_GROUPS = (
    ("판관비", ("판매비", "판관비", "관리비")),
    ("원가", ("원가",)),
    ("채권", ("매출채권", "외상매출금", "받을어음", "공사미수금", "분양미수금", "미수금", "기타유동자산")),
    ("채무", ("매입채무", "외상매입금", "지급어음", "미지급금", "선수금", "기타유동부채")),
    ("매출", ("매출",)),
)
GROUP_NAMES = tuple(g for g, _ in SALES_PURCHASE_GROUPS)


def account_group(account: str) -> str | None:
    """총괄 계정과목 -> 매출및매입 묶음 이름(채권·채무·매출·원가·판관비). 해당 없으면 None.
    미수금은 기타유동자산(D), 미지급금·분양선수금은 기타유동부채(CC) 인덱스에 들어가므로 그 인덱스도 채권·채무로 본다."""
    t = re.sub(r"\s", "", account or "")
    for group, words in SALES_PURCHASE_GROUPS:
        if any(w in t for w in words):
            return group
    return None


def default_category(account: str) -> str:
    return CATEGORIES[0] if account_group(account) else CATEGORIES[1]


def reassign(assignments: list[Assignment], targets: set[str], names: list[str], mode: str = "set") -> list[Assignment]:
    """targets 인덱스의 담당자를 names로 바꾼다 (mode='add'면 기존 담당자에 덧붙인다)."""
    out = []
    for a in assignments:
        if a.index in targets:
            cur = split_names(a.assignee) if mode == "add" else []
            cur += [n for n in names if n not in cur]
            a = Assignment(a.index, a.account, ", ".join(cur))
        out.append(a)
    return out


def read_assign_matrix(path: str | Path, assignments: list[Assignment]) -> list[Assignment]:
    """'assign' 탭 (계정과목 줄 × 담당자 열, 맡은 칸에 V)을 읽어 총괄 인덱스에 담당자를 붙인다.
    계정과목 이름이 총괄 계정과목과 겹치는 인덱스에 붙이고, 못 맞춘 줄은 버린다. 탭이 없으면 그대로 돌려준다."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if "assign" not in wb.sheetnames:
            return assignments
        rows = [tuple(r) for r in wb["assign"].iter_rows(values_only=True)]
    finally:
        wb.close()
    head = next((i for i, r in enumerate(rows) if r and _text(r[0]) == "계정과목"), None)
    if head is None:
        return assignments
    who = {j: _text(v) for j, v in enumerate(rows[head]) if j > 0 and _text(v)}

    def norm(t: str) -> str:
        t = re.sub(r"\(.*?\)|\s", "", t)
        return MATRIX_ALIAS.get(t, t)
    names: dict[str, list[str]] = {}
    for r in rows[head + 1:]:
        area = norm(_text(r[0]) if r else "")
        if not area:
            continue
        marked = [n for j, n in who.items() if j < len(r) and _text(r[j]).upper() in ("V", "O", "✓", "1")]
        if not marked:
            continue
        for a in assignments:
            acc = norm(a.account)
            parts = [p for p in re.split(r"[,및]", acc) if len(p) >= 2]
            if area in acc or acc in area or any(p in area for p in parts):
                lst = names.setdefault(a.index, [])
                lst += [n for n in marked if n not in lst]
    return [Assignment(a.index, a.account, ", ".join(names[a.index]) if a.index in names else a.assignee)
            for a in assignments]


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
    assignments: list[Assignment] | None = None,
) -> PrepareReport:
    """assignments를 주면 총괄 시트 대신 그 배정을 쓴다 (관리 화면에서 고친 담당자)."""
    src, dst = Path(src), Path(dst)
    if src.resolve() == dst.resolve():
        raise ValueError("원본과 다른 파일 이름으로 저장해야 합니다.")
    pkg = XlsxPackage(src)
    if assignments is None:
        assignments = read_assign_matrix(src, read_assignments(src))
    sheet_names = [s.name for s in pkg.sheets()]
    roles = classify_sheets(sheet_names, [a.index for a in assignments])

    my_indices: set[str] | None = None
    if assignee:
        my_indices = {a.index for a in assignments if assignee in split_names(a.assignee)}
        if not my_indices:
            raise ValueError(f"'{assignee}' 담당 인덱스가 없습니다. 담당자 목록: {', '.join(people(assignments))}")

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

    if add_index and INDEX_SHEET in sheet_names:      # 예전 목차는 다시 만든다 (담당자에 맞게)
        pkg.remove_sheet(INDEX_SHEET)
        roles = [r for r in roles if r.name != INDEX_SHEET]
        sheet_names = [n for n in sheet_names if n != INDEX_SHEET]
    if add_index:
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
