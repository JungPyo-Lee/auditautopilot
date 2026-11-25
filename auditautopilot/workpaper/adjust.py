"""수정분개를 조서에 반영한다: 수정분개 양식 → AJE·정산표·수정후BS·수정후IS·현금흐름정산표 탭.

흐름
1. aje_template(조서) : 조서 TB 탭 계정과목으로 수정분개 양식(.xlsx)을 만든다 (계정과목은 목록에서 고름).
2. apply(조서, 수정분개, 출력) : 리드시트까지 채운 조서에 탭을 더한다.
   - AJE        : 올린 수정분개 그대로 (번호·일자·구분·계정과목·차변·대변·적요·근거)
   - 정산표     : TB 탭 잔액(수정전) + AJE(SUMIFS) = 수정후, 손익계산서·재무상태표 칸으로 나누고 당기순이익으로 대차를 맞춘다
   - 수정후BS·수정후IS : 정산표를 가리키는 수식으로 묶음별 소계·총계, 전기는 BS·IS 탭에서
   - 현금흐름정산표 : 간접법. 당기순이익 + 현금 없는 손익 + 재무상태표 계정별 증감(당기-전기)을 영업·투자·재무로 나누고
                     기초·기말 현금과 맞는지 검증 칸을 둔다
   - 리드시트(○1000) I열(Audited F/S) = H열(G/L) + 정산표의 수정분개 금액
   모든 숫자는 수식으로 이어져 회계사가 셀을 눌러 따라가며 검증할 수 있다.

값을 넣는 것은 조서 탭뿐이고 양식(도형·서식)은 그대로 둔다 (xlsx_xml.py).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from ..importers.common import squash, to_number
from .leadsheet import BS_SIDES, TARGETS, is_contra, read_template
from .xlsx_xml import Formula, open_workbook, XlsxPackage, column_letter, grid_xml

NEW_TABS = ("AJE", "정산표", "수정후BS", "수정후IS", "현금흐름정산표")
AJE_HEADERS = ["번호", "일자", "구분", "계정과목", "차변", "대변", "적요", "근거 조서", "새 계정 묶음"]
AJE_KINDS = ("수정", "재분류")

BS_MAJORS = ("유동자산", "비유동자산", "유동부채", "비유동부채", "자본")
IS_ORDER = [("sales", "Ⅰ. 매출액"), ("cogs", "Ⅱ. 매출원가"), ("sga", "Ⅳ. 판매비와관리비"),
            ("nonop_inc", "Ⅵ. 영업외수익"), ("nonop_exp", "Ⅶ. 영업외비용"), ("tax", "Ⅸ. 법인세비용")]
REVENUE_KEYS = {"sales", "nonop_inc"}
MFG_KEYS = {"material", "labor", "overhead"}   # 결산 전 시산표의 제조원가 계정은 매출원가에 넣어 보여 준다

# 현금 유출이 없는 비용·유입이 없는 수익 -> 짝이 되는 재무상태표 계정 (이름에 들어 있는 말)
NONCASH = [
    ("감가상각비", ("감가상각누계액",), +1),
    ("무형자산상각비", ("영업권", "개발비", "산업재산권", "소프트웨어", "무형자산"), +1),
    ("대손상각비", ("대손충당금",), +1),
    ("퇴직급여", ("퇴직급여충당부채", "퇴직급여충당금"), +1),
    ("유형자산처분손실", None, +1),
    ("유형자산처분이익", None, -1),
]
CASH_EXCLUDE = ("정기", "적금", "금융상품", "사용제한", "특정", "장기")

# 시산표에 없는 계정을 수정분개로 새로 쓸 때 묶음을 정하지 않았으면 이름으로 짐작한다
NEW_ACCOUNT_GUESS = [
    (("대손상각비", "감가상각비", "무형자산상각비", "퇴직급여", "지급수수료", "급여", "복리후생비"), "sga"),
    (("이자비용", "외화환산손실", "외환차손", "잡손실", "기타의대손상각비", "처분손실", "손상차손", "기부금"), "nonop_exp"),
    (("이자수익", "외화환산이익", "외환차익", "잡이익", "처분이익", "배당금수익", "환입"), "nonop_inc"),
    (("법인세비용", "법인세등"), "tax"),
    (("미지급법인세", "미지급비용", "선수수익", "예수금", "미지급금"), "other_cl"),
    (("미수수익", "선급비용", "선급법인세", "미수금", "선급금"), "other_quick"),
    (("매출원가",), "cogs"),
]


@dataclass
class TbRow:
    tb_row: int
    name: str
    key: str
    side: str            # asset | liab | is
    major: str
    group_row: int
    group_label: str
    contra: bool
    db: float
    cb: float
    bs_row: int | None = None
    label: str = ""      # 겹치지 않는 이름 (차감계정은 '감가상각누계액(건물)')
    new: bool = False    # 시산표에 없고 수정분개로 새로 생긴 계정 (tb_row는 음수 번호)


@dataclass
class AjeLine:
    no: str
    when: object
    kind: str
    account: str
    debit: float
    credit: float
    memo: str
    ref: str
    group: str = ""


@dataclass
class Result:
    output: Path | None = None
    entries: int = 0
    lines: int = 0
    debit: float = 0.0
    credit: float = 0.0
    notes: list[str] = field(default_factory=list)
    problems: list[str] = field(default_factory=list)

    def summary(self) -> list[str]:
        out = [f"수정분개 {self.entries}건({self.lines}줄), 차변 {self.debit:,.0f}원 · 대변 {self.credit:,.0f}원을 "
               f"정산표에 반영했습니다 ({', '.join(NEW_TABS)} 탭)."]
        out += [f"확인 필요: {p}" for p in self.problems]
        out += [f"참고: {n}" for n in self.notes]
        return out


# ---------------------------------------------------------------- 조서 읽기

def _group_key(major: str, label: str) -> str | None:
    return next((k for k, (maj, pre) in TARGETS.items() if maj == major and label.startswith(pre)), None)


def read_rows(workpaper: str | Path) -> tuple[object, list[TbRow]]:
    """리드시트까지 채운 조서의 TB 탭에서 계정 줄을 읽는다."""
    tpl = read_template(workpaper)
    wb = open_workbook(workpaper, data_only=False)
    try:
        tb = wb["TB"]
        out = []
        for g in tpl.groups:
            key = _group_key(g.major, g.label)
            if key is None:
                continue
            side = BS_SIDES.get(g.major, "is")
            for r in g.slots:
                name = tb.cell(r, 6).value
                if not isinstance(name, str) or not name.strip() or name.startswith("="):
                    continue
                db, cb = to_number(tb.cell(r, 4).value), to_number(tb.cell(r, 8).value)
                out.append(TbRow(r, squash(name), key, side, g.major, g.row, g.label, is_contra(name), db, cb,
                                 tpl.bs_row.get(r)))
        _set_labels(out)
        if not out:
            raise ValueError("조서 TB 탭에 계정이 없습니다. 먼저 [조서 채우기]로 시산표를 넣은 조서를 올려 주세요.")
        return tpl, out
    finally:
        wb.close()


def _set_labels(rows: list[TbRow]) -> None:
    parent = None
    for x in rows:
        x.label = f"{x.name}({parent})" if x.contra and parent else x.name
        if not x.contra:
            parent = x.name
    seen: dict[str, int] = {}
    for x in rows:
        seen[x.label] = seen.get(x.label, 0) + 1
    for x in rows:
        if seen[x.label] > 1:
            x.label = f"{x.label}[{x.group_label}]"


def group_choices(tpl) -> dict[str, object]:
    """'유동부채 > 기타유동부채' -> 묶음"""
    out = {}
    for g in tpl.groups:
        if _group_key(g.major, g.label):
            out.setdefault(f"{g.major} > {g.label}", g)
    return out


# ---------------------------------------------------------------- 수정분개 양식

def aje_template(workpaper: str | Path, output: str | Path) -> Path:
    tpl, rows = read_rows(workpaper)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "수정분개"
    ws.append(["수정분개 (AJE)"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append(["번호가 같은 줄이 한 분개입니다. 계정과목은 목록(조서 TB 탭 과목)에서 고르세요. 구분: 수정 / 재분류. "
               "시산표에 없는 계정(예: 대손상각비)을 새로 쓰면 계정과목을 직접 적고 '새 계정 묶음'을 고르세요."])
    ws.append([])
    ws.append(AJE_HEADERS)
    head = ws.max_row
    for c in ws[head]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDE6F4")
        c.alignment = Alignment(horizontal="center")
    last = head + 300
    names = wb.create_sheet("계정목록")
    names.append(["계정과목", "구분", "조서"])
    for r in rows:
        names.append([r.label, {"asset": "자산", "liab": "부채·자본"}.get(r.side, "손익"), r.group_label])
    choices = list(group_choices(tpl))
    for i, c in enumerate(choices, start=2):
        names.cell(i, 5, c)
    names.cell(1, 5, "새 계정 묶음")
    names.column_dimensions["A"].width = 30
    names.column_dimensions["E"].width = 40
    dv = DataValidation(type="list", formula1=f"=계정목록!$A$2:$A${len(rows) + 1}", allow_blank=True,
                        showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add(f"D{head + 1}:D{last}")
    grp = DataValidation(type="list", formula1=f"=계정목록!$E$2:$E${len(choices) + 1}", allow_blank=True)
    ws.add_data_validation(grp)
    grp.add(f"I{head + 1}:I{last}")
    kinds = DataValidation(type="list", formula1='"' + ",".join(AJE_KINDS) + '"', allow_blank=True)
    ws.add_data_validation(kinds)
    kinds.add(f"C{head + 1}:C{last}")
    for r in range(head + 1, last + 1):
        ws.cell(r, 5).number_format = ws.cell(r, 6).number_format = "#,##0"
        ws.cell(r, 2).number_format = "yyyy-mm-dd"
    ws.cell(last + 2, 4, "합계").font = Font(bold=True)
    ws.cell(last + 2, 5, f"=SUM(E{head + 1}:E{last})").number_format = "#,##0"
    ws.cell(last + 2, 6, f"=SUM(F{head + 1}:F{last})").number_format = "#,##0"
    ws.cell(last + 2, 7, f'=IF(E{last + 2}=F{last + 2},"대차 일치","대차 불일치")')
    for col, w in zip("ABCDEFGHI", (7, 12, 8, 30, 16, 16, 40, 12, 30)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = ws.cell(head + 1, 1)
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return output


def read_aje(path: str | Path) -> list[AjeLine]:
    wb = open_workbook(path, read_only=True, data_only=True)
    try:
        ws = next((wb[n] for n in ("수정분개", "AJE") if n in wb.sheetnames), wb.worksheets[0])
        rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    head = next((i for i, r in enumerate(rows[:20]) if {"계정과목", "차변", "대변"} <= {squash(v) for v in r if v}),
                None)
    if head is None:
        raise ValueError("수정분개 파일에서 '계정과목·차변·대변' 머리글을 못 찾았습니다. 수정분개 양식을 써 주세요.")
    pos = {squash(v): i for i, v in enumerate(rows[head]) if v}

    def get(r, k):
        i = pos.get(k)
        return r[i] if i is not None and i < len(r) else None
    out, no = [], ""
    for r in rows[head + 1:]:
        acc = get(r, "계정과목")
        if not acc or squash(acc) in ("합계", "계"):
            continue
        dr, cr = to_number(get(r, "차변")), to_number(get(r, "대변"))
        if not dr and not cr:
            continue
        no = str(get(r, "번호") or no).strip()
        if no.endswith(".0"):
            no = no[:-2]
        out.append(AjeLine(no, get(r, "일자"), str(get(r, "구분") or "수정").strip(), squash(acc), dr, cr,
                           str(get(r, "적요") or "").strip(), str(get(r, "근거조서") or "").strip(),
                           str(get(r, "새계정묶음") or "").strip()))
    return out


def add_new_accounts(tpl, rows: list[TbRow], lines: list[AjeLine]) -> tuple[list[TbRow], list[str], list[str]]:
    """수정분개에만 있는 계정을 정산표 줄로 더한다. 반환: (줄 목록, 확인 필요, 참고)"""
    problems, notes = [], []
    known = {x.label for x in rows}
    choices = group_choices(tpl)
    by_key = {}
    for label, g in choices.items():
        by_key.setdefault(_group_key(g.major, g.label), g)
    out = list(rows)
    n = 0
    for acc in dict.fromkeys(ln.account for ln in lines):
        if acc in known:
            continue
        plain = [x for x in rows if x.name == acc]
        if len(plain) == 1:
            for ln in lines:
                if ln.account == acc:
                    ln.account = plain[0].label
            continue
        if len(plain) > 1:
            problems.append(f"'{acc}'은(는) 조서에 여러 개라 어느 것인지 모릅니다. 목록의 이름({plain[0].label} 등)으로 고쳐 주세요")
            continue
        want = next((ln.group for ln in lines if ln.account == acc and ln.group), "")
        g = choices.get(want)
        if g is None:
            key = next((k for words, k in NEW_ACCOUNT_GUESS if any(w in acc for w in words)), None)
            g = by_key.get(key)
            if g is None:
                problems.append(f"'{acc}'은(는) 조서 TB 탭에 없는 과목입니다. '새 계정 묶음'을 골라 주세요 (정산표에 안 들어감)")
                continue
            notes.append(f"새 계정 '{acc}'을(를) '{g.major} > {g.label}' 묶음에 넣었습니다 (이름으로 짐작)")
        else:
            notes.append(f"새 계정 '{acc}'을(를) '{want}' 묶음에 넣었습니다")
        n += 1
        x = TbRow(-n, acc, _group_key(g.major, g.label), BS_SIDES.get(g.major, "is"), g.major, g.row, g.label,
                  is_contra(acc), 0.0, 0.0, None, acc, True)
        last = max((i for i, y in enumerate(out) if y.group_row == g.row), default=len(out) - 1)
        out.insert(last + 1, x)
        known.add(acc)
    if n:
        notes.append(f"새 계정 {n}개는 TB 탭·리드시트에는 없고 정산표·수정후BS·IS에만 나옵니다")
    return out, problems, notes


def check_aje(lines: list[AjeLine], rows: list[TbRow]) -> list[str]:
    problems = []
    by_no: dict[str, list[AjeLine]] = {}
    for ln in lines:
        by_no.setdefault(ln.no, []).append(ln)
    for no, ls in by_no.items():
        d, c = sum(x.debit for x in ls), sum(x.credit for x in ls)
        if abs(d - c) >= 1:
            problems.append(f"수정분개 {no or '(번호 없음)'}: 차변 {d:,.0f}원 ≠ 대변 {c:,.0f}원")
    return problems


# ---------------------------------------------------------------- 시트 만들기 도우미

class Grid:
    def __init__(self, styles: dict[str, str]):
        self.cells: dict[str, tuple] = {}
        self.st = styles

    def put(self, ref: str, value, style: str = "text"):
        self.cells[ref] = (value, self.st.get(style))

    def row(self, r: int, values: list, style: str = "text", start: int = 0):
        for i, v in enumerate(values):
            if v is not None:
                self.put(f"{column_letter(start + i)}{r}", v, style)


def _q(sheet: str) -> str:
    return "'" + sheet.replace("'", "''") + "'"


def _aje_sheet(st, lines: list[AjeLine]) -> tuple[bytes, int]:
    g = Grid(st)
    g.put("A1", "수정분개 (AJE)", "title")
    g.put("A2", "수정분개 양식에서 올린 내용 그대로입니다. 정산표 F·G열이 이 탭의 계정과목별 합계(SUMIFS)입니다.")
    g.row(4, AJE_HEADERS, "head")
    r = 5
    for ln in lines:
        when = ln.when.date() if isinstance(ln.when, datetime) else ln.when
        g.row(r, [ln.no, when.isoformat() if isinstance(when, date) else (str(when) if when else None), ln.kind,
                  ln.account], "text")
        g.put(f"E{r}", ln.debit or None, "num")
        g.put(f"F{r}", ln.credit or None, "num")
        g.row(r, [ln.memo or None, ln.ref or None, ln.group or None], "text", start=6)
        r += 1
    last = max(r - 1, 5)
    g.put(f"D{r + 1}", "합계", "bold")
    g.put(f"E{r + 1}", Formula(f"SUM(E5:E{last})"), "boldnum")
    g.put(f"F{r + 1}", Formula(f"SUM(F5:F{last})"), "boldnum")
    g.put(f"G{r + 1}", Formula(f'IF(ROUND(E{r + 1}-F{r + 1},0)=0,"대차 일치","대차 불일치")'), "check")
    return grid_xml(g.cells, [7, 12, 8, 30, 16, 16, 40, 12, 24], freeze="A5"), r + 1


# 정산표 열
WS_COLS = ["구분", "조서", "계정과목", "수정전 차변", "수정전 대변", "수정분개 차변", "수정분개 대변",
           "수정후 차변", "수정후 대변", "손익계산서 차변", "손익계산서 대변", "재무상태표 차변", "재무상태표 대변",
           "수정분개(리드시트 부호)"]


def _worksheet(st, rows: list[TbRow], aje_total_row: int) -> tuple[bytes, dict[int, int], dict[str, int]]:
    """반환: (시트, TB 행 -> 정산표 행, 합계·당기순이익 행 위치)"""
    g = Grid(st)
    ws = "정산표"
    g.put("A1", "정산표", "title")
    g.put("A2", "수정전 = TB 탭 잔액, 수정분개 = AJE 탭 합계(SUMIFS), 수정후 = 수정전 + 수정분개. "
                "손익계산서 칸의 차이가 당기순이익이고, 재무상태표 칸은 당기순이익을 더해 대차가 맞습니다.")
    g.row(4, WS_COLS, "head")
    pos: dict[int, int] = {}
    r = 5
    aje = _q("AJE")
    for x in rows:
        kind = {"asset": "BS-자산", "liab": "BS-부채·자본"}.get(x.side, "IS")
        g.row(r, [kind, None], "text")
        g.put(f"C{r}", x.label, "text")
        if x.new:
            g.put(f"D{r}", 0, "num")
            g.put(f"E{r}", 0, "num")
        else:
            g.put(f"D{r}", Formula(f"N(TB!D{x.tb_row})"), "num")
            g.put(f"E{r}", Formula(f"N(TB!H{x.tb_row})"), "num")
        g.put(f"F{r}", Formula(f"SUMIFS({aje}!$E:$E,{aje}!$D:$D,$C{r})"), "input")
        g.put(f"G{r}", Formula(f"SUMIFS({aje}!$F:$F,{aje}!$D:$D,$C{r})"), "input")
        g.put(f"H{r}", Formula(f"MAX(D{r}-E{r}+F{r}-G{r},0)"), "num")
        g.put(f"I{r}", Formula(f"MAX(E{r}-D{r}+G{r}-F{r},0)"), "num")
        g.put(f"J{r}", Formula(f'IF($A{r}="IS",H{r},0)'), "num")
        g.put(f"K{r}", Formula(f'IF($A{r}="IS",I{r},0)'), "num")
        g.put(f"L{r}", Formula(f'IF($A{r}="IS",0,H{r})'), "num")
        g.put(f"M{r}", Formula(f'IF($A{r}="IS",0,I{r})'), "num")
        credit_view = x.side == "liab" or (x.side == "is" and x.key in REVENUE_KEYS)
        g.put(f"N{r}", Formula(f"G{r}-F{r}" if credit_view else f"F{r}-G{r}"), "num")
        g.put(f"B{r}", x.group_label + (" (신규)" if x.new else ""), "text")
        pos[x.tb_row] = r
        r += 1
    first, last = 5, r - 1
    tot, ni, grand = r + 1, r + 2, r + 3
    g.put(f"C{tot}", "합계", "bold")
    for col in "DEFGHIJKLM":
        g.put(f"{col}{tot}", Formula(f"SUM({col}{first}:{col}{last})"), "boldnum")
    g.put(f"C{ni}", "당기순이익(손실)", "bold")
    g.put(f"J{ni}", Formula(f"K{tot}-J{tot}"), "boldnum")
    g.put(f"M{ni}", Formula(f"J{ni}"), "boldnum")
    g.put(f"C{grand}", "합계(당기순이익 포함)", "bold")
    g.put(f"J{grand}", Formula(f"J{tot}+J{ni}"), "boldnum")
    g.put(f"K{grand}", Formula(f"K{tot}"), "boldnum")
    g.put(f"L{grand}", Formula(f"L{tot}"), "boldnum")
    g.put(f"M{grand}", Formula(f"M{tot}+M{ni}"), "boldnum")
    chk = grand + 2
    g.put(f"C{chk}", "검증 (모두 0이어야 함)", "bold")
    checks = [("수정전 차변-대변", f"D{tot}-E{tot}"), ("수정분개 차변-대변", f"F{tot}-G{tot}"),
              ("AJE 탭 차변 - 정산표 반영액 (TB에 없는 과목)", f"{aje}!E{aje_total_row}-F{tot}"),
              ("수정후 차변-대변", f"H{tot}-I{tot}"), ("재무상태표 차변-대변", f"L{grand}-M{grand}")]
    for i, (label, f) in enumerate(checks, start=1):
        g.put(f"C{chk + i}", label)
        g.put(f"D{chk + i}", Formula(f"ROUND({f},0)"), "check")
    xml = grid_xml(g.cells, [11, 7, 26] + [15] * 11, freeze="D5")
    return xml, pos, {"total": tot, "ni": ni, "grand": grand, "first": first, "last": last, "sheet": ws}


def _bs_sheet(st, rows: list[TbRow], wpos: dict[int, int], w: dict) -> tuple[bytes, dict]:
    """수정후 재무상태표. 반환: 위치 정보 (TB 행 -> 행, 총계 행 등)"""
    g = Grid(st)
    ws = _q("정산표")
    g.put("A1", "수정후 재무상태표", "title")
    g.put("A2", "당기 = 정산표 재무상태표 칸 (자산 차변-대변, 부채·자본 대변-차변). 전기 = BS 탭 전기 금액. "
                "결산 전 시산표면 당기순이익을 이익잉여금 아래 줄로 따로 더합니다.")
    g.row(4, ["과목", "당기(수정후)", "전기", "증감"], "head")
    r = 5
    acc_row: dict[int, int] = {}
    major_total: dict[str, int] = {}
    group_total: dict[int, int] = {}
    retained_rows: list[int] = []
    ni_row = None
    for major in BS_MAJORS:
        items = [x for x in rows if x.major == major]
        if not items:
            continue
        g.put(f"A{r}", major, "bold")
        head = r
        r += 1
        subs = []
        for grow in dict.fromkeys(x.group_row for x in items):
            gi = [x for x in items if x.group_row == grow]
            g.put(f"A{r}", f"  {gi[0].group_label}", "bold")
            gh = r
            r += 1
            for x in gi:
                p = wpos[x.tb_row]
                cur = f"{ws}!L{p}-{ws}!M{p}" if x.side == "asset" else f"{ws}!M{p}-{ws}!L{p}"
                g.put(f"A{r}", f"    {x.label}")
                g.put(f"B{r}", Formula(cur), "num")
                g.put(f"C{r}", Formula(f"N(BS!F{x.bs_row})") if x.bs_row else None, "num")
                g.put(f"D{r}", Formula(f"B{r}-C{r}"), "num")
                acc_row[x.tb_row] = r
                r += 1
            if gi[0].key == "retained":
                g.put(f"A{r}", "    당기순이익(손실) (정산표)")
                g.put(f"B{r}", Formula(f"{ws}!J{w['ni']}"), "num")
                g.put(f"D{r}", Formula(f"B{r}-C{r}"), "num")
                ni_row = r
                r += 1
                retained_rows.append(gh)
            for col in "BCD":
                g.put(f"{col}{gh}", Formula(f"SUM({col}{gh + 1}:{col}{r - 1})"), "boldnum")
            group_total[grow] = gh
            subs.append(gh)
        for col in "BCD":
            g.put(f"{col}{head}", Formula("+".join(f"{col}{s}" for s in subs)), "boldnum")
        major_total[major] = head
    r += 1
    totals = {}

    def total(label, majors):
        nonlocal r
        have = [major_total[m] for m in majors if m in major_total]
        g.put(f"A{r}", label, "bold")
        for col in "BCD":
            g.put(f"{col}{r}", Formula("+".join(f"{col}{h}" for h in have) or "0"), "boldnum")
        totals[label] = r
        r += 1
    total("자산총계", ("유동자산", "비유동자산"))
    total("부채총계", ("유동부채", "비유동부채"))
    total("자본총계", ("자본",))
    g.put(f"A{r}", "부채와자본총계", "bold")
    for col in "BCD":
        g.put(f"{col}{r}", Formula(f"{col}{totals['부채총계']}+{col}{totals['자본총계']}"), "boldnum")
    totals["부채와자본총계"] = r
    r += 2
    g.put(f"A{r}", "검증: 자산총계 - 부채와자본총계 (0이어야 함)", "bold")
    for col in "BC":
        g.put(f"{col}{r}", Formula(f"ROUND({col}{totals['자산총계']}-{col}{totals['부채와자본총계']},0)"), "check")
    if ni_row is None:
        g.put(f"A{r + 1}", "이익잉여금 묶음이 없어 당기순이익 줄을 넣지 못했습니다. 대차가 맞지 않으면 확인해 주세요.")
    return grid_xml(g.cells, [34, 18, 18, 18], freeze="B5"), {
        "acc": acc_row, "group": group_total, "major": major_total, "totals": totals, "ni": ni_row,
        "retained": retained_rows}


def _is_sheet(st, rows: list[TbRow], wpos: dict[int, int], w: dict) -> tuple[bytes, dict]:
    g = Grid(st)
    ws = _q("정산표")
    g.put("A1", "수정후 손익계산서", "title")
    g.put("A2", "당기 = 정산표 손익계산서 칸 (수익 대변-차변, 비용 차변-대변). 전기 = IS 탭에서 같은 과목명(SUMIF). "
                "결산 전 제조원가 계정은 매출원가 안에 따로 보여 줍니다.")
    g.row(4, ["과목", "당기(수정후)", "전기", "증감"], "head")
    r = 5
    sec: dict[str, int] = {}
    acc_row: list[tuple[str, int]] = []
    for key, label in IS_ORDER:
        keys = {key} | (MFG_KEYS if key == "cogs" else set())
        items = [x for x in rows if x.side == "is" and x.key in keys]
        g.put(f"A{r}", label, "bold")
        head = r
        r += 1
        for x in items:
            p = wpos[x.tb_row]
            cur = f"{ws}!K{p}-{ws}!J{p}" if x.key in REVENUE_KEYS else f"{ws}!J{p}-{ws}!K{p}"
            tag = " (제조원가)" if x.key in MFG_KEYS else ""
            g.put(f"A{r}", f"    {x.label}{tag}")
            g.put(f"B{r}", Formula(cur), "num")
            g.put(f"C{r}", Formula(f'SUMIF(IS!$C:$C,"{x.name}",IS!$F:$F)'), "num")
            g.put(f"D{r}", Formula(f"B{r}-C{r}"), "num")
            acc_row.append((x.name, r))
            r += 1
        for col in "BCD":
            g.put(f"{col}{head}", Formula(f"SUM({col}{head + 1}:{col}{r - 1})") if r - 1 > head else 0, "boldnum")
        sec[key] = head
        if key == "cogs":
            g.put(f"A{r}", "Ⅲ. 매출총이익", "bold")
            for col in "BCD":
                g.put(f"{col}{r}", Formula(f"{col}{sec['sales']}-{col}{sec['cogs']}"), "boldnum")
            sec["gross"] = r
            r += 1
        if key == "sga":
            g.put(f"A{r}", "Ⅴ. 영업이익(손실)", "bold")
            for col in "BCD":
                g.put(f"{col}{r}", Formula(f"{col}{sec['gross']}-{col}{sec['sga']}"), "boldnum")
            sec["op"] = r
            r += 1
        if key == "nonop_exp":
            g.put(f"A{r}", "Ⅷ. 법인세비용차감전순이익(손실)", "bold")
            for col in "BCD":
                g.put(f"{col}{r}", Formula(f"{col}{sec['op']}+{col}{sec['nonop_inc']}-{col}{sec['nonop_exp']}"),
                      "boldnum")
            sec["ebt"] = r
            r += 1
    g.put(f"A{r}", "Ⅹ. 당기순이익(손실)", "bold")
    for col in "BCD":
        g.put(f"{col}{r}", Formula(f"{col}{sec['ebt']}-{col}{sec['tax']}"), "boldnum")
    sec["ni"] = r
    other = [x for x in rows if x.side == "is" and x.key not in {k for k, _ in IS_ORDER} | MFG_KEYS]
    r += 2
    g.put(f"A{r}", "검증: 당기순이익 - 정산표 당기순이익 (0이어야 함)", "bold")
    g.put(f"B{r}", Formula(f"ROUND(B{sec['ni']}-{ws}!J{w['ni']},0)"), "check")
    if other:
        g.put(f"A{r + 1}", "묶음을 못 정한 손익 계정: " + ", ".join(x.name for x in other))
    return grid_xml(g.cells, [40, 18, 18, 18], freeze="B5"), {"sec": sec, "acc": acc_row}


def _is_cash(x: TbRow) -> bool:
    return x.key == "cash" and not any(w in x.name for w in CASH_EXCLUDE) and not x.contra


def _cf_category(x: TbRow) -> str:
    """재무상태표 계정 -> op | inv | fin | other"""
    n = x.name
    k = x.key
    if k == "other_inv" and "퇴직" in n:
        return "op"      # 퇴직연금운용자산은 퇴직급여충당부채와 함께 영업활동
    if k in ("cash", "sec", "inv_cur", "lt_cash", "lt_sec", "equity_method", "other_inv", "ppe", "intangible"):
        return "inv"
    if k == "other_quick":
        return "inv" if "대여금" in n else "op"
    if k == "other_nca":
        return "inv" if any(w in n for w in ("보증금", "대여금", "회원권")) else "op"
    if k in ("borrow", "lt_borrow", "capital", "cap_surplus"):
        return "fin"
    if k == "other_cl":
        return "fin" if any(w in n for w in ("배당", "차입", "사채")) else "op"
    if k == "other_ncl":
        return "fin" if "보증금" in n else "op"
    if k in ("cap_adj", "aoci"):
        return "other"
    return "op"     # ar, lt_ar, inventory, tax_asset, ap, lt_ap, retire, provision


def _cf_sheet(st, rows: list[TbRow], bs: dict, is_: dict) -> tuple[bytes, dict]:
    g = Grid(st)
    b, i_ = _q("수정후BS"), _q("수정후IS")
    g.put("A1", "현금흐름정산표 (간접법)", "title")
    g.put("A2", "수정후BS·수정후IS를 가리키는 수식입니다. 재무상태표 계정마다 증감(당기-전기)이 한 번씩 들어가므로 "
                "아래 검증 칸이 0이면 빠진 계정이 없습니다. 분류·현금 없는 손익은 회계사가 검토해 고치세요.")
    g.row(4, ["구분", "금액", "산식 설명"], "head")
    r = 5
    noncash_used: dict[int, list[str]] = {}   # 재무상태표 TB 행 -> 빼 줄 손익 셀

    def is_cells(word: str) -> list[str]:
        return [f"{i_}!B{row}" for name, row in is_["acc"] if word in name]

    sections: dict[str, int] = {}
    g.put(f"A{r}", "Ⅰ. 영업활동으로 인한 현금흐름", "bold")
    sections["op"] = r
    r += 1
    g.put(f"A{r}", "  1. 당기순이익(손실)")
    g.put(f"B{r}", Formula(f"{i_}!B{is_['sec']['ni']}"), "num")
    g.put(f"C{r}", "수정후IS 당기순이익")
    op_lines = [r]
    r += 1
    bs_rows = [x for x in rows if x.side != "is"]
    add_lines = []
    for word, partners, sign in NONCASH:
        cells = is_cells(word)
        if not cells:
            continue
        if partners is None:
            target = [x for x in bs_rows if x.key == "ppe" and not x.contra]
        else:
            target = [x for x in bs_rows if any(p in x.name for p in partners)]
        if not target:
            continue
        expr = "+".join(cells)
        g.put(f"A{r}", f"  {'가산' if sign > 0 else '차감'}: {word}")
        g.put(f"B{r}", Formula(f"{'' if sign > 0 else '-'}({expr})"), "num")
        g.put(f"C{r}", f"현금 없는 손익 → '{target[0].name}' 증감에서 {'빼' if sign > 0 else '더해'} 줌")
        noncash_used.setdefault(target[0].tb_row, []).append(("-" if sign > 0 else "+") + f"({expr})")
        add_lines.append(r)
        r += 1
    op_lines += add_lines

    def bs_line(x: TbRow) -> str | None:
        row = bs["acc"].get(x.tb_row)
        if row is None:
            return None
        base = f"-({b}!D{row})" if x.side == "asset" else f"{b}!D{row}"
        return base + "".join(noncash_used.get(x.tb_row, []))

    cat_lines: dict[str, list[int]] = {"op": [], "inv": [], "fin": [], "other": []}
    retained_done = False
    by_cat: dict[str, list] = {"op": [], "inv": [], "fin": [], "other": []}
    for x in bs_rows:
        if _is_cash(x):
            continue
        if x.key == "retained":
            if not retained_done:
                by_cat["fin"].append(("retained", x))
                retained_done = True
            continue
        by_cat[_cf_category(x)].append(("acc", x))

    def write(cat: str, header: str | None):
        nonlocal r
        if header:
            g.put(f"A{r}", header)
            r += 1
        for kind, x in by_cat[cat]:
            if kind == "retained":
                gh = next(iter(bs["retained"]), None)
                if gh is None:
                    continue
                g.put(f"A{r}", "    이익잉여금 증감(당기순이익 제외: 배당 등)")
                g.put(f"B{r}", Formula(f"{b}!D{gh}-{i_}!B{is_['sec']['ni']}"), "num")
                g.put(f"C{r}", "수정후BS 이익잉여금 묶음 증감 - 당기순이익")
            else:
                f = bs_line(x)
                if f is None:
                    continue
                g.put(f"A{r}", f"    {x.label} {'감소(증가)' if x.side == 'asset' else '증가(감소)'}")
                g.put(f"B{r}", Formula(f), "num")
                g.put(f"C{r}", "수정후BS 증감" + (" - 현금 없는 손익" if x.tb_row in noncash_used else ""))
            cat_lines[cat].append(r)
            r += 1

    write("op", "  2. 영업활동으로 인한 자산·부채의 변동")
    for col in "B":
        g.put(f"{col}{sections['op']}", Formula("+".join(f"B{x}" for x in op_lines + cat_lines["op"]) or "0"),
              "boldnum")
    g.put(f"A{r}", "Ⅱ. 투자활동으로 인한 현금흐름", "bold")
    sections["inv"] = r
    r += 1
    write("inv", None)
    g.put(f"B{sections['inv']}", Formula("+".join(f"B{x}" for x in cat_lines["inv"]) or "0"), "boldnum")
    g.put(f"A{r}", "Ⅲ. 재무활동으로 인한 현금흐름", "bold")
    sections["fin"] = r
    r += 1
    write("fin", None)
    g.put(f"B{sections['fin']}", Formula("+".join(f"B{x}" for x in cat_lines["fin"]) or "0"), "boldnum")
    if by_cat["other"]:
        g.put(f"A{r}", "Ⅳ. 현금 없는 자본 변동 등 (분석 필요)", "bold")
        sections["other"] = r
        r += 1
        write("other", None)
        g.put(f"B{sections['other']}", Formula("+".join(f"B{x}" for x in cat_lines["other"]) or "0"), "boldnum")
    r += 1
    g.put(f"A{r}", "현금의 증가(감소)", "bold")
    g.put(f"B{r}", Formula("+".join(f"B{sections[k]}" for k in sections)), "boldnum")
    inc = r
    r += 1
    cash = [x for x in bs_rows if _is_cash(x) and x.tb_row in bs["acc"]]
    g.put(f"A{r}", "기초의 현금", "bold")
    g.put(f"B{r}", Formula("+".join(f"{b}!C{bs['acc'][x.tb_row]}" for x in cash) or "0"), "boldnum")
    g.put(f"C{r}", "현금으로 본 계정: " + (", ".join(x.name for x in cash) or "없음"))
    beg = r
    r += 1
    g.put(f"A{r}", "기말의 현금", "bold")
    g.put(f"B{r}", Formula(f"B{beg}+B{inc}"), "boldnum")
    end = r
    r += 2
    g.put(f"A{r}", "검증: 기말 현금(수정후BS) - 기말의 현금 (0이어야 함)", "bold")
    end_cash = "+".join(f"{b}!B{bs['acc'][x.tb_row]}" for x in cash) or "0"
    g.put(f"B{r}", Formula(f"ROUND(({end_cash})-B{end},0)"), "check")
    g.put(f"C{r}", "0이 아니면 전기 금액이 빠졌거나(BS 탭 전기), 재무상태표 대차가 맞지 않는 것")
    return grid_xml(g.cells, [46, 18, 56], freeze="B5"), {"sections": sections, "check": r}


# ---------------------------------------------------------------- 수정후 금액 (파이썬 계산)

@dataclass
class Figure:
    row: TbRow
    current: float       # 수정후 당기 (자산·비용은 차변-대변, 부채·자본·수익은 대변-차변)
    prior: float | None  # 전기 (BS·IS 탭)
    adjustment: float    # 수정분개 (같은 부호)


def figures(workpaper: str | Path, aje_path: str | Path | None = None) -> tuple[list[Figure], float, list[str]]:
    """수정분개까지 반영한 계정별 금액을 계산한다 (주석 자료 양식·감사보고서용).
    aje_path가 없으면 조서 안 AJE 탭을 쓴다. 반환: (계정별 금액, 당기순이익, 참고)"""
    tpl, rows = read_rows(workpaper)
    notes: list[str] = []
    wb = open_workbook(workpaper, read_only=True, data_only=True)
    try:
        has_aje = "AJE" in wb.sheetnames
        bs_prior, is_prior = {}, {}
        ws = wb["BS"]
        for i, r in enumerate(ws.iter_rows(min_row=1, max_col=6, values_only=True), start=1):
            bs_prior[i] = r[5] if len(r) > 5 else None
        for r in wb["IS"].iter_rows(min_row=1, max_col=6, values_only=True):
            if len(r) > 5 and isinstance(r[2], str) and isinstance(r[5], (int, float)):
                is_prior.setdefault(squash(r[2]), r[5])
    finally:
        wb.close()
    lines = read_aje(aje_path) if aje_path else (read_aje(workpaper) if has_aje else [])
    if not aje_path and not has_aje:
        notes.append("수정분개가 없어 시산표 금액 그대로 썼습니다")
    rows, problems, more = add_new_accounts(tpl, rows, lines)
    notes += problems
    net: dict[str, float] = {}
    for ln in lines:
        net[ln.account] = net.get(ln.account, 0.0) + ln.debit - ln.credit
    out = []
    for x in rows:
        credit_view = x.side == "liab" or (x.side == "is" and x.key in REVENUE_KEYS)
        sign = -1 if credit_view else 1
        adj = sign * net.get(x.label, 0.0)
        cur = sign * (x.db - x.cb) + adj
        if x.side == "is":
            prior = is_prior.get(x.name)
        else:
            prior = to_number(bs_prior.get(x.bs_row)) if x.bs_row and bs_prior.get(x.bs_row) is not None else None
        out.append(Figure(x, cur, prior, adj))
    ni = sum(f.current * (1 if f.row.key in REVENUE_KEYS else -1) for f in out if f.row.side == "is")
    return out, ni, notes


# ---------------------------------------------------------------- 반영

def apply(workpaper: str | Path, aje_path: str | Path, output: str | Path) -> Result:
    workpaper, output = Path(workpaper), Path(output)
    if workpaper.resolve() == output.resolve():
        raise ValueError("원본 조서에 덮어쓰지 않습니다. 다른 이름으로 저장해 주세요.")
    tpl, rows = read_rows(workpaper)
    lines = read_aje(aje_path)
    res = Result(entries=len({ln.no for ln in lines}), lines=len(lines),
                 debit=sum(ln.debit for ln in lines), credit=sum(ln.credit for ln in lines))
    rows, problems, notes = add_new_accounts(tpl, rows, lines)
    res.problems = problems + check_aje(lines, rows)
    res.notes += notes
    pkg = XlsxPackage(workpaper)
    for name in NEW_TABS:
        if pkg.has_sheet(name):
            pkg.remove_sheet(name)
            res.notes.append(f"전에 만든 '{name}' 탭을 새로 만들었습니다")
    st = pkg.add_basic_styles()
    aje_xml, aje_total = _aje_sheet(st, lines)
    ws_xml, wpos, w = _worksheet(st, rows, aje_total)
    bs_xml, bs = _bs_sheet(st, rows, wpos, w)
    is_xml, is_ = _is_sheet(st, rows, wpos, w)
    cf_xml, _ = _cf_sheet(st, rows, bs, is_)
    for name, xml in zip(NEW_TABS, (aje_xml, ws_xml, bs_xml, is_xml, cf_xml)):
        pkg.add_sheet(name, xml)
    # 리드시트 Audited(F/S) = G/L + 수정분개
    n = 0
    for lead, pairs in tpl.lead_rows.items():
        ed = pkg.sheet(lead)
        for lr, tr in pairs:
            p = wpos.get(tr)
            ed.set(f"I{lr}", Formula(f"H{lr}+'정산표'!N{p}" if p else f"H{lr}"))
            n += 1
    res.notes.append(f"리드시트 {len(tpl.lead_rows)}개 탭 {n}줄의 Audited(F/S) 칸을 'G/L + 수정분개' 수식으로 바꿨습니다")
    if not any(x.bs_row for x in rows if x.side != "is"):
        res.notes.append("BS 탭 전기 금액이 없어 현금흐름정산표 검증이 맞지 않을 수 있습니다")
    pkg.recalc_on_open()
    res.output = pkg.save(output)
    return res

