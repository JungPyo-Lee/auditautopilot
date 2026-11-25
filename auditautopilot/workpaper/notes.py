"""주석 자료요청 양식: 회사에 보내 주석에 필요한 자료를 채워 받는 엑셀.

두 가지로 나눈다.
- 정산표 금액 채움: 수정분개까지 반영한 정산표 금액(당기·전기)을 미리 넣고, 회사는 내역·변동·조건만 채운다
  (현금·금융상품, 매출채권·대손충당금, 재고자산, 유형·무형자산 변동, 차입금, 퇴직급여, 자본금, 판관비, 영업외손익, 법인세)
- 회사 작성: 장부에 없는 정보라 회사가 처음부터 채운다
  (특수관계자, 담보, 보험, 토지 공시지가, 지급보증, 약정, 소송·우발부채, 사용제한 예금, 주주, 리스, 보고기간후사건)

계정이 없는 주석(예: 무형자산이 없으면 무형자산 변동)은 넣지 않는다.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .adjust import Figure, figures

KIND_FILLED = "정산표 금액 채움"
KIND_COMPANY = "회사 작성"
PREFILL = PatternFill("solid", fgColor="E7E6E6")
INPUT = PatternFill("solid", fgColor="FFF8DC")
HEAD = PatternFill("solid", fgColor="DDE6F4")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


@dataclass
class Note:
    title: str
    kind: str
    columns: list[str]
    rows: list[list] = field(default_factory=list)     # 미리 채운 줄 (값 또는 None)
    prefilled_cols: int = 0                            # 앞에서 몇 칸이 미리 채운 금액인지 (과목 칸 제외)
    blank_rows: int = 0                                # 회사가 채울 빈 줄
    guide: str = ""
    total: bool = False
    sum_cols: int = 0                                  # 합계 줄을 만들 금액 칸 수 (과목 칸 다음부터)


def _pick(figs: list[Figure], *keys, contra: bool | None = None, words: tuple[str, ...] = ()) -> list[Figure]:
    out = []
    for f in figs:
        if f.row.key not in keys:
            continue
        if contra is not None and f.row.contra != contra:
            continue
        if words and not any(w in f.row.name for w in words):
            continue
        out.append(f)
    return out


def _amounts(fs: list[Figure]) -> list[list]:
    return [[f.row.label, f.current, f.prior] for f in fs]


def _movement(figs: list[Figure], key: str) -> list[list]:
    """유형·무형자산: 자산별 기초 장부금액(전기 취득원가-누계액)과 기말 장부금액을 미리 넣는다."""
    rows, current = [], None
    for f in [x for x in figs if x.row.key == key]:
        if not f.row.contra:
            current = [f.row.label, f.prior or 0.0, f.current]
            rows.append(current)
        elif current is not None:
            current[1] += f.prior or 0.0
            current[2] += f.current
    return [[name, beg, None, None, None, None, end] for name, beg, end in rows]


def build_notes(figs: list[Figure]) -> list[Note]:
    notes: list[Note] = []
    amt = ["과목", "당기말(정산표)", "전기말"]

    def filled(title, fs, extra, guide, blank=0):
        if fs:
            notes.append(Note(title, KIND_FILLED, amt + extra, [r + [None] * len(extra) for r in _amounts(fs)], 2,
                              blank, guide, total=True, sum_cols=2))

    filled("현금및현금성자산·금융상품", _pick(figs, "cash", "lt_cash"),
           ["금융기관", "계좌 종류", "이자율(%)", "만기", "사용제한 여부·사유"],
           "계좌별로 나눠야 하면 아래 빈 줄에 같은 과목으로 더 적어 주세요.", 5)
    ar = _pick(figs, "ar", "lt_ar")
    if ar:
        notes.append(Note("매출채권과 대손충당금", KIND_FILLED, amt + ["3개월 이내", "3~6개월", "6개월~1년", "1년 초과", "비고"],
                          [r + [None] * 5 for r in _amounts(ar)], 2, 0,
                          "채권은 연령별로 나눠 주세요(합계가 당기말과 같아야 합니다). 대손충당금 줄은 설정 기준을 비고에 적어 주세요.",
                          total=True, sum_cols=6))
        allow = [f for f in ar if f.row.contra]
        if allow:
            notes.append(Note("대손충당금 변동", KIND_FILLED, ["과목", "기초", "설정(환입)", "제각", "기타", "기말(정산표)"],
                              [[f.row.label, -(f.prior or 0.0), None, None, None, -f.current] for f in allow], 0, 0,
                              "기초·기말은 정산표 금액입니다. 기초 + 설정 - 제각 + 기타 = 기말이 되게 채워 주세요."))
    filled("재고자산", _pick(figs, "inventory"), ["평가충당금", "담보 제공액", "보관 장소", "비고"],
           "저가법 평가 결과와 담보로 제공한 재고가 있으면 적어 주세요.")
    for key, title in (("ppe", "유형자산 변동"), ("intangible", "무형자산 변동")):
        mv = _movement(figs, key)
        if mv:
            notes.append(Note(title, KIND_FILLED, ["과목", "기초 장부금액", "취득", "처분", "감가상각(상각)", "대체·기타",
                                                   "기말 장부금액(정산표)"], mv, 0, 0,
                              "기초·기말 장부금액(취득원가-누계액)은 정산표 금액입니다. 기초+취득-처분-상각±대체 = 기말이 되게 채워 주세요."
                              " 담보 제공 자산은 '담보 제공 자산' 시트에 적어 주세요.", total=True, sum_cols=6))
    filled("차입금", _pick(figs, "borrow", "lt_borrow"), ["차입처", "이자율(%)", "만기일", "상환 방법", "담보·보증"],
           "차입처마다 나눠 아래 빈 줄에 적어 주세요(과목별 합계가 당기말과 같아야 합니다).", 8)
    filled("퇴직급여충당부채·퇴직연금", _pick(figs, "retire") + _pick(figs, "other_inv", words=("퇴직",)),
           ["당기 지급액", "당기 설정액", "퇴직연금 종류(DB/DC)", "운용기관"], "퇴직연금 가입 형태와 운용기관을 적어 주세요.")
    filled("자본금", _pick(figs, "capital"), ["발행할 주식 총수", "발행주식수", "1주당 액면금액", "당기 변동 내용"],
           "주주별 지분은 '주주 현황' 시트에 적어 주세요.")
    filled("판매비와관리비", _pick(figs, "sga"), ["비고"], "금액은 정산표 기준입니다. 다른 점이 있으면 비고에 적어 주세요.")
    filled("영업외수익·영업외비용", _pick(figs, "nonop_inc", "nonop_exp"), ["주요 내용"],
           "잡이익·잡손실 등 큰 금액은 내용을 적어 주세요.")
    filled("법인세비용", _pick(figs, "tax"), ["법인세 신고 예정액", "이연법인세 반영 여부"],
           "세무조정계산서(예상)가 있으면 함께 보내 주세요.")

    def company(title, cols, guide, n=10):
        notes.append(Note(title, KIND_COMPANY, cols, [], 0, n, guide))
    company("특수관계자 현황·거래", ["특수관계자명", "관계(지배·종속·임원 등)", "매출 등", "매입 등", "채권 잔액", "채무 잔액",
                               "자금 대여·차입", "지급보증"], "당기 중 거래와 기말 잔액을 특수관계자별로 적어 주세요.")
    company("담보 제공 자산", ["담보 제공 자산", "장부금액", "담보 설정액", "담보권자", "관련 차입금"],
            "차입금 등을 위해 담보로 제공한 자산을 적어 주세요.")
    company("보험 가입 내역", ["보험 종류", "부보 자산", "부보 금액", "보험회사", "만기"], "화재·재산·배상책임 보험 등을 적어 주세요.")
    land = [f for f in figs if f.row.key == "ppe" and "토지" in f.row.name]
    if land:
        notes.append(Note("토지 공시지가", KIND_COMPANY, ["소재지", "면적(㎡)", "장부금액", "공시지가(원/㎡)", "공시지가 합계"],
                          [[None, None, sum(f.current for f in land), None, None]], 0, 5,
                          f"토지 장부금액 합계(정산표) {sum(f.current for f in land):,.0f}원을 필지별로 나눠 적어 주세요. "
                          "공시지가는 기말 기준 최근 공시 금액입니다."))
    company("지급보증", ["구분(제공받은/제공한)", "보증처·피보증처", "보증 금액", "보증 내용"],
            "금융기관·특수관계자 등에게서 받거나 준 보증을 적어 주세요.")
    company("약정 사항", ["금융기관", "약정 종류(한도대출·L/C 등)", "약정 한도", "실행 금액", "만기"],
            "금융기관과 맺은 한도 약정을 적어 주세요.")
    company("소송·우발부채", ["사건 내용", "원고·피고", "소송 금액", "진행 상황", "예상 결과"],
            "진행 중인 소송과 우발 상황을 적어 주세요. 없으면 '없음'이라고 적어 주세요.", 5)
    company("사용이 제한된 예금", ["금융기관", "계좌", "금액", "제한 사유"], "질권 설정 등으로 쓸 수 없는 예금을 적어 주세요.", 5)
    company("주주 현황", ["주주명", "소유 주식수", "지분율(%)", "특수관계 여부"], "기말 현재 주주명부 기준으로 적어 주세요.")
    company("리스", ["리스 자산", "리스 회사", "리스 기간", "연간 리스료", "구분(금융·운용)"], "리스 계약이 있으면 적어 주세요.", 5)
    company("보고기간후사건", ["일자", "내용", "금액"], "결산일 이후 감사보고서일까지 생긴 중요한 일(차입, 증자, 소송, 재해 등)을 적어 주세요.", 5)
    return notes


def _write(ws, note: Note, company_name: str) -> None:
    ws.append([f"{company_name} 주석 자료 - {note.title}"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append([f"[{note.kind}] {note.guide}"])
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.append([])
    ws.append(note.columns)
    head = ws.max_row
    for c in ws[head]:
        c.font = Font(bold=True)
        c.fill = HEAD
        c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in note.rows:
        ws.append(r)
    for _ in range(note.blank_rows):
        ws.append([None] * len(note.columns))
    last = ws.max_row
    for row in ws.iter_rows(min_row=head + 1, max_row=last, max_col=len(note.columns)):
        for i, c in enumerate(row):
            c.border = BOX
            if isinstance(c.value, (int, float)) or (i > 0 and c.value is None):
                c.number_format = "#,##0"
            prefilled = row[0].row <= head + len(note.rows) and c.value is not None and i > 0 and (
                note.kind == KIND_FILLED)
            c.fill = PREFILL if prefilled else (INPUT if c.value is None else PatternFill())
    if note.total and note.rows:
        ws.append(["합계"] + [f"=SUM({col}{head + 1}:{col}{last})" for col in "BCDEFGHIJ"[:note.sum_cols]])
        for c in ws[ws.max_row]:
            c.font = Font(bold=True)
            c.number_format = "#,##0"
            c.border = BOX
    ws.column_dimensions["A"].width = 28
    for i in range(2, len(note.columns) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 16
    ws.freeze_panes = ws.cell(head + 1, 2)


def _sheet_name(i: int, title: str) -> str:
    bad = '[]:*?/\\'
    t = "".join(ch for ch in title if ch not in bad)
    return f"{i:02d}.{t}"[:31]


@dataclass
class Result:
    output: Path
    filled: list[str]
    company: list[str]
    notes: list[str]


def request_workbook(workpaper: str | Path, output: str | Path, company_name: str = "",
                     aje_path: str | Path | None = None) -> Result:
    figs, _, notes = figures(workpaper, aje_path)
    items = build_notes(figs)
    if not company_name:
        from .report import read_info
        name = read_info(workpaper).company
        company_name = "" if name.startswith("[") else name
    wb = openpyxl.Workbook()
    idx = wb.active
    idx.title = "목차"
    idx.append([f"{company_name} 주석 작성 자료 요청"])
    idx["A1"].font = Font(bold=True, size=14)
    idx.append(["회색 칸은 감사인 정산표(수정분개 반영) 금액입니다. 다르면 비고에 적어 주시고, 노란 칸을 채워 주세요."])
    idx.append(["'회사 작성' 시트는 장부에 없는 정보라 회사에서 처음부터 채워 주셔야 합니다. 해당 없으면 '없음'이라고 적어 주세요."])
    idx.append([])
    idx.append(["번호", "시트", "구분", "제출 여부", "비고"])
    for c in idx[idx.max_row]:
        c.font = Font(bold=True)
        c.fill = HEAD
    for i, n in enumerate(sorted(items, key=lambda n: n.kind != KIND_FILLED), start=1):
        name = _sheet_name(i, n.title)
        _write(wb.create_sheet(name), n, company_name)
        idx.append([i, name, n.kind, None, None])
        cell = idx.cell(idx.max_row, 2)
        cell.hyperlink = f"#'{name}'!A1"
        cell.font = Font(color="0563C1", underline="single")
        idx.cell(idx.max_row, 4).fill = INPUT
    for col, w in zip("ABCDE", (6, 34, 16, 12, 30)):
        idx.column_dimensions[col].width = w
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return Result(output, [n.title for n in items if n.kind == KIND_FILLED],
                  [n.title for n in items if n.kind == KIND_COMPANY], notes)
