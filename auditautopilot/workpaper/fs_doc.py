"""재무제표·주석 초안 (워드 .docx).

- 재무상태표·손익계산서·자본변동표·현금흐름표: 수정분개까지 반영한 정산표 금액 (조서의 수정후BS·수정후IS·
  정산표-CF 탭과 같은 합계를 파이썬으로 계산한 값). 워드 표라 그대로 복사해 붙여 넣을 수 있다.
- 주석: 회사가 채워 보낸 주석 자료 엑셀(주석 자료요청 양식)의 시트마다 주석 하나. 채운 파일이 없으면
  정산표 금액만 들어간 양식 내용으로 만들고, 회사가 채울 주석은 [ ]로 비워 둔다.

표준 라이브러리로 docx를 만든다 (report.write_docx).
"""

from __future__ import annotations

import tempfile
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from .adjust import BS_MAJORS, IS_ORDER, MFG_KEYS, NONCASH, REVENUE_KEYS, Figure, _cf_category, _is_cash, figures
from .report import PAGE_BREAK, _kdate, _para, _run, read_info, write_docx
from .xlsx_xml import open_workbook

EQUITY_COLS = [("capital", "자본금"), ("cap_surplus", "자본잉여금"), ("cap_adj", "자본조정"),
               ("aoci", "기타포괄손익누계액"), ("retained", "이익잉여금")]


# ---------------------------------------------------------------- 표 한 줄

@dataclass
class Line:
    label: str
    current: float | None = None
    prior: float | None = None
    level: int = 0          # 0 대분류·합계(굵게), 1 묶음(굵게), 2 계정
    bold: bool = False


def _shown(f: Figure) -> bool:
    """당기·전기 모두 0인 계정은 재무제표에 넣지 않는다."""
    return abs(f.current) >= 0.5 or abs(f.prior or 0) >= 0.5


def _sum(vals) -> float:
    return sum(v or 0.0 for v in vals)


def _prior_or_none(fs: list[Figure]):
    return None if fs and all(f.prior is None for f in fs) else _sum(f.prior for f in fs)


# ---------------------------------------------------------------- 재무상태표

def balance_sheet(figs: list[Figure], ni: float) -> tuple[list[Line], dict]:
    out: list[Line] = []
    major_tot: dict[str, tuple[float, float]] = {}
    for major in BS_MAJORS:
        items = [f for f in figs if f.row.side != "is" and f.row.major == major]
        if not items:
            continue
        head = Line(major, level=0, bold=True)
        out.append(head)
        mc = mp = 0.0
        for grow in dict.fromkeys(f.row.group_row for f in items):
            gi = [f for f in items if f.row.group_row == grow]
            gl = Line(gi[0].row.group_label, level=1, bold=True)
            out.append(gl)
            for f in gi:
                if _shown(f):
                    out.append(Line(f.row.label, f.current, f.prior, 2))
            cur, pri = _sum(f.current for f in gi), _sum(f.prior for f in gi)
            if gi[0].row.key == "retained":
                out.append(Line("당기순이익(손실)", ni, None, 2))
                cur += ni
            gl.current, gl.prior = cur, pri
            if out[-1] is gl and abs(cur) < 0.5 and abs(pri) < 0.5:
                out.pop()           # 계정이 모두 0인 묶음
            mc, mp = mc + cur, mp + pri
        head.current, head.prior = mc, mp
        major_tot[major] = (mc, mp)

    def tot(majors):
        return (sum(major_tot.get(m, (0, 0))[0] for m in majors), sum(major_tot.get(m, (0, 0))[1] for m in majors))
    totals = {"자산총계": tot(("유동자산", "비유동자산")), "부채총계": tot(("유동부채", "비유동부채")),
              "자본총계": tot(("자본",))}
    totals["부채와자본총계"] = (totals["부채총계"][0] + totals["자본총계"][0], totals["부채총계"][1] + totals["자본총계"][1])
    # 자산총계는 자산 묶음 뒤, 부채·자본 총계는 각 묶음 뒤에 넣는다
    lines: list[Line] = []
    asset_done = liab_done = False
    for ln in out:
        if ln.level == 0 and ln.label in ("유동부채", "비유동부채") and not asset_done:
            lines.append(Line("자산총계", *totals["자산총계"], bold=True))
            asset_done = True
        if ln.level == 0 and ln.label == "자본" and not liab_done:
            if not asset_done:
                lines.append(Line("자산총계", *totals["자산총계"], bold=True))
                asset_done = True
            lines.append(Line("부채총계", *totals["부채총계"], bold=True))
            liab_done = True
        lines.append(ln)
    if not asset_done:
        lines.append(Line("자산총계", *totals["자산총계"], bold=True))
    if not liab_done:
        lines.append(Line("부채총계", *totals["부채총계"], bold=True))
    lines.append(Line("자본총계", *totals["자본총계"], bold=True))
    lines.append(Line("부채와자본총계", *totals["부채와자본총계"], bold=True))
    return lines, totals


# ---------------------------------------------------------------- 손익계산서

def _diff(a, b):
    c = a[0] - b[0]
    p = None if a[1] is None and b[1] is None else (a[1] or 0) - (b[1] or 0)
    return c, p


def income_statement(figs: list[Figure], ni: float) -> tuple[list[Line], dict]:
    out: list[Line] = []
    sec: dict[str, tuple[float, float | None]] = {}
    for key, label in IS_ORDER:
        keys = {key} | (MFG_KEYS if key == "cogs" else set())
        items = [f for f in figs if f.row.side == "is" and f.row.key in keys]
        cur, pri = _sum(f.current for f in items), _prior_or_none(items)
        out.append(Line(label, cur, pri, 0, True))
        for f in items:
            if _shown(f):
                out.append(Line(f.row.label, f.current, f.prior, 2))
        sec[key] = (cur, pri)
        if key == "cogs":
            c, p = _diff(sec["sales"], sec["cogs"])
            sec["gross"] = (c, p)
            out.append(Line("Ⅲ. 매출총이익", c, p, 0, True))
        if key == "sga":
            c, p = _diff(sec["gross"], sec["sga"])
            sec["op"] = (c, p)
            out.append(Line("Ⅴ. 영업이익(손실)", c, p, 0, True))
        if key == "nonop_exp":
            c = sec["op"][0] + sec["nonop_inc"][0] - sec["nonop_exp"][0]
            ps = (sec["op"][1], sec["nonop_inc"][1], sec["nonop_exp"][1])
            p = None if all(x is None for x in ps) else (ps[0] or 0) + (ps[1] or 0) - (ps[2] or 0)
            sec["ebt"] = (c, p)
            out.append(Line("Ⅷ. 법인세비용차감전순이익(손실)", c, p, 0, True))
    p_ebt, p_tax = sec["ebt"][1], sec["tax"][1]
    p_ni = None if p_ebt is None and p_tax is None else (p_ebt or 0) - (p_tax or 0)
    out.append(Line("Ⅹ. 당기순이익(손실)", ni, p_ni, 0, True))
    sec["ni"] = (ni, p_ni)
    return out, sec


# ---------------------------------------------------------------- 자본변동표

def equity_changes(figs: list[Figure], ni: float) -> tuple[list[str], list[list[float]]]:
    """열: 자본 묶음, 줄: 기초·당기순이익·기타·기말"""
    cols = []
    for key, label in EQUITY_COLS:
        fs = [f for f in figs if f.row.side == "liab" and f.row.major == "자본" and f.row.key == key]
        if fs or key == "retained":
            cols.append((label, _sum(f.prior for f in fs), _sum(f.current for f in fs) + (ni if key == "retained" else 0),
                         ni if key == "retained" else 0.0))
    other = [f for f in figs if f.row.side == "liab" and f.row.major == "자본"
             and f.row.key not in {k for k, _ in EQUITY_COLS}]
    if other:
        cols.append(("기타 자본", _sum(f.prior for f in other), _sum(f.current for f in other), 0.0))
    heads = [c[0] for c in cols] + ["합계"]
    beg = [c[1] for c in cols]
    nil = [c[3] for c in cols]
    end = [c[2] for c in cols]
    etc = [e - b - n for b, n, e in zip(beg, nil, end)]
    rows = [beg + [sum(beg)], nil + [sum(nil)], etc + [sum(etc)], end + [sum(end)]]
    return heads, rows


# ---------------------------------------------------------------- 현금흐름표 (간접법)

def cash_flow(figs: list[Figure], ni: float) -> tuple[list[Line], dict]:
    """간접법 현금흐름표 (조서 정산표-CF 탭과 활동별 합계가 같다)."""
    bs = [f for f in figs if f.row.side != "is" and f.row.major in BS_MAJORS]
    is_ = [f for f in figs if f.row.side == "is"]
    adj: dict[int, float] = {}
    op = [Line("1. 당기순이익(손실)", ni, level=1)]
    for word, partners, sign in NONCASH:
        amt = [f.current for f in is_ if word in f.row.label]
        if not amt:
            continue
        if partners is None:
            target = [f for f in bs if f.row.key == "ppe" and not f.row.contra]
        else:
            target = [f for f in bs if any(p in f.row.name for p in partners)]
        if not target:
            continue
        v = sign * sum(amt)
        op.append(Line(f"{'가산' if sign > 0 else '차감'}: {word}", v, level=2))
        adj[target[0].row.tb_row] = adj.get(target[0].row.tb_row, 0.0) - v

    def change(f: Figure) -> float:
        d = f.current - (f.prior or 0.0)
        return (-d if f.row.side == "asset" else d) + adj.get(f.row.tb_row, 0.0)

    cats: dict[str, list[Line]] = {"op": [], "inv": [], "fin": [], "other": []}
    retained = [f for f in bs if f.row.key == "retained"]
    for f in bs:
        if _is_cash(f.row) or f.row.key == "retained":
            continue
        word = "감소(증가)" if f.row.side == "asset" else "증가(감소)"
        cats[_cf_category(f.row)].append(Line(f"{f.row.label} {word}", change(f), level=2))
    if retained:
        v = _sum(f.current for f in retained) + ni - _sum(f.prior for f in retained) - ni
        cats["fin"].append(Line("이익잉여금 증감(당기순이익 제외: 배당 등)", v, level=2))
    out: list[Line] = []
    sections: dict[str, float] = {}
    opv = _sum(x.current for x in op) + _sum(x.current for x in cats["op"])
    out.append(Line("Ⅰ. 영업활동으로 인한 현금흐름", opv, bold=True))
    out += op
    if cats["op"]:
        out.append(Line("2. 영업활동으로 인한 자산·부채의 변동", _sum(x.current for x in cats["op"]), level=1))
        out += cats["op"]
    sections["op"] = opv
    for key, label in (("inv", "Ⅱ. 투자활동으로 인한 현금흐름"), ("fin", "Ⅲ. 재무활동으로 인한 현금흐름"),
                       ("other", "Ⅳ. 현금 없는 자본 변동 등 (분석 필요)")):
        if key == "other" and not cats["other"]:
            continue
        v = _sum(x.current for x in cats[key])
        out.append(Line(label, v, bold=True))
        out += cats[key]
        sections[key] = v
    inc = sum(sections.values())
    cash = [f for f in bs if _is_cash(f.row)]
    beg = _sum(f.prior for f in cash)
    end = beg + inc
    out += [Line("Ⅴ. 현금의 증가(감소)", inc, bold=True), Line("Ⅵ. 기초의 현금", beg, bold=True),
            Line("Ⅶ. 기말의 현금", end, bold=True)]
    return out, {"increase": inc, "begin": beg, "end": end, "check": _sum(f.current for f in cash) - end}


# ---------------------------------------------------------------- 주석

@dataclass
class NoteTable:
    title: str
    kind: str = ""
    head: list[str] = field(default_factory=list)
    rows: list[list] = field(default_factory=list)
    text: str = ""          # 표 대신 글 (해당사항 없음 등)


def _cell_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, datetime):
        v = v.date()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, bool):
        return "예" if v else "아니오"
    if isinstance(v, (int, float)):
        return _amt(v)
    return str(v).strip()


def read_filled_notes(path: str | Path) -> list[NoteTable]:
    """회사가 채운 주석 자료 엑셀 (notes.request_workbook 양식)을 읽는다.

    시트마다 1행 '회사명 주석 자료 - 제목', 2행 '[구분] 안내', 4행부터 표. '합계' 줄의 수식은 다시 더한다
    (엑셀에서 열고 저장하지 않으면 수식 값이 없으므로)."""
    wb_f = open_workbook(path, read_only=True, data_only=False)
    wb_v = open_workbook(path, read_only=True, data_only=True)
    out: list[NoteTable] = []
    try:
        for name in wb_v.sheetnames:
            if name == "목차":
                continue
            vals = [list(r) for r in wb_v[name].iter_rows(values_only=True)]
            forms = [list(r) for r in wb_f[name].iter_rows(values_only=True)]
            if not vals:
                continue
            a1 = str(vals[0][0] or name) if vals[0] else name
            title = a1.split(" - ", 1)[1].strip() if " - " in a1 else name.split(".", 1)[-1]
            kind = ""
            if len(vals) > 1 and vals[1] and isinstance(vals[1][0], str) and vals[1][0].startswith("["):
                kind = vals[1][0][1:vals[1][0].find("]")]
            head_i = next((i for i in range(2, min(len(vals), 12)) if sum(1 for v in vals[i] if v not in (None, "")) >= 2),
                          None)
            if head_i is None:
                continue
            head = [str(v or "").strip() for v in vals[head_i]]
            body = []
            for i in range(head_i + 1, len(vals)):
                row = (vals[i] + [None] * len(head))[:len(head)]
                frow = ((forms[i] if i < len(forms) else []) + [None] * len(head))[:len(head)]
                if str(row[0] or frow[0] or "").strip() == "합계":
                    sums = []
                    for j in range(1, len(head)):
                        nums = [r[j] for r in body if isinstance(r[j], (int, float)) and not isinstance(r[j], bool)]
                        is_formula = isinstance(frow[j], str) and frow[j].startswith("=")
                        sums.append(sum(nums) if (is_formula or row[j] is not None) and nums else row[j])
                    body.append(["합계"] + sums)
                    continue
                if all(v in (None, "") for v in row[1:]) and (row[0] in (None, "")):
                    continue
                body.append(row)
            # 아무것도 안 채운 열은 뺀다 (과목 칸은 둔다)
            keep = [0] + [j for j in range(1, len(head)) if any(r[j] not in (None, "") for r in body if r[0] != "합계")]
            body_data = [r for r in body if r[0] != "합계"]
            if not body_data or (kind == "회사 작성" and not any(
                    isinstance(v, str) and v.strip() for r in body_data for v in r)):
                out.append(NoteTable(title, kind, text=_no_data_text(kind, body_data)))
                continue
            only_none = len(body_data) == 1 and all(
                isinstance(v, str) and v.strip() in ("없음", "해당없음", "해당 없음") for v in body_data[0] if v)
            if only_none:
                out.append(NoteTable(title, kind, text="보고기간말 현재 해당사항이 없습니다."))
                continue
            out.append(NoteTable(title, kind, [head[j] for j in keep], [[r[j] for j in keep] for r in body]))
    finally:
        wb_f.close()
        wb_v.close()
    return out


def _no_data_text(kind: str, rows) -> str:
    if kind == "회사 작성":
        return "[회사 자료를 받지 못했습니다. 해당 없으면 '보고기간말 현재 해당사항이 없습니다.'로 두세요.]"
    return "[자료를 받지 못했습니다.]"


# ---------------------------------------------------------------- docx

def _amt(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer() or isinstance(v, int):
        v = int(v)
        return f"({-v:,})" if v < 0 else f"{v:,}"
    if abs(v - round(v)) < 1e-9:
        v = round(v)
        return f"({-v:,})" if v < 0 else f"{v:,}"
    return f"{v:,.2f}"


def _tc(text: str, bold=False, align="left", width: int | None = None, indent: int = 0) -> str:
    jc = {"left": "", "right": '<w:jc w:val="right"/>', "center": '<w:jc w:val="center"/>'}[align]
    ind = f'<w:ind w:left="{indent}"/>' if indent else ""
    w = f'<w:tcW w:w="{width}" w:type="dxa"/>' if width else '<w:tcW w:w="0" w:type="auto"/>'
    return (f'<w:tc><w:tcPr>{w}</w:tcPr><w:p><w:pPr><w:spacing w:after="0"/>{ind}{jc}</w:pPr>'
            f"{_run(text, bold, 18)}</w:p></w:tc>")


def _tbl(head: list[str], rows: list[tuple[list[str], bool, int]], widths: list[int] | None = None) -> str:
    """rows: [(칸 글, 굵게, 들여쓰기)]. 첫 칸은 왼쪽, 나머지는 오른쪽 정렬."""
    borders = "".join(f'<w:{b} w:val="single" w:sz="4" w:color="808080"/>'
                      for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
    out = [f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>{borders}</w:tblBorders>'
           '<w:tblCellMar><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr>']
    if widths:
        out.append("<w:tblGrid>" + "".join(f'<w:gridCol w:w="{w}"/>' for w in widths) + "</w:tblGrid>")

    def wd(i):
        return widths[i] if widths and i < len(widths) else None
    out.append('<w:tr><w:trPr><w:tblHeader/></w:trPr>' +
               "".join(_tc(h, True, "center", wd(i)) for i, h in enumerate(head)) + "</w:tr>")
    for cells, bold, indent in rows:
        out.append("<w:tr>" + "".join(_tc(t, bold, "left" if i == 0 else "right", wd(i), indent if i == 0 else 0)
                                       for i, t in enumerate(cells)) + "</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


def _statement(lines: list[Line], cur_head: str, pri_head: str | None) -> str:
    head = ["과목", cur_head] + ([pri_head] if pri_head else [])
    rows = []
    for ln in lines:
        cells = [ln.label, _amt(ln.current)] + ([_amt(ln.prior)] if pri_head else [])
        rows.append((cells, ln.bold or ln.level < 2, {0: 0, 1: 200, 2: 400}[ln.level]))
    return _tbl(head, rows, [5000, 2200, 2200] if pri_head else [6000, 3200])


def _title_block(title: str, company: str, period: str) -> str:
    return (_para(title, "title") + _para(period, "sign") +
            f'<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="9300"/></w:tabs><w:spacing w:after="60"/></w:pPr>{_run(company)}<w:r><w:tab/></w:r>'
            f'{_run("(단위: 원)")}</w:p>')


@dataclass
class Result:
    output: Path
    notes_used: str                    # 쓴 주석 자료 파일 이름 (없으면 "")
    notes: list[str] = field(default_factory=list)


def draft(workpaper: str | Path, output: str | Path, notes_xlsx: str | Path | None = None,
          aje_path: str | Path | None = None) -> Result:
    info = read_info(workpaper)
    figs, ni, extra = figures(workpaper, aje_path)
    co = "" if info.company.startswith("[") else info.company
    ye, pe = _kdate(info.year_end), _kdate(info.prior_end)
    cur_h, pri_h = (f"당기 ({info.year_end.year}년)" if info.year_end else "당기"), (
        f"전기 ({info.prior_end.year}년)" if info.prior_end else "전기")
    body = []
    body.append(_para(f"{co or '[회사명]'}", "title") + _para("재 무 제 표 (초안)", "title") +
                _para(f"당기 {ye} 현재 · 전기 {pe} 현재", "sign") +
                _para("감사인 정산표(수정분개 반영) 금액으로 만든 초안입니다. 회사 재무제표와 대사해 주세요.", "note"))
    bs, tot = balance_sheet(figs, ni)
    body.append(PAGE_BREAK + _title_block("재 무 상 태 표", co, f"당기 {ye} 현재 · 전기 {pe} 현재") +
                _statement(bs, cur_h, pri_h))
    is_, sec = income_statement(figs, ni)
    period = f"당기 {ye}로 종료되는 보고기간 · 전기 {pe}로 종료되는 보고기간"
    body.append(PAGE_BREAK + _title_block("손 익 계 산 서", co, period) + _statement(is_, cur_h, pri_h))
    heads, eq = equity_changes(figs, ni)
    labels = [f"{pe} (전기말)", "당기순이익(손실)", "기타 변동 (배당·증자 등, 확인 필요)", f"{ye} (당기말)"]
    eq_rows = [([lb] + [_amt(v) for v in vals], i in (0, 3), 0) for i, (lb, vals) in enumerate(zip(labels, eq))]
    n = len(heads) + 1
    body.append(PAGE_BREAK + _title_block("자 본 변 동 표", co, f"당기 {ye}로 종료되는 보고기간") +
                _tbl(["구분"] + heads, eq_rows, [2600] + [max(1100, 6800 // (n - 1))] * (n - 1)))
    cf, cfi = cash_flow(figs, ni)
    body.append(PAGE_BREAK + _title_block("현 금 흐 름 표", co, f"당기 {ye}로 종료되는 보고기간") +
                _statement(cf, cur_h, None))
    if abs(cfi["check"]) >= 1:
        extra.append(f"현금흐름표 기말 현금이 재무상태표 현금과 {cfi['check']:,.0f}원 다릅니다 (전기 금액을 확인해 주세요)")
    if abs(tot["자산총계"][0] - tot["부채와자본총계"][0]) >= 1:
        extra.append(f"재무상태표 대차가 {tot['자산총계'][0] - tot['부채와자본총계'][0]:,.0f}원 맞지 않습니다")

    # 주석
    used = ""
    if notes_xlsx:
        items = read_filled_notes(notes_xlsx)
        used = Path(notes_xlsx).name
    else:
        from . import notes as notes_mod
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            res = notes_mod.request_workbook(workpaper, Path(tmp) / "n.xlsx", co, aje_path)
            items = read_filled_notes(res.output)
        extra.append("회사가 채운 주석 자료가 없어 정산표 금액만 넣었습니다. 회사 작성 주석은 [ ]로 비어 있습니다")
    body.append(PAGE_BREAK + _para("주석", "title") + _para(f"당기 {ye} 현재 · 전기 {pe} 현재", "sign") +
                _para(f"{co or '[회사명]'}", "p"))
    no = 1
    body.append(_para(f"{no}. 회사의 개요", "h") +
                _para("[설립일, 본점 소재지, 주요 사업, 자본금·주주 현황을 적어 주세요.]", "note"))
    no += 1
    body.append(_para(f"{no}. 중요한 회계정책", "h") +
                _para("회사는 일반기업회계기준에 따라 재무제표를 작성하였습니다. [수익인식, 재고자산 평가방법, "
                      "유형자산 감가상각 방법·내용연수, 대손충당금 설정 기준 등 적용한 회계정책을 적어 주세요.]", "note"))
    for nt in items:
        no += 1
        body.append(_para(f"{no}. {nt.title}", "h"))
        if nt.text:
            body.append(_para(nt.text, "note" if nt.text.startswith("[") else "p"))
            continue
        rows = [([_cell_text(v) for v in r], str(r[0] or "").strip() == "합계", 0) for r in nt.rows]
        k = len(nt.head)
        first = 2400 if k > 4 else 3200
        rest = max(900, (9300 - first) // max(1, k - 1))
        body.append(_tbl(nt.head, rows, [first] + [rest] * (k - 1)))
        body.append(_para("", "p"))
    out = write_docx(Path(output), "".join(body))
    return Result(out, used, extra)
