"""회계사가 보던 조서 양식 그대로의 현금흐름 정산표(정산표-CF)와 보고서 형식 재무제표 탭.

- 정산표-CF        : 재무상태표 계정마다 전기·당기·증감액을 두고, 증감을 영업(Dr 유출·Cr 유입·자산부채변동)·
                     투자(Dr·Cr)·재무(Dr·Cr)·비현금흐름(Dr·Cr) 칸에 나눈다. 비고 칸은 계정마다 '나눈 금액 - 증감액'으로
                     0이면 다 나눈 것이다. 아래에 열 합계, 활동별 현금흐름, 현금 없는 손익 대사표(손익계산서 vs 배분)를 둔다.
                     배분은 수식으로 미리 채운다 (운전자본 → 자산부채변동, 감가상각비 등 → 짝 계정의 Cr(유입),
                     투자·재무 계정 → 증감 방향에 따라 Dr/Cr). 회계사가 셀을 고쳐 쓰면 보고서현금흐름표가 따라간다.
- 보고서 BS·보고서IS : 한 해에 열 2개 (안쪽 열 = 계정, 바깥 열 = 소계·합계). 당기 C·D, 전기 E·F.
- 자본변동표       : 자본 묶음(자본금·자본잉여금·…·이익잉여금)별 전기초 → 전기말 → 당기초 → 당기말.
- 보고서현금흐름표 : 정산표-CF 칸을 가리키는 간접법 현금흐름표. G열은 정산표-CF와 맞는지 보는 검증 칸.

숫자는 모두 수정후BS·수정후IS·정산표-CF를 가리키는 수식이다. 금액이 0인 계정 줄은 보고서 탭에서 뺀다.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .adjust import (BS_MAJORS, MFG_KEYS, NONCASH, REVENUE_KEYS, Figure, Grid, TbRow, _cf_category, _is_cash, _q)
from .xlsx_xml import Formula, grid_xml

TABS = ("정산표-CF", "보고서 BS", "보고서IS", "자본변동표", "보고서현금흐름표")
ROMAN = ["Ⅰ.", "Ⅱ.", "Ⅲ.", "Ⅳ.", "Ⅴ.", "Ⅵ.", "Ⅶ.", "Ⅷ.", "Ⅸ.", "Ⅹ."]
KOR = [f"{c}." for c in "가나다라마바사아자차카타파하거너더러머버서어저처커터퍼허고노도로모보소오조초코토포호"]
SUBCLASS = {"cash": "당좌자산", "sec": "당좌자산", "ar": "당좌자산", "other_quick": "당좌자산",
            "inv_cur": "당좌자산", "tax_asset": "당좌자산", "inventory": "재고자산",
            "lt_cash": "투자자산", "lt_sec": "투자자산", "equity_method": "투자자산", "other_inv": "투자자산",
            "ppe": "유형자산", "intangible": "무형자산", "other_nca": "기타비유동자산", "lt_ar": "기타비유동자산"}
SUBCLASS_ORDER = {"유동자산": ["당좌자산", "재고자산"],
                  "비유동자산": ["투자자산", "유형자산", "무형자산", "기타비유동자산"]}
EQUITY_ORDER = ["capital", "cap_surplus", "cap_adj", "aoci", "retained"]
EQUITY_NAME = {"capital": "자본금", "cap_surplus": "자본잉여금", "cap_adj": "자본조정",
               "aoci": "기타포괄손익누계액", "retained": "이익잉여금"}
CF_COLS = "GHIJKLMNO"
GEN = "일반사항"


def _kor(i: int) -> str:
    return KOR[i] if i < len(KOR) else f"{i + 1}."


def _sum(cells: list[str]) -> str:
    return "+".join(cells) if cells else "0"


@dataclass
class Ctx:
    """시트 사이에 넘기는 위치 정보"""
    rows: list[TbRow]
    figs: dict[int, Figure]          # TB 행 -> 금액 (0인 줄을 빼는 데만 쓴다)
    bs: dict
    is_: dict
    general: bool                    # 일반사항 탭이 있으면 머리글을 수식으로
    ni_value: float = 0.0
    cf: dict = field(default_factory=dict)

    def cur(self, x: TbRow) -> float:
        f = self.figs.get(x.tb_row)
        return f.current if f else 0.0

    def prior(self, x: TbRow) -> float:
        f = self.figs.get(x.tb_row)
        return (f.prior or 0.0) if f else 0.0

    def nonzero(self, x: TbRow) -> bool:
        return abs(self.cur(x)) >= 0.5 or abs(self.prior(x)) >= 0.5


# ---------------------------------------------------------------- 머리글

def _report_head(g: Grid, ctx: Ctx, title: str) -> list[str]:
    g.put("A1", title, "ctitle")
    if ctx.general:
        g.put("A2", Formula(f"{GEN}!$D$6"), "bold")
    g.put("F2", "(단위 : 원)", "right")
    g.put("A3", "과목", "head")
    g.put("C3", "당기", "head")
    g.put("E3", "전기", "head")
    g.put("C4", "금          액", "head")
    g.put("E4", "금          액", "head")
    for ref in ("B3", "A4", "B4", "D3", "F3", "D4", "F4"):
        g.put(ref, None, "head")
    return ["A1:F1", "A3:B4", "C3:D3", "E3:F3", "C4:D4", "E4:F4"]


# ---------------------------------------------------------------- 정산표-CF

def _partner_rows(ctx: Ctx, word: str, partners) -> list[TbRow]:
    bs_rows = [x for x in ctx.rows if x.side != "is" and x.tb_row in ctx.bs["acc"] and not _is_cash(x)]
    if partners is None:
        return [x for x in bs_rows if x.key == "ppe" and not x.contra][:1]
    return [x for x in bs_rows if any(p in x.name for p in partners)]


def cf_sheet(st, ctx: Ctx) -> bytes:
    g = Grid(st)
    b, i_ = _q("수정후BS"), _q("수정후IS")
    if ctx.general:
        g.put("B1", Formula(f"{GEN}!$D$5"), "bold")
        g.put("B2", Formula(f"{GEN}!$D$6"), "bold")
        g.put("B3", Formula(f"{GEN}!$D$8"), "date")
        g.put("O2", "Date")
        g.put("P2", Formula(f"{GEN}!$D$12"), "date")
        g.put("O3", "Prepared by")
        g.put("P3", Formula(f"{GEN}!$D$9"))
        g.put("O4", "Reviewed by")
        g.put("P4", Formula(f"{GEN}!$D$10"))
    g.put("B4", "현금흐름표 검토", "title")
    head6 = {"A": "co\nde", "C": "과목", "D": "전기", "E": "당기", "F": "증감액", "G": "영업활동", "J": "투자활동",
             "L": "재무활동", "N": "비현금흐름", "P": "비  고"}
    head7 = {"D": "금액", "E": "금액", "G": "Dr(유출)", "H": "Cr(유입)", "I": "자산부채변동", "J": "Dr(유출)",
             "K": "Cr(유입)", "L": "Dr(유출)", "M": "Cr(유입)", "N": "Dr", "O": "Cr"}
    for col in "ABCDEFGHIJKLMNOP":
        g.put(f"{col}6", head6.get(col), "head")
        g.put(f"{col}7", head7.get(col), "head")
    merges = ["A6:A7", "C6:C7", "F6:F7", "G6:I6", "J6:K6", "L6:M6", "N6:O6", "P6:P7"]

    # 현금 없는 손익: (TB 행, 열) -> 수식 조각. 비용은 짝 계정의 H(Cr 유입), 수익은 G(Dr 유출)
    is_cells: dict[str, list[str]] = {}
    for word, partners, sign in NONCASH:
        cells = [f"{i_}!B{row}" for name, row in ctx.is_["acc"] if word in name]
        targets = _partner_rows(ctx, word, partners) if cells else []
        if targets:
            is_cells[word] = cells
    row_of: dict[int, int] = {}
    r = 8
    for kind, obj, _ in ctx.bs["order"]:
        if kind == "acc":
            row_of[obj.tb_row] = r
        r += 1
    noncash: dict[tuple[int, str], list[str]] = {}
    noncash_v: dict[tuple[int, str], float] = {}     # 같은 배분을 파이썬으로 (보고서에서 0인 줄을 빼는 데만)
    is_value = {name: 0.0 for name, _ in ctx.is_["acc"]}
    for x in ctx.rows:
        if x.side == "is":
            is_value[x.name] = is_value.get(x.name, 0.0) + ctx.cur(x)

    def change(x: TbRow) -> float:
        return ctx.cur(x) - ctx.prior(x)
    recon: list[tuple[str, str, list[str], list[int]]] = []   # (항목, 열, 손익계산서 셀, 배분한 행)
    for word, partners, sign in NONCASH:
        if word not in is_cells:
            continue
        col = "H" if sign > 0 else "G"
        targets = _partner_rows(ctx, word, partners)
        total = _sum(is_cells[word])
        total_v = sum(v for name, v in is_value.items() if word in name)
        if len(targets) == 1:
            noncash.setdefault((targets[0].tb_row, col), []).append(f"({total})")
            noncash_v[(targets[0].tb_row, col)] = noncash_v.get((targets[0].tb_row, col), 0.0) + total_v
        else:
            others, used = [], 0.0
            for x in targets[:-1]:
                rr = row_of[x.tb_row]
                inc = f"MAX(-F{rr},0)" if x.side == "asset" else f"MAX(F{rr},0)"
                noncash.setdefault((x.tb_row, col), []).append(inc)
                v = max(-change(x), 0.0) if x.side == "asset" else max(change(x), 0.0)
                noncash_v[(x.tb_row, col)] = noncash_v.get((x.tb_row, col), 0.0) + v
                used += v
                others.append(f"{col}{rr}")
            last = targets[-1]
            noncash.setdefault((last.tb_row, col), []).append(f"({total})-({'+'.join(others)})")
            noncash_v[(last.tb_row, col)] = noncash_v.get((last.tb_row, col), 0.0) + total_v - used
        recon.append((word, col, is_cells[word], [row_of[x.tb_row] for x in targets]))

    cash_rows: list[int] = []
    retained_rows: list[int] = []
    info: dict[int, tuple[TbRow, str]] = {}    # 정산표-CF 행 -> (계정, 분류)
    ni_row = None
    rest_v: dict[int, float] = {}
    r = 8
    for kind, obj, src in ctx.bs["order"]:
        if kind == "major":
            g.put(f"C{r}", obj, "bold")
        elif kind == "group":
            g.put(f"C{r}", f"  {obj.group_label}", "bold")
        elif kind == "ni":
            g.put(f"A{r}", 20, "text")
            g.put(f"C{r}", "    당기순이익(손실)")
            g.put(f"D{r}", Formula(f"N({b}!C{src})"), "num")
            g.put(f"E{r}", Formula(f"N({b}!B{src})"), "num")
            g.put(f"F{r}", Formula(f"E{r}-D{r}"), "num")
            g.put(f"I{r}", Formula(f"F{r}"), "input")
            g.put(f"P{r}", Formula(f"(H{r}+I{r}+K{r}+M{r}+O{r})-(G{r}+J{r}+L{r}+N{r})-F{r}"), "check")
            ni_row = r
        else:
            x: TbRow = obj
            asset = x.side == "asset"
            g.put(f"A{r}", 10 if asset else 20, "text")
            g.put(f"C{r}", f"    {x.label}")
            g.put(f"D{r}", Formula(f"N({b}!C{src})"), "num")
            g.put(f"E{r}", Formula(f"N({b}!B{src})"), "num")
            g.put(f"F{r}", Formula(f"E{r}-D{r}"), "num")
            g.put(f"P{r}", Formula(f"(G{r}+J{r}+L{r}+N{r})-(H{r}+I{r}+K{r}+M{r}+O{r})-F{r}" if asset else
                                   f"(H{r}+I{r}+K{r}+M{r}+O{r})-(G{r}+J{r}+L{r}+N{r})-F{r}"), "check")
            if _is_cash(x):
                cash_rows.append(r)
                info[r] = (x, "cash")
            else:
                for col in "GH":
                    parts = noncash.get((x.tb_row, col))
                    if parts:
                        g.put(f"{col}{r}", Formula("+".join(parts)), "input")
                # 현금 없는 손익을 뺀 나머지 (자산: 차변 순액, 부채·자본: 대변 순액)
                rest = f"(F{r}+H{r}-G{r})" if asset else f"(F{r}-H{r}+G{r})"
                cat = "fin" if x.key == "retained" else _cf_category(x)
                if x.key == "retained":
                    g.put(f"L{r}", Formula(f"-{rest}"), "input")     # 배당·이익처분 (전기말 이익잉여금 감소)
                    retained_rows.append(r)
                elif cat == "op":
                    g.put(f"I{r}", Formula(f"-{rest}" if asset else rest), "input")
                else:
                    dr, cr = {"inv": "JK", "fin": "LM", "other": "NO"}[cat]
                    g.put(f"{dr}{r}", Formula(f"MAX({rest},0)" if asset else f"MAX(-{rest},0)"), "input")
                    g.put(f"{cr}{r}", Formula(f"MAX(-{rest},0)" if asset else f"MAX({rest},0)"), "input")
                info[r] = (x, cat)
                h, gv = noncash_v.get((x.tb_row, "H"), 0.0), noncash_v.get((x.tb_row, "G"), 0.0)
                rest_v[r] = change(x) + h - gv if asset else change(x) - h + gv
        r += 1
    first, last = 8, r - 1
    s = r + 1
    g.put(f"C{s}", "합계", "bold")
    for col in "F" + CF_COLS + "P":
        g.put(f"{col}{s}", Formula(f"SUM({col}{first}:{col}{last})"), "boldnum")
    fl = s + 1
    g.put(f"C{fl}", "활동별 현금흐름", "bold")
    g.put(f"G{fl}", "영업 = G+H+I", "right")
    g.put(f"I{fl}", Formula(f"H{s}+I{s}-G{s}"), "boldnum")
    g.put(f"K{fl}", Formula(f"K{s}-J{s}"), "boldnum")
    g.put(f"M{fl}", Formula(f"M{s}-L{s}"), "boldnum")
    g.put(f"O{fl}", "현금 증감", "right")
    g.put(f"P{fl}", Formula(f"I{fl}+K{fl}+M{fl}"), "boldnum")
    chk = fl + 2
    g.put(f"C{chk}", "검증 (모두 0이어야 함)", "bold")
    cash_f = _sum([f"F{x}" for x in cash_rows])
    cash_p = _sum([f"P{x}" for x in cash_rows])
    checks = [("비현금흐름 Dr - Cr", f"N{s}-O{s}"),
              ("덜 나눈 금액 (비고 합계 - 현금 계정 비고)", f"P{s}-({cash_p})"),
              ("현금흐름 합계 - 현금 증감(수정후BS)", f"P{fl}-({cash_f})")]
    for k, (label, f) in enumerate(checks, start=1):
        g.put(f"C{chk + k}", label)
        g.put(f"F{chk + k}", Formula(f"ROUND({f},0)"), "check")
    note = chk + len(checks) + 1
    g.put(f"C{note}", "노란 칸은 자동으로 나눈 수식입니다. 취득·처분을 따로 보이려면 Dr·Cr 칸에 총액을 적고, "
                      "비고가 0인지 확인하세요. 현금으로 본 계정: "
                      + (", ".join(info[x][0].name for x in cash_rows) or "없음"))
    rc = note + 2
    g.put(f"C{rc}", "현금 없는 손익 대사", "bold")
    g.row(rc, ["구분", "손익계산서", "정산표-CF 배분", "차액"], "head", start=6)
    recon_row: dict[str, int] = {}
    for k, (word, col, cells, trows) in enumerate(recon, start=1):
        rr = rc + k
        g.put(f"G{rr}", word)
        g.put(f"H{rr}", Formula(_sum(cells)), "num")
        g.put(f"I{rr}", Formula(_sum([f"{col}{t}" for t in trows])), "num")
        g.put(f"J{rr}", Formula(f"ROUND(I{rr}-H{rr},0)"), "check")
        recon_row[word] = rr
    widths = [4, 3, 30, 15, 15, 15] + [14] * 9 + [13]
    ctx.cf = {"info": info, "cash": cash_rows, "retained": retained_rows, "ni": ni_row, "sum": s, "flow": fl,
              "recon": recon_row, "recon_col": {w: c for w, c, _, _ in recon},
              "rest": rest_v}
    return grid_xml(g.cells, widths, freeze="D8", merges=merges)


# ---------------------------------------------------------------- 보고서 BS

def bs_report(st, ctx: Ctx) -> bytes:
    g = Grid(st)
    merges = _report_head(g, ctx, "재 무 상 태 표")
    b = _q("수정후BS")
    bs = ctx.bs
    accs = [(obj, src) for kind, obj, src in bs["order"] if kind == "acc"]
    retained = [(x, src) for x, src in accs if x.key == "retained"]
    ni_host = None
    if bs.get("ni") and retained:
        ni_host = next((x for x, _ in retained if "미처분" in x.name), None) or \
            next((x for x, _ in retained if "이월" in x.name), None) or retained[-1][0]
    r = 5
    totals: dict[str, int] = {}

    def acc_line(x: TbRow, src: int) -> bool:
        nonlocal r
        host = ni_host is not None and x.tb_row == ni_host.tb_row
        if not ctx.nonzero(x) and not host:
            return False
        g.put(f"B{r}", ("    " if x.contra else "  ") + x.name)
        g.put(f"C{r}", Formula(f"{b}!B{src}" + (f"+{b}!B{bs['ni']}" if host else "")), "num")
        g.put(f"E{r}", Formula(f"N({b}!C{src})"), "num")
        r += 1
        return True

    def block(num: str, label: str, items: list[tuple[TbRow, int]]) -> int:
        """번호·이름 줄 + 계정 줄. 소계는 D·F열. 반환: 소계 행"""
        nonlocal r
        head = r
        g.put(f"A{head}", num, "bold")
        g.put(f"B{head}", label, "bold")
        r += 1
        start = r
        for x, src in items:
            acc_line(x, src)
        for inner, outer in (("C", "D"), ("E", "F")):
            g.put(f"{outer}{head}", Formula(f"SUM({inner}{start}:{inner}{r - 1})") if r > start else 0, "boldnum")
        return head

    def side(title: str, majors: list[str], total_label: str):
        nonlocal r
        g.put(f"B{r}", title, "bold")
        r += 1
        heads = []
        if title == "자 본":
            n = 0
            for key in EQUITY_ORDER:
                items = [(x, s) for x, s in accs if x.major == "자본" and x.key == key]
                if not items:
                    continue
                heads.append(block(ROMAN[n], EQUITY_NAME[key], items))
                n += 1
        else:
            for n, major in enumerate([m for m in majors if any(x.major == m for x, _ in accs)]):
                items = [(x, s) for x, s in accs if x.major == major]
                if major in SUBCLASS_ORDER:
                    mh = r
                    g.put(f"A{mh}", ROMAN[n], "bold")
                    g.put(f"B{mh}", major, "bold")
                    r += 1
                    subs = []
                    k = 0
                    for sub in SUBCLASS_ORDER[major]:
                        si = [(x, s) for x, s in items if SUBCLASS.get(x.key) == sub]
                        if not si:
                            continue
                        subs.append(block(f"({k + 1})", sub, si))
                        k += 1
                    for col in "DF":
                        g.put(f"{col}{mh}", Formula(_sum([f"{col}{h}" for h in subs])), "boldnum")
                    heads.append(mh)
                else:
                    heads.append(block(ROMAN[n], major, items))
        g.put(f"B{r}", total_label, "bold")
        for col in "DF":
            g.put(f"{col}{r}", Formula(_sum([f"{col}{h}" for h in heads])), "boldnum")
        totals[total_label] = r
        r += 2

    side("자 산", ["유동자산", "비유동자산"], "자산총계")
    side("부 채", ["유동부채", "비유동부채"], "부채총계")
    side("자 본", [], "자본총계")
    r -= 1
    g.put(f"B{r}", "부채와자본총계", "bold")
    for col in "DF":
        g.put(f"{col}{r}", Formula(f"{col}{totals['부채총계']}+{col}{totals['자본총계']}"), "boldnum")
    totals["부채와자본총계"] = r
    r += 2
    g.put(f"B{r}", "검증 (0이어야 함)", "bold")
    t = bs["totals"]
    checks = [("자산총계 - 부채와자본총계", "D", "F", f"{{c}}{totals['자산총계']}-{{c}}{totals['부채와자본총계']}"),
              ("자산총계 - 수정후BS", "D", "F", f"{{c}}{totals['자산총계']}-{b}!{{s}}{t['자산총계']}"),
              ("부채와자본총계 - 수정후BS", "D", "F", f"{{c}}{totals['부채와자본총계']}-{b}!{{s}}{t['부채와자본총계']}")]
    for k, (label, cur, pri, f) in enumerate(checks, start=1):
        g.put(f"B{r + k}", label)
        g.put(f"D{r + k}", Formula("ROUND(" + f.format(c="D", s="B") + ",0)"), "check")
        g.put(f"F{r + k}", Formula("ROUND(" + f.format(c="F", s="C") + ",0)"), "check")
    return grid_xml(g.cells, [4, 30, 16, 16, 16, 16], freeze="A5", merges=merges)


# ---------------------------------------------------------------- 보고서IS

def is_report(st, ctx: Ctx) -> bytes:
    g = Grid(st)
    merges = _report_head(g, ctx, "손 익 계 산 서")
    i_ = _q("수정후IS")
    by_key = ctx.is_["by_key"]
    sec: dict[str, int] = {}
    r = 5

    def block(num: str, label: str, keys: list[str]):
        nonlocal r
        head = r
        g.put(f"A{head}", num, "bold")
        g.put(f"B{head}", label, "bold")
        r += 1
        start = r
        for key in keys:
            for x, src in by_key.get(key, []):
                if not ctx.nonzero(x):
                    continue
                g.put(f"B{r}", "  " + x.name + (" (제조원가)" if x.key in MFG_KEYS else ""))
                g.put(f"C{r}", Formula(f"{i_}!B{src}"), "num")
                g.put(f"E{r}", Formula(f"{i_}!C{src}"), "num")
                r += 1
        for inner, outer in (("C", "D"), ("E", "F")):
            g.put(f"{outer}{head}", Formula(f"SUM({inner}{start}:{inner}{r - 1})") if r > start else 0, "boldnum")
        return head

    def calc(num: str, label: str, f: str):
        nonlocal r
        g.put(f"A{r}", num, "bold")
        g.put(f"B{r}", label, "bold")
        for col in "DF":
            g.put(f"{col}{r}", Formula(f.replace("#", col)), "boldnum")
        r += 1
        return r - 1

    sec["sales"] = block("Ⅰ.", "매출액", ["sales"])
    sec["cogs"] = block("Ⅱ.", "매출원가", ["cogs"])
    sec["gross"] = calc("Ⅲ.", "매출총이익(손실)", f"#{sec['sales']}-#{sec['cogs']}")
    sec["sga"] = block("Ⅳ.", "판매비와관리비", ["sga"])
    sec["op"] = calc("Ⅴ.", "영업이익(손실)", f"#{sec['gross']}-#{sec['sga']}")
    sec["nonop_inc"] = block("Ⅵ.", "영업외수익", ["nonop_inc"])
    sec["nonop_exp"] = block("Ⅶ.", "영업외비용", ["nonop_exp"])
    sec["ebt"] = calc("Ⅷ.", "법인세비용차감전순이익(손실)", f"#{sec['op']}+#{sec['nonop_inc']}-#{sec['nonop_exp']}")
    sec["tax"] = block("Ⅸ.", "법인세비용", ["tax"])
    sec["ni"] = calc("Ⅹ.", "당기순이익(손실)", f"#{sec['ebt']}-#{sec['tax']}")
    r += 1
    ni = ctx.is_["sec"]["ni"]
    g.put(f"B{r}", "검증: 당기순이익 - 수정후IS (0이어야 함)", "bold")
    g.put(f"D{r}", Formula(f"ROUND(D{sec['ni']}-{i_}!B{ni},0)"), "check")
    g.put(f"F{r}", Formula(f"ROUND(F{sec['ni']}-{i_}!C{ni},0)"), "check")
    return grid_xml(g.cells, [4, 30, 16, 16, 16, 16], freeze="A5", merges=merges)


# ---------------------------------------------------------------- 자본변동표

def equity_sheet(st, ctx: Ctx) -> bytes:
    g = Grid(st)
    b, i_ = _q("수정후BS"), _q("수정후IS")
    groups: dict[str, list[int]] = {}
    for kind, obj, src in ctx.bs["order"]:
        if kind == "group" and obj.major == "자본":
            groups.setdefault(obj.key, []).append(src)
    keys = [k for k in EQUITY_ORDER if k in groups] + [k for k in groups if k not in EQUITY_ORDER]
    cols = [chr(ord("B") + i) for i in range(len(keys))]
    tot = chr(ord("B") + len(keys))
    g.put("A1", "자 본 변 동 표", "ctitle")
    if ctx.general:
        g.put("A2", Formula(f"{GEN}!$D$6"), "bold")
    g.put(f"{tot}2", "(단위 : 원)", "right")
    g.put("A3", "과         목", "head")
    for c, k in zip(cols, keys):
        g.put(f"{c}3", EQUITY_NAME.get(k, k), "head")
    g.put(f"{tot}3", "총         계", "head")

    def label(text: str, when: str | None) -> object:
        if ctx.general and when:
            return Formula(f'TEXT({when},"yyyy.mm.dd")&"({text})"')
        return text

    prior_end = f"{GEN}!$D$7"
    rows_def = [
        (4, label("전기초", f"EDATE({prior_end},-12)+1"), "begin_p"),
        (5, "당기순이익(손실)", "ni_p"),
        (6, "배당금 지급 등 (입력)", "other_p"),
        (7, label("전기말", prior_end), "end_p"),
        (8, label("당기초", f"{prior_end}+1"), "begin_c"),
        (9, "당기순이익(손실)", "ni_c"),
        (10, "배당금 지급 등 (이익잉여금 기타 변동)", "div_c"),
        (11, "그 밖의 자본 변동 (증자·자기주식 등)", "other_c"),
        (12, label("당기말", f"{GEN}!$D$8"), "end_c"),
    ]
    ni_is = ctx.is_["sec"]["ni"]
    for r, text, kind in rows_def:
        g.put(f"A{r}", text, "bold" if kind.startswith(("begin", "end")) else "text")
        for c, k in zip(cols, keys):
            cur = "+".join(f"{b}!B{h}" for h in groups[k])
            pri = "+".join(f"N({b}!C{h})" for h in groups[k])
            ret = k == "retained"
            f = {"begin_p": f"{c}7-{c}5-{c}6",
                 "ni_p": f"N({i_}!C{ni_is})" if ret else "0",
                 "other_p": None,
                 "end_p": pri,
                 "begin_c": f"{c}7",
                 "ni_c": f"{i_}!B{ni_is}" if ret else "0",
                 "div_c": f"({cur})-{c}8-{c}9" if ret else "0",
                 "other_c": "0" if ret else f"({cur})-{c}8",
                 "end_c": f"SUM({c}8:{c}11)"}[kind]
            if f is None:
                g.put(f"{c}{r}", 0, "input")
            else:
                g.put(f"{c}{r}", Formula(f), "boldnum" if kind.startswith(("begin", "end")) else "num")
        if cols:
            g.put(f"{tot}{r}", Formula(f"SUM({cols[0]}{r}:{cols[-1]}{r})"),
                  "boldnum" if kind.startswith(("begin", "end")) else "num")
    g.put("A14", "검증: 당기말 - 수정후BS (0이어야 함)", "bold")
    for c, k in zip(cols, keys):
        cur = "+".join(f"{b}!B{h}" for h in groups[k])
        g.put(f"{c}14", Formula(f"ROUND({c}12-({cur}),0)"), "check")
    g.put("A15", "전기 배당금 등은 노란 칸에 적어 주세요. 전기초는 전기말에서 거꾸로 계산합니다.")
    merges = [f"A1:{tot}1"]
    return grid_xml(g.cells, [34] + [16] * (len(keys) + 1), freeze="B4", merges=merges)


# ---------------------------------------------------------------- 보고서현금흐름표

def cf_report(st, ctx: Ctx) -> bytes:
    g = Grid(st)
    merges = _report_head(g, ctx, "현 금 흐 름 표")
    g.put("G3", "검증(0)", "head")
    cf = _q("정산표-CF")
    i_ = _q("수정후IS")
    c = ctx.cf
    s, fl = c["sum"], c["flow"]
    info = c["info"]
    rest = c["rest"]     # 계정별로 활동에 나눌 금액 (자산: 차변 순액, 부채·자본: 대변 순액)
    r = 5

    def line(num: str, label: str, col: str, f: str, style: str = "num"):
        nonlocal r
        g.put(f"A{r}", num, "bold" if style == "boldnum" else "text")
        g.put(f"B{r}", label, "bold" if style == "boldnum" else "text")
        g.put(f"{col}{r}", Formula(f), style)
        r += 1
        return r - 1

    def check(row: int, f: str):
        g.put(f"G{row}", Formula(f"ROUND({f},0)"), "check")

    def detail_block(num: str, label: str, items: list[tuple[str, str]], negate: bool, check_f: str | None):
        nonlocal r
        head = r
        r += 1
        start = r
        for k, (text, f) in enumerate(items):
            line(_kor(k), text, "C", f)
        body = f"SUM(C{start}:C{r - 1})" if r > start else "0"
        g.put(f"A{head}", num, "bold")
        g.put(f"B{head}", label, "bold")
        g.put(f"C{head}", Formula(f"-{body}" if negate else body), "boldnum")
        if check_f:
            check(head, check_f.replace("@", f"C{head}"))
        return head

    op = r
    g.put(f"A{op}", "Ⅰ.", "bold")
    g.put(f"B{op}", "영업활동으로 인한 현금흐름", "bold")
    r += 1
    parts = []
    parts.append(line("1.", "당기순이익(손실)", "C", f"{i_}!B{ctx.is_['sec']['ni']}"))
    exp = [(w, f"{cf}!I{rr}") for w, rr in c["recon"].items() if c["recon_col"][w] == "H"]
    inc = [(w, f"{cf}!I{rr}") for w, rr in c["recon"].items() if c["recon_col"][w] == "G"]
    parts.append(detail_block("2.", "현금의 유출이 없는 비용등의 가산", exp, False, f"@-{cf}!H{s}"))
    parts.append(detail_block("3.", "현금의 유입이 없는 수익등의 차감", inc, True, f"-@-{cf}!G{s}"))
    ops = []
    for row, (x, cat) in info.items():
        if cat != "op":
            continue
        if abs(rest[row]) < 0.5:
            continue
        word = "감소(증가)" if x.side == "asset" else "증가(감소)"
        ops.append((f"{x.label}의 {word}", f"{cf}!I{row}"))
    ni_i = f"-{cf}!I{c['ni']}" if c["ni"] else ""
    parts.append(detail_block("4.", "영업활동으로 인한 자산부채의 변동", ops, False, f"@-({cf}!I{s}{ni_i})"))
    g.put(f"D{op}", Formula("+".join(f"C{p}" for p in parts)), "boldnum")
    check(op, f"D{op}-{cf}!I{fl}")

    def activity(num: str, label: str, cr: str, dr: str, cat: str, fl_col: str, extra_out=None):
        nonlocal r
        head = r
        g.put(f"A{head}", num, "bold")
        g.put(f"B{head}", label, "bold")
        r += 1
        ins, outs = [], []
        for row, (x, kind) in info.items():
            if kind != cat or x.key == "retained" or abs(rest[row]) < 0.5:
                continue
            asset = x.side == "asset"
            # 나눌 금액의 방향으로 고른다 (자산 감소·부채 증가 = 유입)
            inflow = (rest[row] < 0) if asset else (rest[row] > 0)
            n = x.label
            if asset:
                text_in, text_out = f"{n}의 감소", f"{n}의 증가"
            elif x.key in ("borrow", "lt_borrow"):
                text_in, text_out = f"{n}의 차입", f"{n}의 상환"
            else:
                text_in, text_out = f"{n}의 증가", f"{n}의 감소"
            if inflow:
                ins.append((text_in, f"{cf}!{cr}{row}"))
            else:
                outs.append((text_out, f"{cf}!{dr}{row}"))
        if extra_out:
            outs += extra_out
        a = detail_block("1.", label.replace("현금흐름", "현금유입액"), ins, False, f"@-{cf}!{cr}{s}")
        b2 = detail_block("2.", label.replace("현금흐름", "현금유출액"), outs, True, f"-@-{cf}!{dr}{s}")
        g.put(f"D{head}", Formula(f"C{a}+C{b2}"), "boldnum")
        check(head, f"D{head}-{cf}!{fl_col}{fl}")
        return head

    inv = activity("Ⅱ.", "투자활동으로 인한 현금흐름", "K", "J", "inv", "K")
    div = []
    if any(abs(rest[x]) >= 0.5 for x in c["retained"]):
        div = [("배당금의 지급 등 (이익잉여금 감소)", "+".join(f"{cf}!L{x}" for x in c["retained"]))]
    fin = activity("Ⅲ.", "재무활동으로 인한 현금흐름", "M", "L", "fin", "M", div)
    tot = line("Ⅳ.", "현금의 증가(감소)(Ⅰ+Ⅱ+Ⅲ)", "D", f"D{op}+D{inv}+D{fin}", "boldnum")
    check(tot, f"D{tot}-{cf}!P{fl}")
    beg = line("Ⅴ.", "기초의 현금", "D", _sum([f"{cf}!D{x}" for x in c["cash"]]), "boldnum")
    end = line("Ⅵ.", "기말의 현금", "D", f"D{tot}+D{beg}", "boldnum")
    check(end, f"D{end}-({_sum([f'{cf}!E{x}' for x in c['cash']])})")
    r += 1
    g.put(f"B{r}", "전기(E·F열)는 전기 감사보고서 현금흐름표를 옮겨 적어 주세요. G열 검증은 모두 0이어야 합니다.")
    other = [x.name for x, kind in info.values() if kind == "other" and ctx.nonzero(x)]
    if other:
        g.put(f"B{r + 1}", "비현금흐름으로 둔 계정 (현금 거래가 있으면 정산표-CF에서 옮겨 주세요): " + ", ".join(other))
    return grid_xml(g.cells, [4, 40, 16, 16, 16, 16, 12], freeze="A5", merges=merges)


def build(st, rows: list[TbRow], figs: list[Figure], bs: dict, is_: dict, general: bool) -> list[tuple[str, bytes]]:
    """반환: [(탭 이름, 시트 XML)] (TABS 순서)"""
    ctx = Ctx(rows, {f.row.tb_row: f for f in figs}, bs, is_, general)
    cf = cf_sheet(st, ctx)       # 먼저 만들어야 보고서현금흐름표가 행 위치를 안다
    return list(zip(TABS, (cf, bs_report(st, ctx), is_report(st, ctx), equity_sheet(st, ctx), cf_report(st, ctx))))
