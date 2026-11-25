"""합계잔액시산표로 조서의 TB·BS·IS 탭을 채워 리드시트(○1000)가 저절로 채워지게 하고,
거래처원장으로 ○1100 탭에 거래처별 명세를 쓴다.

조서 양식(2026-11 받은 '조서_작업용' 기준):
- TB 탭: B 조서 기호, D 차변잔액, E 차변합계, F 계정과목, G 대변합계, H 대변잔액,
  I 기초(=BS!F), J 증가, K 감소, L 기말, M 검증. 묶음 머리 행의 D열이 =SUM(D12:D16) 꼴이고
  그 범위가 계정을 넣을 칸이다. '[유동자산]' 같은 대괄호 줄이 큰 구분이다.
- BS 탭: C열이 =TB!C(행) 으로 TB 칸과 이어져 있다. D 당기, F 전기.
- IS 탭: 묶음 머리 행의 E열(또는 D열)이 =SUM(D6:D10) 꼴. C 과목, D 당기, F 전기.
- 리드시트(A1000 등)는 TB·IS 탭을 수식으로 끌어 쓰므로 위 세 탭만 채우면 된다.

양식은 바꾸지 않고 값만 넣는다. 도형이 사라지지 않게 XML 수준에서 고친다 (xlsx_xml.py).
재무제표를 대신 작성하는 것이 아니라 회사가 준 시산표를 조서에 옮겨 적는 일이다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from ..importers import books as books_mod
from ..importers.common import squash
from ..importers.douzone import section_label
from .xlsx_xml import Formula, XlsxPackage

# 계정 묶음 -> (조서 TB 탭의 큰 구분, 묶음 이름 앞부분)
TARGETS = {
    "cash": ("유동자산", "현금"), "sec": ("유동자산", "유가증권"), "ar": ("유동자산", "매출채권"),
    "other_quick": ("유동자산", "기타당좌자산"), "inv_cur": ("유동자산", "투자자산"),
    "tax_asset": ("유동자산", "법인세"), "inventory": ("유동자산", "재고자산"),
    "lt_cash": ("비유동자산", "현금"), "lt_ar": ("비유동자산", "매출채권"), "lt_sec": ("비유동자산", "유가증권"),
    "equity_method": ("비유동자산", "지분법"), "other_inv": ("비유동자산", "기타투자자산"),
    "ppe": ("비유동자산", "유형자산"), "intangible": ("비유동자산", "무형자산"),
    "other_nca": ("비유동자산", "기타비유동자산"),
    "ap": ("유동부채", "매입채무"), "borrow": ("유동부채", "차입금"), "other_cl": ("유동부채", "기타유동부채"),
    "lt_ap": ("비유동부채", "매입채무"), "lt_borrow": ("비유동부채", "차입금"),
    "retire": ("비유동부채", "퇴직급여"), "provision": ("비유동부채", "제충당부채"),
    "other_ncl": ("비유동부채", "기타비유동부채"),
    "capital": ("자본", "자본금"), "cap_surplus": ("자본", "자본잉여금"), "cap_adj": ("자본", "자본조정"),
    "aoci": ("자본", "기타포괄"), "retained": ("자본", "이익잉여금"),
    "sales": ("매출", "매출"), "cogs": ("매출원가", "매출원가"), "material": ("매출원가", "재료비"),
    "labor": ("매출원가", "노무비"), "overhead": ("매출원가", "경비"), "sga": ("판매관리비", "판매관리비"),
    "nonop_inc": ("영업외수익", "영업외수익"), "nonop_exp": ("영업외비용", "영업외비용"),
    "tax": ("법인세비용", "법인세비용"),
}
CREDIT_NATURE = {"ap", "borrow", "other_cl", "lt_ap", "lt_borrow", "retire", "provision", "other_ncl",
                 "capital", "cap_surplus", "cap_adj", "aoci", "retained", "sales", "nonop_inc"}
BS_SIDES = {"유동자산": "asset", "비유동자산": "asset", "유동부채": "liab", "비유동부채": "liab", "자본": "liab"}
ASSET_SUBS = {"당좌자산": "유동자산", "재고자산": "유동자산", "투자자산": "비유동자산", "유형자산": "비유동자산",
              "무형자산": "비유동자산", "기타비유동자산": "비유동자산"}
COST_SUBS = {"재료비", "노무비", "제조경비", "경비"}
CONTRA_WORDS = ("대손충당금", "감가상각누계액", "상각누계액", "손상차손누계액", "평가충당금", "국고보조금",
                "현재가치할인차금", "할인발행차금")
# IS 탭 묶음 이름 -> 계정 묶음
IS_TARGETS = {"매출액": "sales", "제품매출원가": "cogs_product", "상품매출원가": "cogs_goods",
              "기타매출원가": "cogs_other", "판매비와관리비": "sga", "영업외수익": "nonop_inc",
              "영업외비용": "nonop_exp", "법인세비용": "tax"}


# ---------------------------------------------------------------- 조서 양식 읽기

@dataclass
class Group:
    row: int
    letter: str
    label: str
    major: str
    slots: list[int]

    @property
    def side(self) -> str:
        return BS_SIDES.get(self.major, "is")


@dataclass
class Template:
    groups: list[Group]
    bs_row: dict[int, int]                      # TB 행 -> BS 행
    is_groups: dict[str, list[int]]             # IS 묶음 -> 칸 행
    lead_rows: dict[str, list[tuple[int, int]]]  # 리드시트 -> [(리드 행, TB 행)]
    schedule_tabs: list[str]                    # ○1100 탭 이름
    styles: dict[str, str | None] = field(default_factory=dict)
    table_head: dict[str, str] = field(default_factory=dict)  # ○1100 탭 -> 'No.' 머리글 셀 (서식 본보기)
    missing_from_total: list[Group] = field(default_factory=list)  # TB 합계 수식에 빠진 묶음

    def group(self, key: str) -> Group | None:
        major, prefix = TARGETS[key]
        return next((g for g in self.groups if g.major == major and g.label.startswith(prefix)), None)


_SUM_RE = re.compile(r"^=SUM\(\$?([A-Z]+)\$?(\d+):\$?[A-Z]+\$?(\d+)\)$")


def _label(v) -> str:
    return section_label(v) if v is not None else ""


def read_template(path: str | Path) -> Template:
    wb = openpyxl.load_workbook(path)
    try:
        if not {"TB", "BS", "IS"} <= set(wb.sheetnames):
            raise ValueError("조서에 TB·BS·IS 탭이 없습니다. 조서 양식 파일을 넣어 주세요.")
        tb = wb["TB"]
        groups, major, d_formula, total_row = [], "", {}, None
        for r in range(1, tb.max_row + 1):
            f = tb.cell(r, 6).value
            text = squash(f) if isinstance(f, str) else ""
            if text.startswith("[") and text.endswith("]"):
                major = _label(text)
            if text == "합계":
                total_row = r
            d = tb.cell(r, 4).value
            if isinstance(d, str) and d.startswith("="):
                d_formula[r] = d
            m = _SUM_RE.match(d) if isinstance(d, str) else None
            if m and m.group(1) == "D":
                b = tb.cell(r, 2).value
                groups.append(Group(r, str(b).strip() if b else "", _label(f), major,
                                    list(range(int(m.group(2)), int(m.group(3)) + 1))))
        bs_row = {}
        bs = wb["BS"]
        for r in range(1, bs.max_row + 1):
            c = bs.cell(r, 3).value
            m = re.fullmatch(r"=TB!\$?C\$?(\d+)", c) if isinstance(c, str) else None
            if m:
                bs_row[int(m.group(1))] = r
        is_groups: dict[str, list[int]] = {}
        ws = wb["IS"]
        for r in range(1, ws.max_row + 1):
            key = IS_TARGETS.get(_label(ws.cell(r, 3).value))
            for col in (5, 4):
                v = ws.cell(r, col).value
                m = _SUM_RE.match(v) if isinstance(v, str) else None
                if key and m:
                    is_groups[key] = list(range(int(m.group(2)), int(m.group(3)) + 1))
                    break
        lead_rows: dict[str, list[tuple[int, int]]] = {}
        tabs, heads = [], {}
        for name in wb.sheetnames:
            if re.fullmatch(r"[A-Z,]+1100", name):
                tabs.append(name)
                for row in wb[name].iter_rows(min_row=9, max_row=120, max_col=3):
                    for c in row:
                        if isinstance(c.value, str) and c.value.strip() == "No.":
                            heads.setdefault(name, c.coordinate)
            if not re.fullmatch(r"[A-Z,]+1000", name):
                continue
            s = wb[name]
            for r in range(1, min(s.max_row, 200) + 1):
                c = s.cell(r, 3).value
                m = re.fullmatch(r"=TB!\$?C\$?(\d+)", c) if isinstance(c, str) else None
                if m:
                    lead_rows.setdefault(name, []).append((r, int(m.group(1))))
        missing = []
        if total_row and total_row in d_formula:
            seen, todo = set(), [total_row]
            while todo:
                r = todo.pop()
                for ref in re.findall(r"D(\d+)", d_formula.get(r, "")):
                    if int(ref) not in seen and not _SUM_RE.match(d_formula.get(r, "")):
                        seen.add(int(ref))
                        todo.append(int(ref))
            missing = [g for g in groups if g.row not in seen]
        return Template(groups, bs_row, is_groups, lead_rows, tabs, table_head=heads, missing_from_total=missing)
    finally:
        wb.close()


# ---------------------------------------------------------------- 계정 나누기

def is_contra(name: str) -> bool:
    return any(w in name for w in CONTRA_WORDS)


def classify(major: str, sub: str | None, name: str) -> str | None:
    """시산표의 구분(큰 구분·작은 구분)과 과목명으로 조서 묶음을 정한다. 모르면 None."""
    def has(*words):
        return any(w in name for w in words)

    if sub == "재고자산":
        return "inventory"
    if major == "유동자산":
        if has("법인세"):
            return "tax_asset"
        if has("현금", "예금", "적금", "금융상품"):
            return "cash"
        if has("단기매매", "유가증권", "단기투자증권", "당기손익"):
            return "sec"
        if has("외상매출금", "받을어음", "매출채권", "공사미수금"):
            return "ar"
        return "other_quick"
    if major == "비유동자산":
        if sub == "유형자산":
            return "ppe"
        if sub == "무형자산":
            return "intangible"
        if sub == "투자자산":
            if has("장기금융상품", "장기성예금", "특정현금", "예금"):
                return "lt_cash"
            if has("장기매출채권", "장기외상매출금", "장기받을어음", "장기성매출채권"):
                return "lt_ar"
            if has("지분법"):
                return "equity_method"
            if has("매도가능", "만기보유", "투자증권", "공정가치", "상각후원가"):
                return "lt_sec"
            return "other_inv"
        return "other_nca"
    if major == "유동부채":
        if has("외상매입금", "지급어음", "매입채무"):
            return "ap"
        if has("차입금", "유동성장기", "사채"):
            return "borrow"
        return "other_cl"
    if major == "비유동부채":
        if has("장기외상매입금", "장기지급어음", "장기매입채무"):
            return "lt_ap"
        if has("차입금", "사채"):
            return "lt_borrow"
        if has("퇴직", "연금"):
            return "retire"
        if has("충당부채"):
            return "provision"
        return "other_ncl"
    if major.startswith("자본") or major in ("이익잉여금", "기타포괄손익누계액", "자본"):
        if major == "자본금" or (major == "자본" and has("자본금")):
            return "capital"
        if major == "자본잉여금" or has("주식발행초과금", "감자차익"):
            return "cap_surplus"
        if major == "자본조정":
            return "cap_adj"
        if major.startswith("기타포괄") or has("평가이익", "평가손실", "재평가"):
            return "aoci"
        if major == "이익잉여금" or has("잉여금", "준비금", "결손금"):
            return "retained"
        return "cap_adj"
    if major.startswith("기타포괄"):
        return "aoci"
    if major == "이익잉여금":
        return "retained"
    if major.startswith("매출원가"):
        return "cogs"
    if major.startswith("매출"):
        return "sales"
    if major == "제조원가":
        if sub == "재료비" or has("원재료", "부재료", "재료비"):
            return "material"
        if sub == "노무비" or (sub is None and has("급여", "임금", "상여", "수당", "퇴직")):
            return "labor"
        return "overhead"
    if "판매비" in major or "관리비" in major:
        return "sga"
    if major == "영업외수익":
        return "nonop_inc"
    if major == "영업외비용":
        return "nonop_exp"
    plain = re.sub(r"\(.*?\)", "", major)
    if "법인세" in plain or "소득세" in plain:
        return "tax"
    return None


@dataclass
class Account:
    code: str | None
    name: str
    key: str | None
    contra: bool
    parent: str | None
    db: float = 0.0       # 차변잔액
    dt: float = 0.0       # 차변합계
    ct: float = 0.0       # 대변합계
    cb: float = 0.0       # 대변잔액
    prior: float | None = None
    tb_row: int | None = None

    @property
    def credit(self) -> bool:
        """대변 성격 계정인지 (대변잔액이 있으면 대변, 없으면 묶음 성격과 차감계정 여부로)"""
        if self.cb and not self.db:
            return True
        if self.db and not self.cb:
            return False
        natural = self.key in CREDIT_NATURE
        return natural != self.contra

    @property
    def fs_key(self) -> str:
        """재무상태표에서 쓰는 이름 (차감계정은 '대손충당금(외상매출금)')"""
        return f"{self.name}({self.parent})" if self.contra and self.parent else self.name


def accounts_from_tb(tb: list[dict]) -> list[Account]:
    major, sub, out = "", None, []
    last: Account | None = None
    for r in tb:
        if r["level"] == "group":
            n = section_label(r["name"])
            if n in ASSET_SUBS:
                major, sub = ASSET_SUBS[n], n
            elif n in COST_SUBS and major == "제조원가":
                sub = n
            else:
                major, sub = n, None
            last = None
            continue
        if r["level"] != "account":
            continue
        contra = is_contra(r["name"])
        key = last.key if contra and last is not None else classify(major, sub, r["name"])
        acc = Account(r["code"], r["name"], key, contra, last.name if contra and last else None,
                      r["debit_balance"], r["debit_total"], r["credit_total"], r["credit_balance"])
        out.append(acc)
        if not contra:
            last = acc
    return out


def _bs_sections(bs: list[dict]) -> list[tuple[str, str | None, dict]]:
    """재무상태표 행마다 (큰 구분, 작은 구분, 행)."""
    major, sub, out = "", None, []
    for r in bs:
        if not r["is_account"]:
            n = section_label(r["name"])
            if n in ASSET_SUBS:
                major, sub = ASSET_SUBS[n], n
            elif n in BS_SIDES or n in ("자본금", "자본잉여금", "자본조정", "기타포괄손익누계액", "이익잉여금"):
                major, sub = n, None
            elif n in ("자산", "부채"):
                major, sub = "", None
            continue
        out.append((major, sub, r))
    return out


def attach_prior(accounts: list[Account], bs: list[dict], notes: list[str]) -> None:
    """재무상태표(과목별)의 전기 금액을 붙인다. 시산표에 없는 계정은 전기만 있는 계정으로 더한다."""
    if not bs:
        notes.append("재무상태표(과목별) 파일이 없어 전기 금액(BS 탭 F열)은 비워 두었습니다. "
                     "이 경우 TB 탭 증가 칸에 전기이월이 함께 들어갑니다.")
        return
    by_name: dict[str, list[tuple[str, str | None, dict]]] = {}
    by_code: dict[str, list[tuple[str, str | None, dict]]] = {}
    for major, sub, r in _bs_sections(bs):
        by_name.setdefault(r["name"], []).append((major, sub, r))
        if r.get("code"):
            by_code.setdefault(r["code"], []).append((major, sub, r))
    used = set()
    for a in accounts:
        if a.key is None or a.key not in TARGETS or TARGETS[a.key][0] not in BS_SIDES:
            continue
        hits = (by_code.get(a.code) if a.code else None) or by_name.get(a.fs_key) \
            or ([] if a.contra else by_name.get(a.name)) or []
        hit = next((h for h in hits if id(h[2]) not in used), None)
        if hit:
            used.add(id(hit[2]))
            a.prior = hit[2]["prior"]
    for major, sub, r in _bs_sections(bs):
        if id(r) in used or not r["prior"]:
            continue
        m = re.fullmatch(r"(.+)\((.+)\)", r["name"])
        contra = bool(m) and is_contra(m.group(1))
        name, parent = (m.group(1), m.group(2)) if contra else (r["name"], None)
        parent_acc = next((a for a in accounts if a.name == parent and not a.contra), None) if parent else None
        key = parent_acc.key if parent_acc else classify(major, sub, name)
        if key is None:
            continue
        acc = Account(None, name, key, contra, parent, prior=r["prior"])
        # 같은 묶음의 마지막 계정 뒤(차감계정은 원래 계정 바로 뒤)에 넣는다
        anchor = parent_acc or next((a for a in reversed(accounts) if a.key == key), None)
        pos = accounts.index(anchor) + 1 if anchor else len(accounts)
        while anchor and pos < len(accounts) and accounts[pos].contra and accounts[pos].parent == anchor.name:
            pos += 1
        accounts.insert(pos, acc)
        notes.append(f"'{name}'은(는) 시산표에 없고 전기 금액({r['prior']:,.0f})만 있어 당기 0으로 넣었습니다")


# ---------------------------------------------------------------- 채우기

@dataclass
class Result:
    placed: list[Account] = field(default_factory=list)
    unplaced: list[Account] = field(default_factory=list)
    overflow: list[Account] = field(default_factory=list)
    schedules: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    output: Path | None = None

    def lines(self) -> list[str]:
        out = [f"TB 탭에 계정 {len(self.placed)}개를 넣었습니다."]
        for a in self.overflow:
            out.append(f"  칸 부족: [{a.code or '-'}] {a.name} ({TARGETS[a.key][1]} 묶음 칸이 모자람. 조서에 행을 늘린 뒤 다시 해 주세요)")
        for a in self.unplaced:
            out.append(f"  넣을 곳을 못 정함: [{a.code or '-'}] {a.name} 차변잔액 {a.db:,.0f} / 대변잔액 {a.cb:,.0f}")
        out += [f"거래처 명세: {s}" for s in self.schedules]
        out += [f"참고: {n}" for n in self.notes]
        return out


def _signed(a: Account, side: str, value: float) -> float:
    """BS 탭 금액 부호: 자산 묶음은 차변 성격 +, 부채·자본 묶음은 대변 성격 +. 차감계정은 -"""
    natural_credit = side == "liab"
    return value if a.credit == natural_credit else -value


def _num(v: float | None):
    return None if v in (None, 0) else v


def fill_tb(pkg: XlsxPackage, tpl: Template, accounts: list[Account], res: Result) -> None:
    tb, bs = pkg.sheet("TB"), pkg.sheet("BS")
    slots: dict[int, list[int]] = {g.row: list(g.slots) for g in tpl.groups}
    for g in tpl.groups:                          # 양식에 들어 있는 예시 값 지우기
        for r in g.slots:
            for col in "DEFGH":
                tb.set(f"{col}{r}", None)
            if g.side != "is" and r in tpl.bs_row:
                bs.set(f"D{tpl.bs_row[r]}", None)
                bs.set(f"F{tpl.bs_row[r]}", None)
    for a in accounts:
        g = tpl.group(a.key) if a.key in TARGETS else None
        if g is None:
            if a.db or a.cb or a.dt or a.ct:
                res.unplaced.append(a)
            continue
        if not slots[g.row]:
            res.overflow.append(a)
            continue
        r = slots[g.row].pop(0)
        a.tb_row = r
        tb.set(f"F{r}", a.name)
        tb.set(f"D{r}", _num(a.db))
        tb.set(f"E{r}", _num(a.dt))
        tb.set(f"G{r}", _num(a.ct))
        tb.set(f"H{r}", _num(a.cb))
        if g.side == "asset":
            j, k, m = (f"-G{r}-I{r}", f"E{r}", f"-L{r}=H{r}") if a.credit else (f"E{r}-I{r}", f"-G{r}", f"L{r}=D{r}")
        elif g.side == "liab":
            j, k, m = (f"G{r}-I{r}", f"-E{r}", f"L{r}=H{r}") if a.credit else (f"-E{r}-I{r}", f"G{r}", f"-L{r}=D{r}")
        else:
            j = k = m = None
        if j:
            tb.set(f"J{r}", Formula(j))
            tb.set(f"K{r}", Formula(k))
            tb.set(f"M{r}", Formula(m))
            if r in tpl.bs_row:
                cur = (a.db - a.cb) if g.side == "asset" else (a.cb - a.db)
                bs.set(f"D{tpl.bs_row[r]}", _num(cur))
                bs.set(f"F{tpl.bs_row[r]}", _num(_signed(a, g.side, abs(a.prior)) if a.prior else None))
        res.placed.append(a)
    for lead, pairs in tpl.lead_rows.items():      # 리드시트 B열(계정코드)
        by_row = {a.tb_row: a for a in res.placed if a.tb_row}
        ed = pkg.sheet(lead)
        for lr, tr in pairs:
            a = by_row.get(tr)
            ed.set(f"B{lr}", a.code if a and a.code else None)


# IS 탭: 손익계산서(과목별) 또는 시산표에서
_PL_SECTIONS = [("매출원가", None), ("매출총", None), ("매출액", "sales"), ("판매비", "sga"),
                ("영업이익", None), ("영업외수익", "nonop_inc"), ("영업외비용", "nonop_exp"),
                ("법인세차감전", None), ("법인세", "tax"), ("당기순", None)]


def pl_lines(pl: list[dict]) -> list[tuple[str, str, float, float]]:
    """손익계산서(과목별) -> [(IS 묶음, 과목, 당기, 전기)]. 매출원가의 기말재고·타계정대체는 음수로."""
    out, key = [], None
    for r in pl:
        n = section_label(r["name"])
        if not r["is_account"]:
            if n.startswith("매출원가"):
                key = "cogs_other"
            elif key and key.startswith("cogs") and "제품매출원가" in n:
                key = "cogs_product"
            elif key and key.startswith("cogs") and "상품매출원가" in n:
                key = "cogs_goods"
            elif key and key.startswith("cogs") and "매출원가" in n:
                key = "cogs_other"
            else:
                key = next((k for word, k in _PL_SECTIONS if n.startswith(word)), key)
            continue
        if key is None:
            continue
        sign = -1 if key.startswith("cogs") and ("기말" in n or "타계정" in n) else 1
        out.append((key, n, sign * r["current"], sign * r["prior"]))
    return out


def tb_is_lines(accounts: list[Account]) -> list[tuple[str, str, float, float | None]]:
    out = []
    for a in accounts:
        if a.key in ("sales", "nonop_inc"):
            out.append((a.key, a.name, a.cb - a.db, None))
        elif a.key == "cogs":
            k = "cogs_goods" if "상품" in a.name else "cogs_product" if "제품" in a.name else "cogs_other"
            out.append((k, a.name, a.db - a.cb, None))
        elif a.key in ("sga", "nonop_exp", "tax"):
            out.append((a.key, a.name, a.db - a.cb, None))
    return out


def fill_is(pkg: XlsxPackage, tpl: Template, lines, res: Result) -> None:
    ws = pkg.sheet("IS")
    free = {k: list(v) for k, v in tpl.is_groups.items()}
    for rows in tpl.is_groups.values():
        for r in rows:
            for col in "CDF":
                ws.set(f"{col}{r}", None)
    for key, name, cur, pri in lines:
        if not free.get(key):
            res.notes.append(f"IS 탭 칸 부족 또는 묶음 없음: {name}")
            continue
        r = free[key].pop(0)
        ws.set(f"C{r}", name)
        ws.set(f"D{r}", _num(cur))
        ws.set(f"F{r}", _num(pri))


# ---------------------------------------------------------------- 거래처 명세 (○1100)

def schedule_tab(tpl: Template, letter: str) -> str | None:
    for name in tpl.schedule_tabs:
        if letter in name[:-4].split(","):
            return name
    return None


def _free_area(ed) -> tuple[int, int]:
    """'감사절차 및 감사결과' 아래 빈 줄의 (시작, 다음 내용 행)."""
    xml = ed.xml
    rows = [int(r) for r in re.findall(r'<row\b[^>]*?\br="(\d+)"', xml)]
    texts = {}
    for m in re.finditer(r'<c\b[^>]*?\br="([A-Z]+)(\d+)"[^>]*?(?:/>|>(.*?)</c>)', xml, re.S):
        if m.group(3) and re.search(r"<(v|t|f)\b", m.group(3)):
            texts.setdefault(int(m.group(2)), True)
    used = sorted(texts)
    start = next((r for r in used if r >= 9), 9) + 2
    nxt = next((r for r in used if r >= start), max(rows or [start]) + 1)
    return start, nxt


def fill_schedules(pkg: XlsxPackage, tpl: Template, accounts: list[Account], partners: list[dict],
                   res: Result) -> None:
    by_account: dict[tuple, list[dict]] = {}
    for p in partners:
        by_account.setdefault((p.get("account_code") or "", p.get("account_name") or ""), []).append(p)
    for (code, acc_name), rows in by_account.items():
        a = next((x for x in res.placed if code and x.code == code and not x.contra), None) or \
            next((x for x in res.placed if acc_name and x.name == acc_name and not x.contra), None)
        if a is None:
            res.notes.append(f"거래처원장 '{acc_name or code}': 시산표에서 같은 계정을 못 찾아 명세를 쓰지 않았습니다")
            continue
        g = tpl.group(a.key)
        tab = schedule_tab(tpl, g.letter) if g and g.letter else None
        if tab is None:
            res.notes.append(f"거래처원장 '{acc_name}': {g.letter if g else '?'}1100 탭이 없어 명세를 쓰지 않았습니다")
            continue
        rows = [r for r in rows if any(r.get(k) for k in ("opening", "debit", "credit", "balance"))]
        rows.sort(key=lambda r: -abs(r["balance"]))
        ed = pkg.sheet(tab)
        st = dict(tpl.styles)
        head_ref = tpl.table_head.get(tab)
        if head_ref:
            col, row = re.fullmatch(r"([A-Z]+)(\d+)", head_ref).groups()
            st["head"] = ed.style_of(head_ref) or st["head"]
            st["text"] = st["num"] = ed.style_of(f"{col}{int(row) + 1}") or st["num"]
        start, nxt = _free_area(ed)
        need = len(rows) + 6
        if start + need > nxt - 1:
            try:
                pkg.insert_rows(tab, nxt, start + need - nxt + 1)
            except ValueError:
                start = ed.max_row() + 2
        r = start
        ed.set(f"B{r}", f"거래처별 명세: {a.name}" + (f" (계정 {a.code})" if a.code else ""), st.get("title"))
        ed.set(f"I{r}", "출처: 거래처원장", st.get("text"))
        r += 1
        for col, h in zip("BCDEFGHI", ("No.", "거래처", "전기이월", "차변", "대변", "잔액", "구성비", "비고")):
            ed.set(f"{col}{r}", h, st.get("head"))
        first = r + 1
        last = first + len(rows) - 1
        total = last + 1
        for i, p in enumerate(rows, start=1):
            r += 1
            ed.set(f"B{r}", i, st.get("text"))
            ed.set(f"C{r}", p["partner"], st.get("text"))
            for col, k in zip("DEFG", ("opening", "debit", "credit", "balance")):
                ed.set(f"{col}{r}", p[k] or 0, st.get("num"))
            ed.set(f"H{r}", Formula(f"IFERROR(G{r}/G${total},0)"), st.get("pct"))
            ed.set(f"I{r}", None, st.get("text"))
        r = total
        ed.set(f"C{r}", "합계", st.get("head"))
        for col in "DEFG":
            ed.set(f"{col}{r}", Formula(f"SUM({col}{first}:{col}{last})") if rows else 0, st.get("num"))
        ed.set(f"H{r}", Formula(f"SUM(H{first}:H{last})") if rows else None, st.get("pct"))
        ed.set(f"C{r + 1}", "시산표 잔액 (TB)", st.get("text"))
        ed.set(f"G{r + 1}", Formula(f"TB!L{a.tb_row}"), st.get("num"))
        ed.set(f"C{r + 2}", "차이", st.get("text"))
        ed.set(f"G{r + 2}", Formula(f"G{total}-G{r + 1}"), st.get("num"))
        ledger_total = sum(p["balance"] for p in rows)
        tb_bal = (a.db - a.cb) if not a.credit else (a.cb - a.db)
        diff = ledger_total - tb_bal
        res.schedules.append(f"{tab} ← {a.name} 거래처 {len(rows)}곳"
                             + ("" if abs(diff) < 1 else f" (원장 합계와 시산표 차이 {diff:,.0f}원: 기준일 확인 필요)"))


def _styles(pkg: XlsxPackage, tpl: Template) -> dict[str, str | None]:
    tb = pkg.sheet("TB")
    st = {"num": tb.style_of("D12"), "text": tb.style_of("F12"), "head": tb.style_of("D7"), "title": None, "pct": None}
    lead = next(iter(tpl.lead_rows), None)
    if lead:
        r = tpl.lead_rows[lead][0][0]
        st["pct"] = pkg.sheet(lead).style_of(f"L{r}")
    return st


# ---------------------------------------------------------------- 전체

def build(workpaper: str | Path, sources: list[str | Path], output: str | Path,
          ledger_names: dict[str, str] | None = None) -> Result:
    workpaper, output = Path(workpaper), Path(output)
    if workpaper.resolve() == output.resolve():
        raise ValueError("원본 조서에 덮어쓰지 않습니다. 다른 이름으로 저장해 주세요.")
    b = books_mod.load(sources, ledger_names)
    if not b.tb:
        raise ValueError("합계잔액시산표가 없습니다. 더존 '합계잔액시산표' 엑셀을 함께 넣어 주세요.")
    res = Result(notes=list(b.notes))
    tpl = read_template(workpaper)
    accounts = accounts_from_tb(b.tb)
    attach_prior(accounts, b.bs, res.notes)
    pkg = XlsxPackage(workpaper)
    tpl.styles = _styles(pkg, tpl)
    fill_tb(pkg, tpl, accounts, res)
    used = {tpl.group(a.key).row for a in res.placed}
    gaps = [g for g in tpl.missing_from_total if g.row in used]
    if gaps:
        res.notes.append("조서 TB 탭의 합계 행 수식에 " + "·".join(g.label for g in gaps)
                         + " 묶음이 빠져 있어 TB 검증(차변=대변)이 FALSE로 나옵니다. 양식의 합계 수식을 확인해 주세요.")
    if b.pl:
        lines = pl_lines(b.pl)
        tb_sales = sum(c for k, _, c, _ in tb_is_lines(accounts) if k == "sales")
        pl_sales = sum(c for k, _, c, _ in lines if k == "sales")
        if abs(tb_sales - pl_sales) >= 1:
            res.notes.append(f"손익계산서 매출액 {pl_sales:,.0f}원과 시산표 매출 {tb_sales:,.0f}원이 다릅니다. "
                             "두 파일의 기준일이 같은지 확인해 주세요 (IS 탭은 손익계산서 금액으로 넣음).")
    else:
        lines = tb_is_lines(accounts)
        res.notes.append("손익계산서(과목별) 파일이 없어 IS 탭은 시산표 금액으로 넣고 전기는 비워 두었습니다.")
    fill_is(pkg, tpl, lines, res)
    if b.partners:
        fill_schedules(pkg, tpl, accounts, b.partners, res)
    pkg.recalc_on_open()
    pkg.drop_cached_values()
    res.output = pkg.save(output)
    return res
