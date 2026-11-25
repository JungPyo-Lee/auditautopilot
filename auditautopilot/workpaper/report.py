"""감사보고서 초안 (워드 .docx).

- 본문: 독립된 감사인의 감사보고서. 문구는 감사기준서 700 예시(일반기업회계기준, 적정의견) 꼴이고
  회사명·결산일·회계법인·보고서일은 조서 '일반사항' 탭에서 가져온다. 의견 종류(적정·한정·부적정·의견거절)를 고를 수 있고
  적정이 아니면 근거 단락 자리를 [ ] 로 비워 둔다.
- 첨부: 수정분개까지 반영한 재무상태표·손익계산서 요약 (감사인 정산표 기준, 회사 재무제표와 대사용).

표준 라이브러리로 docx(zip + XML)를 직접 만든다.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl

from .adjust import Figure, figures

OPINIONS = ("적정", "한정", "부적정", "의견거절")


@dataclass
class Info:
    company: str = "[회사명]"
    firm: str = "[회계법인]"
    year_end: date | None = None
    prior_end: date | None = None
    report_date: date | None = None
    partner: str = ""


def _d(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    return v if isinstance(v, date) else None


def read_info(workpaper: str | Path) -> Info:
    info = Info()
    wb = openpyxl.load_workbook(workpaper, read_only=True, data_only=True)
    try:
        if "일반사항" not in wb.sheetnames:
            return info
        ws = wb["일반사항"]
        vals = {}
        for row in ws.iter_rows(min_row=1, max_row=40, max_col=6, values_only=True):
            if len(row) >= 4 and isinstance(row[2], str):
                vals[row[2].replace(" ", "")] = row[3]
        info.company = str(vals.get("회사명") or info.company).strip()
        info.firm = str(vals.get("회계법인명") or info.firm).strip()
        info.year_end = _d(vals.get("당기결산일"))
        info.prior_end = _d(vals.get("전기결산일"))
        info.report_date = _d(vals.get("감사종료일"))
        info.partner = str(vals.get("담당이사") or "").strip()
    finally:
        wb.close()
    return info


def _kdate(d: date | None) -> str:
    return f"{d.year}년 {d.month}월 {d.day}일" if d else "20XX년 XX월 XX일"


# ---------------------------------------------------------------- 본문

def paragraphs(info: Info, opinion: str = "적정") -> list[tuple[str, str]]:
    """[(스타일, 글)] 스타일: title | to | h | p | bullet | sign | note"""
    if opinion not in OPINIONS:
        raise ValueError(f"감사의견은 {', '.join(OPINIONS)} 중 하나로 주세요.")
    co, ye = info.company, _kdate(info.year_end)
    fs = (f"해당 재무제표는 {ye} 현재의 재무상태표, 동일로 종료되는 보고기간의 손익계산서, 자본변동표, 현금흐름표 "
          "그리고 유의적인 회계정책의 요약을 포함한 재무제표의 주석으로 구성되어 있습니다.")
    fair = (f"회사의 {ye} 현재의 재무상태와 동일로 종료되는 보고기간의 재무성과 및 현금흐름을 "
            "일반기업회계기준에 따라, 중요성의 관점에서 공정하게 표시하고 있습니다")
    out = [("title", "독립된 감사인의 감사보고서"), ("to", f"{co}"), ("to", "주주 및 이사회 귀중"), ("p", "")]
    head = {"적정": "감사의견", "한정": "한정의견", "부적정": "부적정의견", "의견거절": "의견거절"}[opinion]
    out.append(("h", head))
    if opinion == "의견거절":
        out.append(("p", f"우리는 {co}(이하 \"회사\")의 재무제표에 대한 감사계약을 체결하였습니다. {fs}"))
        out.append(("p", "우리는 별첨된 회사의 재무제표에 대하여 의견을 표명하지 않습니다. 우리는 이 감사보고서의 의견거절근거 "
                         "단락에서 기술된 사항의 유의성 때문에 재무제표에 대한 감사의견의 근거를 제공하는 충분하고 적합한 "
                         "감사증거를 입수할 수 없었습니다."))
    else:
        out.append(("p", f"우리는 {co}(이하 \"회사\")의 재무제표를 감사하였습니다. {fs}"))
        if opinion == "적정":
            out.append(("p", f"우리의 의견으로는 별첨된 회사의 재무제표는 {fair}."))
        elif opinion == "한정":
            out.append(("p", "우리의 의견으로는 별첨된 회사의 재무제표는 이 감사보고서의 한정의견근거 단락에서 기술된 사항이 "
                             f"미칠 수 있는 영향을 제외하고는, {fair}."))
        else:
            out.append(("p", "우리의 의견으로는 별첨된 회사의 재무제표는 이 감사보고서의 부적정의견근거 단락에서 기술된 사항의 "
                             f"유의성으로 인하여 {fair.replace('표시하고 있습니다', '표시하고 있지 않습니다')}."))
    out.append(("h", {"적정": "감사의견근거", "한정": "한정의견근거", "부적정": "부적정의견근거", "의견거절": "의견거절근거"}[opinion]))
    if opinion != "적정":
        out.append(("note", f"[{head}의 근거가 되는 사항과 재무제표에 미치는 영향을 적어 주세요.]"))
    if opinion != "의견거절":
        basis_end = {"적정": "감사의견", "한정": "한정의견", "부적정": "부적정의견"}[opinion]
        out.append(("p", "우리는 대한민국의 회계감사기준에 따라 감사를 수행하였습니다. 이 기준에 따른 우리의 책임은 이 감사보고서의 "
                         "재무제표감사에 대한 감사인의 책임 단락에 기술되어 있습니다. 우리는 재무제표감사와 관련된 대한민국의 "
                         "윤리적 요구사항에 따라 회사로부터 독립적이며, 그러한 요구사항에 따른 기타의 윤리적 책임들을 이행하였습니다. "
                         f"우리가 입수한 감사증거가 {basis_end}을 위한 근거로서 충분하고 적합하다고 우리는 믿습니다."))
    out.append(("h", "기타사항"))
    pe = _kdate(info.prior_end)
    out.append(("note", f"[해당 시 기재. 예: 비교표시된 {pe}로 종료되는 보고기간의 재무제표는 ○○회계법인이 감사하였으며, "
                        "이 감사인의 감사보고서에는 적정의견이 표명되었습니다. 해당 없으면 이 단락을 지워 주세요.]"))
    out.append(("h", "재무제표에 대한 경영진과 지배기구의 책임"))
    out.append(("p", "경영진은 일반기업회계기준에 따라 이 재무제표를 작성하고 공정하게 표시할 책임이 있으며, 부정이나 오류로 인한 "
                     "중요한 왜곡표시가 없는 재무제표를 작성하는데 필요하다고 결정한 내부통제에 대해서도 책임이 있습니다."))
    out.append(("p", "경영진은 재무제표를 작성할 때, 회사의 계속기업으로서의 존속능력을 평가하고 해당되는 경우, 계속기업 관련 사항을 "
                     "공시할 책임이 있습니다. 그리고 경영진이 기업을 청산하거나 영업을 중단할 의도가 없는 한, 회계의 계속기업전제의 "
                     "사용에 대해서도 책임이 있습니다."))
    out.append(("p", "지배기구는 회사의 재무보고절차의 감시에 대한 책임이 있습니다."))
    out.append(("h", "재무제표감사에 대한 감사인의 책임"))
    if opinion == "의견거절":
        out.append(("p", "우리의 책임은 대한민국의 회계감사기준에 따라 회사의 재무제표에 대한 감사를 수행하고 감사보고서를 발행하는 데 "
                         "있습니다. 그러나 우리는 이 감사보고서의 의견거절근거 단락에서 기술된 사항 때문에 재무제표에 대한 감사의견의 "
                         "근거를 제공하는 충분하고 적합한 감사증거를 입수할 수 없었습니다. 우리는 재무제표감사와 관련된 대한민국의 "
                         "윤리적 요구사항에 따라 회사로부터 독립적이며, 그러한 요구사항에 따른 기타의 윤리적 책임들을 이행하였습니다."))
    else:
        out.append(("p", "우리의 목적은 회사의 재무제표에 전체적으로 부정이나 오류로 인한 중요한 왜곡표시가 없는지에 대하여 합리적인 "
                         "확신을 얻어 우리의 의견이 포함된 감사보고서를 발행하는데 있습니다. 합리적인 확신은 높은 수준의 확신을 "
                         "의미하나, 감사기준에 따라 수행된 감사가 항상 중요한 왜곡표시를 발견한다는 것을 보장하지는 않습니다. "
                         "왜곡표시는 부정이나 오류로부터 발생할 수 있으며, 왜곡표시가 재무제표를 근거로 하는 이용자의 경제적 "
                         "의사결정에 개별적으로 또는 집합적으로 영향을 미칠 것이 합리적으로 예상되면, 그 왜곡표시는 중요하다고 "
                         "간주됩니다."))
        out.append(("p", "감사기준에 따른 감사의 일부로서 우리는 감사의 전 과정에 걸쳐 전문가적 판단을 수행하고 전문가적 의구심을 "
                         "유지하고 있습니다. 또한, 우리는:"))
        for b in (
            "부정이나 오류로 인한 재무제표의 중요왜곡표시위험을 식별하고 평가하며 그러한 위험에 대응하는 감사절차를 설계하고 "
            "수행합니다. 그리고 감사의견의 근거로서 충분하고 적합한 감사증거를 입수합니다. 부정은 공모, 위조, 의도적인 누락, "
            "허위진술 또는 내부통제 무력화가 개입될 수 있기 때문에 부정으로 인한 중요한 왜곡표시를 발견하지 못할 위험은 "
            "오류로 인한 위험보다 큽니다.",
            "상황에 적합한 감사절차를 설계하기 위하여 감사와 관련된 내부통제를 이해합니다. 그러나 이는 내부통제의 효과성에 "
            "대한 의견을 표명하기 위한 것이 아닙니다.",
            "재무제표를 작성하기 위하여 경영진이 적용한 회계정책의 적합성과 경영진이 도출한 회계추정치와 관련 공시의 합리성에 "
            "대하여 평가합니다.",
            "경영진이 사용한 회계의 계속기업전제의 적절성과, 입수한 감사증거를 근거로 계속기업으로서의 존속능력에 대하여 "
            "유의적 의문을 초래할 수 있는 사건이나 상황과 관련된 중요한 불확실성이 존재하는지 여부에 대하여 결론을 내립니다. "
            "중요한 불확실성이 존재한다고 결론을 내리는 경우, 우리는 재무제표의 관련 공시에 대하여 감사보고서에 주의를 "
            "환기시키고, 이들 공시가 부적절한 경우 의견을 변형할 것을 요구받고 있습니다. 우리의 결론은 감사보고서일까지 "
            "입수된 감사증거에 기초하나, 미래의 사건이나 상황이 회사의 계속기업으로서 존속을 중단하게 할 수 있습니다.",
            "공시를 포함한 재무제표의 전반적인 표시와 구조 및 내용을 평가하고, 재무제표의 기초가 되는 거래와 사건을 재무제표가 "
            "공정한 방식으로 표시하고 있는지 여부를 평가합니다.",
        ):
            out.append(("bullet", b))
        out.append(("p", "우리는 여러 가지 사항들 중에서 계획된 감사범위와 감사시기 그리고 감사 중 식별된 유의적 미비점 등 유의적인 "
                         "감사의 발견사항에 대하여 지배기구와 커뮤니케이션합니다."))
    rd = _kdate(info.report_date)
    out += [("p", ""), ("sign", f"{info.firm}"), ("sign", "대 표 이 사   [대표이사 성명]"), ("sign", "[주소]"), ("sign", rd),
            ("p", ""),
            ("note", f"이 감사보고서는 감사보고서일({rd}) 현재로 유효한 것입니다. 따라서 감사보고서일 후 이 보고서를 열람하는 "
                     "시점 사이에 첨부된 회사의 재무제표에 중대한 영향을 미칠 수 있는 사건이나 상황이 발생할 수도 있으며 이로 "
                     "인하여 이 감사보고서가 수정될 수도 있습니다.")]
    return out


# ---------------------------------------------------------------- 첨부 요약

def fs_summary(figs: list[Figure], ni: float) -> tuple[list[tuple[str, float, float | None]], list[tuple[str, float, float | None]]]:
    def tot(pred, use_prior=False):
        vals = [(f.prior if use_prior else f.current) for f in figs if pred(f)]
        if use_prior and any(v is None for v in vals) and all(v is None for v in vals):
            return None
        return sum(v or 0 for v in vals)

    def side(s, majors):
        return lambda f: f.row.side == s and f.row.major in majors
    bs = []
    for label, s, majors in (("유동자산", "asset", ("유동자산",)), ("비유동자산", "asset", ("비유동자산",)),
                             ("자산총계", "asset", ("유동자산", "비유동자산")),
                             ("유동부채", "liab", ("유동부채",)), ("비유동부채", "liab", ("비유동부채",)),
                             ("부채총계", "liab", ("유동부채", "비유동부채")), ("자본총계", "liab", ("자본",))):
        cur, pri = tot(side(s, majors)), tot(side(s, majors), True)
        if label == "자본총계":
            cur += ni
        bs.append((label, cur, pri))
    liab = next(v for v in bs if v[0] == "부채총계")
    eq = next(v for v in bs if v[0] == "자본총계")
    bs.append(("부채와자본총계", liab[1] + eq[1], (liab[2] or 0) + (eq[2] or 0) if liab[2] is not None else None))

    def is_tot(keys, use_prior=False):
        return tot(lambda f: f.row.side == "is" and f.row.key in keys, use_prior)
    sales, cogs = is_tot({"sales"}), is_tot({"cogs", "material", "labor", "overhead"})
    sga, ninc, nexp, tax = is_tot({"sga"}), is_tot({"nonop_inc"}), is_tot({"nonop_exp"}), is_tot({"tax"})
    p = {k: is_tot(v, True) for k, v in (("sales", {"sales"}), ("cogs", {"cogs", "material", "labor", "overhead"}),
                                          ("sga", {"sga"}), ("ninc", {"nonop_inc"}), ("nexp", {"nonop_exp"}),
                                          ("tax", {"tax"}))}
    pv = {k: (v or 0) for k, v in p.items()}

    def pr(x):
        return x if p["sales"] is not None else None
    is_ = [("매출액", sales, p["sales"]), ("매출원가", cogs, p["cogs"]),
           ("매출총이익", sales - cogs, pr(pv["sales"] - pv["cogs"])), ("판매비와관리비", sga, p["sga"]),
           ("영업이익(손실)", sales - cogs - sga, pr(pv["sales"] - pv["cogs"] - pv["sga"])),
           ("영업외수익", ninc, p["ninc"]), ("영업외비용", nexp, p["nexp"]),
           ("법인세비용차감전순이익(손실)", sales - cogs - sga + ninc - nexp,
            pr(pv["sales"] - pv["cogs"] - pv["sga"] + pv["ninc"] - pv["nexp"])),
           ("법인세비용", tax, p["tax"]), ("당기순이익(손실)", ni,
                                       pr(pv["sales"] - pv["cogs"] - pv["sga"] + pv["ninc"] - pv["nexp"] - pv["tax"]))]
    return bs, is_


# ---------------------------------------------------------------- docx

def _run(text: str, bold=False, size=None) -> str:
    props = ""
    if bold or size:
        props = "<w:rPr>" + ("<w:b/>" if bold else "") + (f'<w:sz w:val="{size}"/>' if size else "") + "</w:rPr>"
    return f'<w:r>{props}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def _para(text: str, style: str) -> str:
    ppr = {"title": '<w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr>',
           "h": '<w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>',
           "sign": '<w:pPr><w:jc w:val="center"/></w:pPr>',
           "bullet": '<w:pPr><w:ind w:left="400" w:hanging="200"/><w:jc w:val="both"/></w:pPr>',
           "p": '<w:pPr><w:jc w:val="both"/><w:spacing w:after="120"/></w:pPr>',
           "note": '<w:pPr><w:jc w:val="both"/><w:spacing w:after="120"/></w:pPr>',
           "to": ""}.get(style, "")
    if style == "title":
        body = _run(text, True, 32)
    elif style == "h":
        body = _run(text, True, 22)
    elif style == "bullet":
        body = _run("• " + text)
    elif style == "note":
        body = _run(text, size=18)
    else:
        body = _run(text)
    return f"<w:p>{ppr}{body}</w:p>"


def _table(head: list[str], rows: list[list[str]]) -> str:
    def cell(t, bold=False, right=False):
        jc = '<w:jc w:val="right"/>' if right else ""
        return (f'<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:pPr>{jc}</w:pPr>'
                f"{_run(t, bold)}</w:p></w:tc>")
    borders = "".join(f'<w:{b} w:val="single" w:sz="4" w:color="999999"/>'
                      for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
    out = [f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr>']
    out.append("<w:tr>" + "".join(cell(h, True) for h in head) + "</w:tr>")
    for r in rows:
        out.append("<w:tr>" + "".join(cell(v, False, i > 0) for i, v in enumerate(r)) + "</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


def _amt(v) -> str:
    return "" if v is None else f"{v:,.0f}"


def write_docx(path: Path, body_xml: str) -> Path:
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" '
          'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" '
          'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
            'officeDocument" Target="word/document.xml"/></Relationships>')
    doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
                'styles" Target="styles.xml"/></Relationships>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="맑은 고딕" w:eastAsia="맑은 고딕" w:hAnsi="맑은 고딕"/>'
              '<w:sz w:val="20"/><w:lang w:eastAsia="ko-KR"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '</w:styles>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body>{body_xml}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1300" w:bottom="1440" w:left="1300" w:header="851" w:footer="992" w:gutter="0"/>'
           '</w:sectPr></w:body></w:document>')
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc)
    return path


PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def draft(workpaper: str | Path, output: str | Path, opinion: str = "적정",
          aje_path: str | Path | None = None) -> tuple[Path, list[str]]:
    info = read_info(workpaper)
    figs, ni, notes = figures(workpaper, aje_path)
    body = "".join(_para(t, s) for s, t in paragraphs(info, opinion))
    bs, is_ = fs_summary(figs, ni)
    ye, pe = _kdate(info.year_end), _kdate(info.prior_end)
    body += PAGE_BREAK + _para("(첨부) 재무제표 요약 - 감사인 정산표 기준 (회사 재무제표와 대사용, 감사보고서에 넣지 않음)", "h")
    body += _para(f"재무상태표  당기 {ye} 현재 · 전기 {pe} 현재 (단위: 원)", "p")
    body += _table(["과목", "당기", "전기"], [[a, _amt(b), _amt(c)] for a, b, c in bs])
    body += _para("", "p") + _para("손익계산서 (단위: 원)", "p")
    body += _table(["과목", "당기", "전기"], [[a, _amt(b), _amt(c)] for a, b, c in is_])
    asset = next(v for v in bs if v[0] == "자산총계")[1]
    le = next(v for v in bs if v[0] == "부채와자본총계")[1]
    if abs(asset - le) >= 1:
        notes.append(f"재무상태표 대차가 {asset - le:,.0f}원 맞지 않습니다 (정산표 검증 칸을 확인해 주세요)")
    if info.company.startswith("["):
        notes.append("조서 '일반사항' 탭에서 회사명을 못 찾아 [회사명]으로 두었습니다")
    out = write_docx(Path(output), body)
    return out, notes
