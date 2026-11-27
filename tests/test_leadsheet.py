"""합계잔액시산표 → 조서 TB·BS·IS 탭, ○1100 거래처 명세. 가짜 회사·숫자만 쓴다."""

import io
import zipfile

import openpyxl
import pytest

from auditautopilot.cli import main
from auditautopilot.importers import books, douzone
from auditautopilot.workpaper import leadsheet
from auditautopilot.workpaper.xlsx_xml import SheetXml, shift_formula

from test_douzone import _save, dz  # noqa: F401  (픽스처 재사용)
from test_manage import _multipart, _open, server, store  # noqa: F401


def _slot_dr(ws, r, bs_row):
    ws[f"C{r}"] = f"=F{r}"
    ws[f"I{r}"] = f"=BS!F{bs_row}"
    ws[f"J{r}"], ws[f"K{r}"] = f"=E{r}-I{r}", f"=-G{r}"
    ws[f"L{r}"], ws[f"M{r}"] = f"=I{r}+J{r}+K{r}", f"=L{r}=D{r}"


def _slot_cr(ws, r, bs_row):
    ws[f"C{r}"] = f"=F{r}"
    ws[f"I{r}"] = f"=BS!F{bs_row}"
    ws[f"J{r}"], ws[f"K{r}"] = f"=G{r}-I{r}", f"=-E{r}"
    ws[f"L{r}"], ws[f"M{r}"] = f"=I{r}+J{r}+K{r}", f"=L{r}=H{r}"


@pytest.fixture
def template(tmp_path):
    """실제 조서 양식과 같은 모양의 작은 양식."""
    wb = openpyxl.Workbook()
    wb.active.title = "일반사항"
    tb = wb.create_sheet("TB")
    tb["B6"], tb["F6"] = "조서", "계  정  과  목"
    tb["D7"], tb["H7"] = "잔     액", "잔      액"
    tb["F9"], tb["D9"] = "[유동자산]", "=D10"
    tb["F10"], tb["D10"] = "<당좌자산>", "=D11+D15"
    tb["B11"], tb["F11"], tb["D11"] = "A", "1. 현금및현금성자산, 장단기금융상품", "=SUM(D12:D14)"
    tb["B15"], tb["F15"], tb["D15"] = "C", "3. 매출채권(유동성)", "=SUM(D16:D19)"
    tb["F20"], tb["D20"] = "[유동부채]", "=D21"
    tb["B21"], tb["F21"], tb["D21"] = "AA", "<매입채무>", "=SUM(D22:D23)"
    tb["F24"], tb["D24"] = "[자본]", "=D25"
    tb["B25"], tb["F25"], tb["D25"] = "GG", "<이익잉여금>", "=SUM(D26:D26)"
    tb["B27"], tb["F27"], tb["D27"] = "P", "[매출]", "=SUM(D28:D29)"
    tb["B30"], tb["F30"], tb["D30"] = "R", "[판매관리비]", "=SUM(D31:D33)"
    tb["F35"], tb["D35"] = "합계", "=D9+D20+D24+D27+D30"
    bs = wb.create_sheet("BS")
    for r in (12, 13, 14, 16, 17, 18, 19):
        _slot_dr(tb, r, r)
        bs[f"C{r}"] = f"=TB!C{r}"
    for r in (22, 23, 26):
        _slot_cr(tb, r, r)
        bs[f"C{r}"] = f"=TB!C{r}"
    for r in (28, 29, 31, 32, 33):
        tb[f"C{r}"] = f"=F{r}"
    tb["F16"], tb["D16"], tb["F17"] = "공사미수금", 999, "대손충당금"   # 양식에 남아 있는 예시 값
    bs["D16"] = 999
    is_ = wb.create_sheet("IS")
    is_["B5"], is_["C5"], is_["E5"] = "P", "Ⅰ.매출액", "=SUM(D6:D7)"
    is_["C6"] = "(1) 제품매출"
    is_["B8"], is_["C8"], is_["E8"] = "R", "Ⅳ.판매비와관리비", "=SUM(D9:D11)"
    lead = wb.create_sheet("C1000")
    for i, r in enumerate(range(16, 20)):
        lead[f"C{12 + i}"], lead[f"E{12 + i}"], lead[f"H{12 + i}"] = f"=TB!C{r}", f"=TB!I{r}", f"=TB!L{r}"
    for name in ("C1100", "AA1100"):
        s = wb.create_sheet(name)
        s["B6"] = f"='{name[:-4]}1000'!B6"
        s["B9"] = "감사절차 및 감사결과"
        s["B20"] = "차이발견사항"
        s["B22"], s["C22"], s["G22"] = "No.", "내용", "차변금액"
        s.merge_cells("C22:D22")
        s["G27"] = "=SUM(G23:G26)"
        s["B30"] = "감사결론"
    p = tmp_path / "조서양식.xlsx"
    wb.save(p)
    return p


@pytest.fixture
def tb_file(tmp_path):
    return _save(tmp_path, "합계잔액시산표20261031.xlsx", [
        ["차 변", None, "계정과목", "대 변", None],
        ["잔액", "합계", None, "합계", "잔액"],
        ["1,500,000", "4,900,000", "◀유    동    자    산▶", "3,450,000", "50,000"],
        ["1,500,000", "4,900,000", "◁당    좌    자    산▷", "3,450,000", "50,000"],
        [500000, 900000, "[103] 보    통    예    금", 400000, ""],
        [1000000, 4000000, "[108] 외  상  매  출  금", 3000000, ""],
        ["", 10000, "[109] 대  손  충  당  금", 60000, 50000],
        ["", 1000000, "◀유    동    부    채▶", 1250000, 250000],
        ["", 1000000, "[251] 외  상  매  입  금", 1250000, 250000],
        ["", "", "◀이  익  잉  여  금▶", 1000000, 1000000],
        ["", "", "[377] 미처분이익잉여금", 1000000, 1000000],
        ["", "", "◀매            출▶", 5000000, 5000000],
        ["", "", "[404] 제    품    매    출", 5000000, 5000000],
        [4800000, 4800000, "◀판  매  관  리  비▶", "", ""],
        [4800000, 4800000, "[811] 복  리  후  생  비", "", ""],
        ["", "", "◀손            익▶", "", ""],
        [10, 10, "[999] 알 수 없는 과목", "", ""],
        ["6,300,000", "10,700,000", "합     계", "10,700,000", "6,300,000"],
    ])


def test_read_tb_both_layouts(tb_file, tmp_path):
    rows = douzone._rows_of(tb_file)
    assert douzone.detect_kind(rows) == "tb"
    tb = douzone.read_tb_rows(rows)
    acc = {r["code"]: r for r in tb if r["level"] == "account"}
    assert acc["108"]["name"] == "외상매출금" and acc["109"]["credit_balance"] == 50000
    assert [r["name"] for r in tb if r["level"] == "group"][:2] == ["유동자산", "당좌자산"]
    other = _save(tmp_path, "tb2.xlsx", [
        ["차    변", None, "계정과목", "대    변", None],
        ["잔액", "합계", "계정과목", "합계", "잔액"],
        [100, 100, "         1.유    동    자    산", 0, 0],
        [100, 100, "         ＜당    좌    자    산＞", 0, 0],
        [100, 100, "  [0103]   보    통    예    금", 0, 0],
        [0, 0, "        13.법 인 ( 소 득 ) 세등", 0, 0],
        [100, 100, "                 합      계", 100, 100],
    ])
    tb2 = douzone.read_tb_rows(douzone._rows_of(other))
    assert [(r["code"], r["name"], r["level"]) for r in tb2] == [
        (None, "유동자산", "group"), (None, "당좌자산", "group"), ("103", "보통예금", "account"),
        (None, "법인(소득)세등", "group"), (None, "합계", "total")]


def test_classify():
    c = leadsheet.classify
    assert c("유동자산", "당좌자산", "정기예금") == "cash"
    assert c("유동자산", "당좌자산", "선급법인세") == "tax_asset"
    assert c("유동자산", "당좌자산", "미수금") == "other_quick"
    assert c("유동자산", "재고자산", "원재료") == "inventory"
    assert c("비유동자산", "투자자산", "장기금융상품") == "lt_cash"
    assert c("비유동자산", "투자자산", "매도가능증권") == "lt_sec"
    assert c("비유동자산", "유형자산", "건물") == "ppe"
    assert c("유동부채", None, "단기차입금") == "borrow"
    assert c("유동부채", None, "미지급금") == "other_cl"
    assert c("비유동부채", None, "퇴직급여충당부채") == "retire"
    assert c("자본조정", None, "자기주식") == "cap_adj"
    assert c("제조원가", "노무비", "임금") == "labor"
    assert c("판매비및일반관리비", None, "복리후생비") == "sga"
    assert c("법인(소득)세등", None, "법인세등") == "tax"
    assert c("손익", None, "손익") is None


def test_ledger_matched_by_tb_balance(tb_file, dz):  # noqa: F811
    b = books.load([tb_file, dz["ar"]])
    assert {p["account_name"] for p in b.partners} == {"외상매출금"}
    assert {p["account_code"] for p in b.partners} == {"108"}


def test_build_fills_tabs(template, tb_file, dz, tmp_path):  # noqa: F811
    out = tmp_path / "결과.xlsx"
    res = leadsheet.build(template, [tb_file, dz["bs"], dz["pl"], dz["ar"], dz["ap"]], out)
    assert [a.name for a in res.unplaced] == ["알수없는과목"]
    wb = openpyxl.load_workbook(out)
    tb, bs, is_ = wb["TB"], wb["BS"], wb["IS"]
    assert (tb["F12"].value, tb["D12"].value) == ("보통예금", 500000)
    assert tb["F13"].value is None
    assert (tb["F16"].value, tb["D16"].value, tb["E16"].value, tb["G16"].value) == ("외상매출금", 1000000, 4000000, 3000000)
    assert tb["F17"].value == "대손충당금" and tb["H17"].value == 50000
    assert tb["J17"].value == "=-G17-I17" and tb["K17"].value == "=E17" and tb["M17"].value == "=-L17=H17"
    assert tb["J16"].value == "=E16-I16"                      # 차변 계정은 양식 수식 그대로
    assert tb["F18"].value is None                            # 양식 예시 값은 지움
    assert (bs["D16"].value, bs["F16"].value) == (1000000, 950000)
    assert (bs["D17"].value, bs["F17"].value) == (-50000, -50000)   # 차감계정은 음수
    assert (tb["F22"].value, tb["H22"].value, bs["F22"].value) == ("외상매입금", 250000, 200000)
    assert tb["F26"].value == "미처분이익잉여금"
    assert (tb["F28"].value, tb["G28"].value) == ("제품매출", 5000000)
    assert bs["D26"].value == 1000000 + 200000          # 마감 전 당기 손익을 미처분이익잉여금에 더함
    assert any("당기 손익 200,000원" in n for n in res.notes)
    assert (is_["C6"].value, is_["D6"].value, is_["F6"].value) == ("제품매출", 5000000, 5500000)
    assert wb["C1000"]["B12"].value == "108"

    c = wb["C1100"]
    assert c["B11"].value.startswith("거래처별 명세: 외상매출금")
    assert [c.cell(12, i).value for i in range(2, 10)] == ["No.", "거래처", "전기이월", "차변", "대변", "잔액", "구성비", "비고"]
    assert [c[f"C{r}"].value for r in (13, 14, 15)] == ["가나상사", "다라상사", "마바상사"]
    assert c["G16"].value == "=SUM(G13:G15)" and c["G17"].value == "=TB!L16" and c["G18"].value == "=G16-G17"
    # 아래 '차이발견사항' 표는 행을 넣어 아래로 밀었다 (병합·수식도 함께)
    assert c["B21"].value == "차이발견사항" and c["B23"].value == "No."
    assert "C23:D23" in {str(m) for m in c.merged_cells.ranges}
    assert c["G28"].value == "=SUM(G24:G27)"
    assert wb["AA1100"]["C13"].value == "사아물산"
    assert any("AA1100" in s for s in res.schedules)

    z = zipfile.ZipFile(out)
    assert 'fullCalcOnLoad="1"' in z.read("xl/workbook.xml").decode()
    assert "xl/calcChain.xml" not in z.namelist()


def test_build_without_fs(template, tb_file, tmp_path):
    res = leadsheet.build(template, [tb_file], tmp_path / "o.xlsx")
    assert any("전기 금액" in n for n in res.notes) and any("손익계산서" in n for n in res.notes)
    wb = openpyxl.load_workbook(tmp_path / "o.xlsx")
    assert wb["BS"]["F16"].value is None
    assert (wb["IS"]["C6"].value, wb["IS"]["D6"].value) == ("제품매출", 5000000)
    assert (wb["IS"]["C9"].value, wb["IS"]["D9"].value) == ("복리후생비", 4800000)


def test_needs_tb_and_refuses_overwrite(template, dz, tmp_path):  # noqa: F811
    with pytest.raises(ValueError, match="합계잔액시산표"):
        leadsheet.build(template, [dz["bs"]], tmp_path / "o.xlsx")
    with pytest.raises(ValueError, match="덮어쓰지"):
        leadsheet.build(template, [dz["bs"]], template)


def test_cli_leadsheet(template, tb_file, dz, tmp_path, capsys):  # noqa: F811
    out = tmp_path / "c.xlsx"
    assert main(["leadsheet", str(template), str(tb_file), str(dz["ap"]), "-o", str(out),
                 "--ledger", f"{dz['ap'].name}=외상매입금"]) == 0
    text = capsys.readouterr().out
    assert "TB 탭에 계정" in text and "AA1100" in text and "넣을 곳을 못 정함" in text


def test_web_leadsheet(server, store, template, tb_file, dz):  # noqa: F811
    base, _ = server
    cid = store.add_client("C009", "가나상사", "김담당", "kim@example.com")
    eid = store.add_engagement(cid, 2026, "기말감사", "2027-01-11", "2027-01-15")
    data, hdr = _multipart([], [("workpaper", "조서양식.xlsx", template.read_bytes()),
                                ("files", tb_file.name, tb_file.read_bytes()),
                                ("files", dz["ar"].name, dz["ar"].read_bytes())])
    st, h, _ = _open(base + f"/e/{eid}/leadsheet", data, hdr)
    assert st == 303 and "err" not in h["Location"]
    rec = store.get_setting(f"leadsheet:{eid}")
    assert rec["path"].endswith("조서양식_리드시트.xlsx")
    st, h, x = _open(base + f"/e/{eid}/leadsheet.xlsx")
    assert st == 200 and x[:2] == b"PK"
    assert openpyxl.load_workbook(io.BytesIO(x))["TB"]["F16"].value == "외상매출금"
    page = _open(base + f"/e/{eid}")[2].decode()
    assert "조서양식_리드시트.xlsx" in page and "C1100" in page
    assert "전체계정 조서 내려받기" in page                      # 담당자 배정과 관계없이 전체 계정 조서
    st, _, full = _open(base + f"/e/{eid}/mine.xlsx")
    assert st == 200 and full[:2] == b"PK"


# ---------------------------------------------------------------- XML 편집 도구

def test_shift_formula():
    assert shift_formula("SUM(G43:G46)+TB!C45+\"A50\"", "C1100", 40, 3, True) == "SUM(G46:G49)+TB!C45+\"A50\""
    assert shift_formula("'C1100'!B41+B41", "C1100", 40, 2, False) == "'C1100'!B43+B41"


def test_unshare_then_set():
    xml = ('<worksheet><sheetData><row r="1"><c r="A1" s="3"><f t="shared" ref="A1:A3" si="0">B1*2</f><v>2</v></c></row>'
           '<row r="2"><c r="A2" s="3"><f t="shared" si="0"/><v>4</v></c></row>'
           '<row r="3"><c r="A3" s="3"><f t="shared" si="0"/><v>6</v></c></row></sheetData></worksheet>')
    ed = SheetXml(xml, "S")
    ed.set("A1", 7)
    assert '<c r="A1" s="3"><v>7</v></c>' in ed.xml
    assert "<f>B2*2</f>" in ed.xml and "<f>B3*2</f>" in ed.xml and 't="shared"' not in ed.xml
    ed.set("C5", "새 글자")
    assert ed.xml.index('r="5"') > ed.xml.index('r="3"') and "새 글자" in ed.xml


def test_fix_template(template, tmp_path, capsys):
    wb = openpyxl.load_workbook(template)
    wb["TB"]["D35"] = "=D9+D20+D24+D27"                 # 판매관리비 묶음이 빠진 합계
    wb["TB"]["E35"] = "=E9+E20+E24+E27"
    wb["TB"]["F14"] = "대손충당금"                       # 차감계정인데 일반 수식
    wb.save(template)
    out = tmp_path / "고친양식.xlsx"
    assert main(["leadsheet", str(template), "--fix-template", "-o", str(out)]) == 0
    text = capsys.readouterr().out
    assert "35행 합계 수식에 판매관리비(30행)" in text and "14, 17행" in text and "예시 금액" in text
    tb = openpyxl.load_workbook(out)["TB"]
    assert tb["D35"].value == "=D9+D20+D24+D27+D30" and tb["E35"].value == "=E9+E20+E24+E27+E30"
    assert tb["J14"].value == "=-G14-I14" and tb["M14"].value == "=-L14=H14"
    assert tb["D16"].value is None and tb["F16"].value == "공사미수금"     # 예시 금액만 지우고 과목은 둠
    assert main(["leadsheet", str(template), "-o", str(tmp_path / "x.xlsx")]) == 1
