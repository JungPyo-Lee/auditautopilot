"""테스트용 가짜 파일을 만든다. 실제 회사 자료는 쓰지 않는다."""

from __future__ import annotations

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName


@pytest.fixture
def semusarang_file(tmp_path):
    wb = openpyxl.Workbook()
    tb = wb.active
    tb.title = "TB"
    tb.append(["차    변", None, "계정과목", "대    변", None])
    tb.append(["잔액", "합계", "계정과목", "합계", "잔액"])
    tb.append([1500000, 9000000, "         1.유    동    자    산", 7500000, 0])
    tb.append([1000000, 5000000, "  [0103]   보    통    예    금", 4000000, 0])
    tb.append([500000, 4000000, "  [0108]   외  상  매  출  금", 3500000, 0])
    tb.append([0, 1000000, "  [0251]   외  상  매  입  금", 1300000, 300000])
    tb.append([0, 0, "  [0404]   제  품  매  출", 8000000, 8000000])
    tb.append([1800000, 18000000, "                 합      계", 18000000, 1800000])

    for title, rows in {
        "BS": [
            ["  Ⅰ.유동자산", 0, 1500000, 0, 1200000],
            ["      [0103] 보통예금", 0, 1000000, 0, 900000],
            ["      [0108] 외상매출금", 0, 500000, 0, 300000],
            ["  자산총계", 0, 1500000, 0, 1200000],
        ],
        "PL": [
            ["Ⅰ.매출액", 0, 8000000, 0, 9000000],
            ["   [0404] 제품매출", 8000000, 0, 9000000, 0],
            ["Ⅷ.법인세차감전이익", 0, 900000, 0, 1000000],
        ],
    }.items():
        ws = wb.create_sheet(title)
        ws.append(["과    목", "제 2(당)기 2026년1월1일 ~ 2026년9월30일", None,
                   "제 1(전)기 2025년1월1일 ~ 2025년12월31일", None])
        ws.append(["과    목", "금액", "당기집계금액", "금액", "전기집계금액"])
        for r in rows:
            ws.append(r)

    j = wb.create_sheet("분개장")
    j.append(["일자", "번호", "구분", "계정과목", "계정과목", "금액", None, "적요", "거래처명", "거래처명", "등록번호"])
    j.append(["일자", "전표번호", "구분", "코드", "계정과목", "차변", "대변", "적요", "코드", "거래처명", "등록번호"])
    j.append(["2026-01-05", "00001", "차변", "0108", "외상매출금", 300000, 0, "1월 매출", "00101", "가나상사", "123-45-67890"])
    j.append(["2026-01-05", "00001", "대변", "0404", "제품매출", 0, 300000, "1월 매출", "00101", "가나상사", "123-45-67890"])
    j.append(["2026-02-10", "00002", "차변", "0103", "보통예금", 200000, 0, "회수", "00101", "가나상사", None])
    j.append(["2026-02-10", "00002", "대변", "0108", "외상매출금", 0, 200000, "회수", "00101", "가나상사", None])

    led = wb.create_sheet("108")
    led.append(["코드", "거 래 처", "등록번호", "대표자명", "입금은행", "예금주", "입금계좌번호",
                "전기이월", "차    변", "대    변", "잔    액"])
    led.append(["00101", "가나상사", "123-45-67890", "김가나", None, None, None, 100000, 300000, 200000, 200000])
    led.append(["00102", "개인고객", "900101-1234567", "홍길동", None, None, None, 0, 300000, 0, 300000])
    led.append(["00103", "다라상사", "222-33-44444", "이다라", None, None, None, 0, 0, 0, 0])
    path = tmp_path / "semusarang.xlsx"
    wb.save(path)
    return path


@pytest.fixture
def douzone_journal_file(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["전표일자", "전표번호", "구분", "Code", "계정과목", "차변", "대변", "거래처", "Code", "적요"])
    ws.append(["2026-03-02", "00001", "차변", "253", "미지급금", 1000, 0, "가나상사", "101", "대금 지급"])
    ws.append(["2026-03-02", "00001", "대변", "103", "보통예금", 0, 1000, "가나상사", "101", "대금 지급"])
    path = tmp_path / "douzone.xlsx"
    wb.save(path)
    return path


@pytest.fixture
def workpaper_file(tmp_path):
    wb = openpyxl.Workbook()
    g = wb.active
    g.title = "일반사항"
    g["C5"], g["D5"] = "회계법인 명", "테스트회계법인"
    g["C6"], g["D6"] = "회사 명", "주식회사 테스트"
    s = wb.create_sheet("총괄")
    s.append([])
    s.append([None, "인덱스", "계정과목", "담당자", "종료일", "인덱스", "계정과목", "담당자", "종료일"])
    s.append([None, "A", "현금및현금성자산", "홍길동", None, "AA", "매입채무", "김감사"])
    s.append([None, "C", "매출채권", "김감사", None])
    s.append([None, "E100", "재고자산실사입회", "홍길동", None])
    for name in ["TB", "A", "A1000", "A1100", "C", "C1000", "C1100", "AA", "AA1000", "E100"]:
        wb.create_sheet(name)
    wb["A1000"]["B2"] = "=usedbad+rate"
    wb["A1000"].sheet_view.tabSelected = True
    wb.defined_names["junk1"] = DefinedName("junk1", attr_text="#REF!$A$1")
    wb.defined_names["usedbad"] = DefinedName("usedbad", attr_text="#REF!$B$2")
    wb.defined_names["rate"] = DefinedName("rate", attr_text="'TB'!$A$1")
    wb.defined_names["hiddenjunk"] = DefinedName("hiddenjunk", attr_text='{#N/A,#N/A,FALSE,"x"}', hidden=True)
    path = tmp_path / "workpaper.xlsx"
    wb.save(path)
    return path
