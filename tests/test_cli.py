import pandas as pd

from auditautopilot.cli import _months_from_period, main


def test_months_from_period():
    assert _months_from_period("제 14(당)기 2026년1월1일 ~ 2026년9월30일") == 9
    assert _months_from_period("제 13(전)기 2025년1월1일 ~ 2025년12월31일") == 12


def test_cli_end_to_end(semusarang_file, workpaper_file, tmp_path, capsys):
    assert main(["materiality", str(semusarang_file), "--benchmark", "매출액"]) == 0
    out = capsys.readouterr().out
    assert "기준 기간(개월): 9" in out

    an = tmp_path / "an.xlsx"
    assert main(["analytics", str(semusarang_file), "-o", str(an), "--threshold-amount", "100000"]) == 0
    assert "질문서 초안" in pd.ExcelFile(an).sheet_names

    sp = tmp_path / "s.xlsx"
    assert main(["sample", str(semusarang_file), "--account", "108", "--method", "mus",
                 "--n", "2", "--seed", "1", "-o", str(sp)]) == 0
    assert set(pd.ExcelFile(sp).sheet_names) == {"표본", "추출근거", "모집단"}

    imp = tmp_path / "imp.xlsx"
    assert main(["import", "semusarang", str(semusarang_file), "-o", str(imp)]) == 0
    partners = pd.read_excel(imp, sheet_name="거래처별잔액", dtype=str)
    assert "900101-1******" in set(partners["등록번호"])

    wp = tmp_path / "wp.xlsx"
    assert main(["workpaper", "prepare", str(workpaper_file), "-o", str(wp), "--assignee", "홍길동"]) == 0
    assert main(["workpaper", "prepare", str(workpaper_file), "-o", str(wp), "--assignee", "없음"]) == 1
