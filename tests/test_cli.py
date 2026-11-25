import openpyxl

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
    assert "질문서 초안" in openpyxl.load_workbook(an).sheetnames

    sp = tmp_path / "s.xlsx"
    assert main(["sample", str(semusarang_file), "--account", "108", "--method", "mus",
                 "--n", "2", "--seed", "1", "-o", str(sp)]) == 0
    assert set(openpyxl.load_workbook(sp).sheetnames) == {"표본", "추출근거", "모집단"}

    imp = tmp_path / "imp.xlsx"
    assert main(["import", "semusarang", str(semusarang_file), "-o", str(imp)]) == 0
    ws = openpyxl.load_workbook(imp)["거래처별잔액"]
    header = [c.value for c in ws[1]]
    col = header.index("등록번호")
    assert "900101-1******" in {row[col] for row in ws.iter_rows(min_row=2, values_only=True)}

    wp = tmp_path / "wp.xlsx"
    assert main(["workpaper", "prepare", str(workpaper_file), "-o", str(wp), "--assignee", "홍길동"]) == 0
    assert main(["workpaper", "prepare", str(workpaper_file), "-o", str(wp), "--assignee", "없음"]) == 1


def test_no_compiled_packages_are_imported():
    """윈도우 앱 제어 정책에 막히는 컴파일된 패키지를 불러오지 않는지 확인한다.
    (lxml은 openpyxl이 설치돼 있을 때만 스스로 불러오므로 여기서는 소스에서 직접 쓰지 않는지만 본다.)"""
    from pathlib import Path

    import auditautopilot

    src = "\n".join(p.read_text(encoding="utf-8") for p in Path(auditautopilot.__file__).parent.rglob("*.py"))
    for mod in ("pandas", "numpy", "lxml"):
        assert f"import {mod}" not in src and f"from {mod}" not in src
    import subprocess
    import sys

    # pandas·numpy·lxml 이 없는 PC를 흉내 낸다 (openpyxl은 설치돼 있으면 스스로 불러오므로 막아둔다)
    code = (
        "import sys, importlib.abc\n"
        "class Block(importlib.abc.MetaPathFinder):\n"
        "    def find_spec(self, name, path=None, target=None):\n"
        "        if name.split('.')[0] in ('pandas', 'numpy', 'lxml'):\n"
        "            raise ImportError('blocked ' + name)\n"
        "sys.meta_path.insert(0, Block())\n"
        "from auditautopilot.cli import main\n"
        "rc = main(['materiality', '--benchmark', '매출액', '--amount', '1000000000'])\n"
        "print('RC', rc)\n"
    )
    out = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, check=True)
    assert "RC 0" in out.stdout

