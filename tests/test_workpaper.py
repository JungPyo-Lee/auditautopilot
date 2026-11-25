import zipfile

import openpyxl
import pytest

from auditautopilot.workpaper import (
    classify_sheets, make_print_copy, prepare_workpaper, read_assignments,
)


def _states(path):
    wb = openpyxl.load_workbook(path)
    return {ws.title: ws.sheet_state for ws in wb.worksheets}, wb


def test_read_assignments_both_blocks(workpaper_file):
    a = {x.index: x.assignee for x in read_assignments(workpaper_file)}
    assert a == {"A": "홍길동", "AA": "김감사", "C": "김감사", "E100": "홍길동"}


def test_classify_sheets():
    roles = {r.name: (r.kind, r.index) for r in classify_sheets(
        ["총괄", "A", "A1000", "A1100", "AA", "AA1000", "E100", "E110", "BB,DD1000"],
        ["A", "AA", "E", "E100", "BB,DD"])}
    assert roles["총괄"] == ("front", None)
    assert roles["A"] == ("procedure", "A")
    assert roles["AA1000"] == ("lead", "AA")
    assert roles["E100"] == ("detail", "E100")
    assert roles["E110"] == ("detail", "E")
    assert roles["BB,DD1000"] == ("lead", "BB,DD")


def test_prepare_hides_procedures_and_adds_index(workpaper_file, tmp_path):
    out = tmp_path / "work.xlsx"
    report = prepare_workpaper(workpaper_file, out)
    states, wb = _states(out)
    assert wb.sheetnames[0] == "목차"
    assert wb.active.title == "목차"
    assert states["A"] == states["C"] == states["AA"] == "hidden"
    assert states["A1000"] == states["C1100"] == states["TB"] == "visible"
    assert report.index_sheet_added
    idx = wb["목차"]
    links = {c.value: c.hyperlink.location for row in idx.iter_rows() for c in row if c.hyperlink}
    assert links["A1000"] == "'A1000'!A1"
    assert idx["A1"].value == "목차 - 주식회사 테스트"
    # 숨긴 시트가 '선택됨'으로 남지 않는다
    assert not wb["A1000"].sheet_view.tabSelected


def test_prepare_cleans_broken_names_but_keeps_used(workpaper_file, tmp_path):
    out = tmp_path / "work.xlsx"
    report = prepare_workpaper(workpaper_file, out)
    names = set(openpyxl.load_workbook(out).defined_names)
    assert "junk1" not in names
    assert {"usedbad", "rate", "hiddenjunk"} <= names
    assert report.names["removed_ref"] == 1

    out2 = tmp_path / "work2.xlsx"
    prepare_workpaper(workpaper_file, out2, aggressive_names=True)
    assert "hiddenjunk" not in set(openpyxl.load_workbook(out2).defined_names)


def test_prepare_for_assignee(workpaper_file, tmp_path):
    out = tmp_path / "mine.xlsx"
    prepare_workpaper(workpaper_file, out, assignee="김감사")
    states, wb = _states(out)
    assert states["C1000"] == states["AA1000"] == "visible"
    assert states["A1000"] == states["A1100"] == states["E100"] == "hidden"
    assert states["총괄"] == "visible"
    values = [c.value for row in wb["목차"].iter_rows() for c in row if c.value]
    assert "A1000" not in values and "C1000" in values


def test_unknown_assignee_lists_people(workpaper_file, tmp_path):
    with pytest.raises(ValueError, match="김감사"):
        prepare_workpaper(workpaper_file, tmp_path / "x.xlsx", assignee="없는사람")


def test_print_copy_shows_procedures(workpaper_file, tmp_path):
    out = tmp_path / "print.xlsx"
    make_print_copy(workpaper_file, out)
    states, wb = _states(out)
    assert set(states.values()) == {"visible"}
    assert "목차" not in wb.sheetnames


def test_other_parts_are_copied_unchanged(workpaper_file, tmp_path):
    out = tmp_path / "work.xlsx"
    prepare_workpaper(workpaper_file, out)
    with zipfile.ZipFile(workpaper_file) as a, zipfile.ZipFile(out) as b:
        changed = {"xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"}
        for name in a.namelist():
            if name in changed or not name.startswith("xl/") or name.startswith("xl/worksheets/"):
                continue
            assert a.read(name) == b.read(name), name


def test_refuses_to_overwrite_source(workpaper_file):
    with pytest.raises(ValueError):
        prepare_workpaper(workpaper_file, workpaper_file)


def test_classify_front_helpers_and_reports():
    roles = {r.name: (r.kind, r.index) for r in classify_sheets(
        ["일반사항", "PL", "조회서", "정산표", "A", "A1000", "C", "C1000", "금융평잔", "E", "E1000", "E100",
         "보고서 BS", "GG1000"],
        ["A", "C", "E", "E100", "GG"])}
    assert roles["PL"] == roles["조회서"] == ("front", None)          # 첫 인덱스 탭 앞은 공통
    assert roles["금융평잔"] == ("detail", "C")                        # 인덱스 사이 보조 탭은 앞 인덱스
    assert roles["E1000"] == ("lead", "E") and roles["E100"] == ("detail", "E100")
    assert roles["정산표"] == roles["보고서 BS"] == ("front", None)


def test_assign_matrix_and_override(workpaper_file, tmp_path):
    from auditautopilot.workpaper.prepare import Assignment, people, read_assign_matrix
    wb = openpyxl.load_workbook(workpaper_file)
    m = wb.create_sheet("assign", 1)
    m.append(["Job Assignments"])
    m.append(["계정과목", None, "박감사", "최감사"])
    m.append(["매출채권", None, "V", None])
    m.append(["현금및현금성자산", None, None, "V"])
    m.append(["매입채무", None, "V", "V"])
    wb.save(workpaper_file)
    a = {x.index: x.assignee for x in read_assign_matrix(workpaper_file, read_assignments(workpaper_file))}
    assert a == {"A": "최감사", "AA": "박감사, 최감사", "C": "박감사", "E100": "홍길동"}
    items = [Assignment("A", "현금", "박감사"), Assignment("C", "매출채권", "최감사, 박감사")]
    assert people(items) == ["박감사", "최감사"]
    out = tmp_path / "mine.xlsx"
    prepare_workpaper(workpaper_file, out, assignee="최감사", assignments=items)
    states, wb = _states(out)
    assert states["C1000"] == states["C1100"] == "visible"
    assert states["A1000"] == states["AA1000"] == states["E100"] == "hidden"     # AA·E100은 배정 표에 없음
    assert states["assign"] == "visible" and wb.sheetnames.count("목차") == 1
