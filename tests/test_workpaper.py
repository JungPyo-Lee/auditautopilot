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
