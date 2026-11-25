"""엑셀(.xlsx) 파일을 XML 수준에서 고친다.

openpyxl로 열고 다시 저장하면 조서 안의 도형(화살표 연결선 등)이 사라지므로,
여기서는 필요한 XML 부분(workbook.xml, 관계 파일, 새 시트)만 바꾸고 나머지는 그대로 복사한다.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr

from lxml import etree

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
WORKSHEET_TYPE = NS_REL + "/worksheet"
WORKSHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

_TOKEN_RE = re.compile(r"[A-Za-z_\\\u0080-￿][\w.\\\u0080-￿]*")
_FORMULA_RE = re.compile(rb"<(?:\w+:)?f\b[^>]*>(.*?)</(?:\w+:)?f>", re.S)


def _q(tag: str, ns: str = NS_MAIN) -> str:
    return f"{{{ns}}}{tag}"


@dataclass
class SheetInfo:
    name: str
    rel_id: str
    state: str  # visible | hidden | veryHidden
    path: str   # zip 안 경로, 예: xl/worksheets/sheet3.xml


class XlsxPackage:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path) as z:
            self._infos = z.infolist()
            self._data = {i.filename: z.read(i.filename) for i in self._infos}
        self._wb = etree.fromstring(self._data["xl/workbook.xml"])
        self._rels = etree.fromstring(self._data["xl/_rels/workbook.xml.rels"])
        self._ct = etree.fromstring(self._data["[Content_Types].xml"])
        self._new_files: dict[str, bytes] = {}

    # ---- 시트 ----
    def _sheet_elements(self):
        return self._wb.find(_q("sheets")).findall(_q("sheet"))

    def _rel_targets(self) -> dict[str, str]:
        out = {}
        for rel in self._rels.findall(_q("Relationship", NS_PKG_REL)):
            target = rel.get("Target")
            if target.startswith("/"):
                target = target[1:]
            elif not target.startswith("xl/"):
                target = "xl/" + target
            out[rel.get("Id")] = target
        return out

    def sheets(self) -> list[SheetInfo]:
        targets = self._rel_targets()
        result = []
        for el in self._sheet_elements():
            rid = el.get(_q("id", NS_REL))
            result.append(SheetInfo(el.get("name"), rid, el.get("state", "visible"), targets.get(rid, "")))
        return result

    def set_state(self, name: str, state: str) -> None:
        for el in self._sheet_elements():
            if el.get("name") == name:
                if state == "visible":
                    el.attrib.pop("state", None)
                else:
                    el.set("state", state)
                return
        raise KeyError(name)

    # ---- 이름 정의 ----
    def defined_names(self) -> list[etree._Element]:
        dn = self._wb.find(_q("definedNames"))
        return [] if dn is None else dn.findall(_q("definedName"))

    def _formula_tokens(self) -> set[str]:
        """시트·차트 수식에 쓰인 식별자(소문자). 이름이 실제로 쓰이는지 판단할 때 쓴다."""
        tokens: set[str] = set()
        for name, data in self._data.items():
            if not (name.startswith("xl/worksheets/") or name.startswith("xl/charts/")):
                continue
            if not name.endswith(".xml"):
                continue
            for m in _FORMULA_RE.finditer(data):
                text = m.group(1).decode("utf-8", "ignore")
                tokens.update(t.lower() for t in _TOKEN_RE.findall(text))
            # 데이터 유효성·조건부 서식 수식
            for m in re.finditer(rb"<(?:\w+:)?formula1?\b[^>]*>(.*?)</", data, re.S):
                text = m.group(1).decode("utf-8", "ignore")
                tokens.update(t.lower() for t in _TOKEN_RE.findall(text))
        return tokens

    def clean_defined_names(self, aggressive: bool = False) -> dict[str, int]:
        """깨진 이름 정의(#REF!)를 지운다. 시트 수식에서 쓰이는 이름은 남긴다.

        aggressive=True 이면 숨김 처리된 이름 중 어디에서도 쓰이지 않는 이름도 지운다.
        엑셀 기본 이름(_xlnm.Print_Area 등)은 건드리지 않는다.
        """
        names = self.defined_names()
        used = self._formula_tokens()
        # 남길 이름의 정의 안에서 참조하는 다른 이름도 쓰이는 것으로 본다
        for el in names:
            if "#REF!" not in (el.text or ""):
                used.update(t.lower() for t in _TOKEN_RE.findall(el.text or ""))
        removed_ref = removed_unused = 0
        for el in names:
            name = el.get("name", "")
            if name.startswith("_xlnm."):
                continue
            if name.lower() in used:
                continue
            value = el.text or ""
            if "#REF!" in value:
                el.getparent().remove(el)
                removed_ref += 1
            elif aggressive and el.get("hidden") in ("1", "true"):
                el.getparent().remove(el)
                removed_unused += 1
        dn = self._wb.find(_q("definedNames"))
        if dn is not None and len(dn) == 0:
            self._wb.remove(dn)
        return {"before": len(names), "removed_ref": removed_ref,
                "removed_unused": removed_unused,
                "after": len(names) - removed_ref - removed_unused}

    # ---- 시트 추가 ----
    def add_sheet_first(self, name: str, sheet_xml: bytes) -> None:
        """맨 앞에 새 시트를 넣는다. 시트 순서가 바뀌므로 localSheetId와 활성 탭도 맞춘다."""
        existing = {s.name for s in self.sheets()}
        if name in existing:
            raise ValueError(f"이미 '{name}' 시트가 있습니다.")
        used_paths = set(self._data) | set(self._new_files)
        n = 1
        while f"xl/worksheets/sheet{n}.xml" in used_paths:
            n += 1
        part = f"xl/worksheets/sheet{n}.xml"
        rel_ids = {r.get("Id") for r in self._rels}
        k = 1
        while f"rId{k}" in rel_ids:
            k += 1
        rid = f"rId{k}"
        rel = etree.SubElement(self._rels, _q("Relationship", NS_PKG_REL))
        rel.set("Id", rid)
        rel.set("Type", WORKSHEET_TYPE)
        rel.set("Target", f"worksheets/sheet{n}.xml")
        ov = etree.SubElement(self._ct, _q("Override", NS_CT))
        ov.set("PartName", "/" + part)
        ov.set("ContentType", WORKSHEET_CT)

        sheets_el = self._wb.find(_q("sheets"))
        max_id = max(int(s.get("sheetId")) for s in sheets_el.findall(_q("sheet")))
        new = etree.Element(_q("sheet"))
        new.set("name", name)
        new.set("sheetId", str(max_id + 1))
        new.set(_q("id", NS_REL), rid)
        sheets_el.insert(0, new)

        for el in self.defined_names():
            if el.get("localSheetId") is not None:
                el.set("localSheetId", str(int(el.get("localSheetId")) + 1))
        self._new_files[part] = sheet_xml
        self.set_active_sheet(0)

    def set_active_sheet(self, index: int) -> None:
        views = self._wb.find(_q("bookViews"))
        if views is not None:
            for v in views.findall(_q("workbookView")):
                v.set("activeTab", str(index))
                v.set("firstSheet", "0")
        # 다른 시트의 '선택됨' 표시를 지운다 (숨긴 시트가 선택된 채로 남으면 엑셀이 그룹으로 묶는다)
        target = self.sheets()[index].path
        for info in self.sheets():
            if info.path == target or info.path not in self._data:
                continue
            data = self._data[info.path]
            if b'tabSelected="1"' in data:
                self._data[info.path] = data.replace(b' tabSelected="1"', b"")

    def first_visible_index(self) -> int:
        for i, s in enumerate(self.sheets()):
            if s.state == "visible":
                return i
        raise ValueError("보이는 시트가 하나도 없습니다.")

    # ---- 저장 ----
    def save(self, path: str | Path) -> Path:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self._data["xl/workbook.xml"] = etree.tostring(self._wb, xml_declaration=True,
                                                       encoding="UTF-8", standalone=True)
        self._data["xl/_rels/workbook.xml.rels"] = etree.tostring(self._rels, xml_declaration=True,
                                                                  encoding="UTF-8", standalone=True)
        self._data["[Content_Types].xml"] = etree.tostring(self._ct, xml_declaration=True,
                                                           encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for info in self._infos:
                z.writestr(info, self._data[info.filename], compress_type=zipfile.ZIP_DEFLATED)
            for name, data in self._new_files.items():
                z.writestr(name, data)
        return path


def column_letter(index: int) -> str:
    letters = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def build_sheet_xml(rows: list[list[str | None]], links: dict[tuple[int, int], str],
                    widths: list[float], bold_rows: set[int] | None = None) -> bytes:
    """문자열만 있는 단순한 시트 XML을 만든다. links: (행, 열) -> 이동할 위치 ('A1000'!A1)."""
    out = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">',
        '<sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>',
        '<sheetFormatPr defaultRowHeight="16.5"/>',
        "<cols>",
    ]
    for i, w in enumerate(widths, start=1):
        out.append(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>')
    out.append("</cols><sheetData>")
    for r, row in enumerate(rows, start=1):
        out.append(f'<row r="{r}">')
        for c, value in enumerate(row):
            if value is None or value == "":
                continue
            ref = f"{column_letter(c)}{r}"
            out.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{escape(str(value))}</t></is></c>')
        out.append("</row>")
    out.append("</sheetData>")
    if links:
        out.append("<hyperlinks>")
        for (r, c), location in sorted(links.items()):
            ref = f"{column_letter(c)}{r + 1}"
            out.append(f"<hyperlink ref={quoteattr(ref)} location={quoteattr(location)}/>")
        out.append("</hyperlinks>")
    out.append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
    out.append("</worksheet>")
    return "".join(out).encode("utf-8")


def sheet_link(name: str, cell: str = "A1") -> str:
    return "'" + name.replace("'", "''") + "'!" + cell
