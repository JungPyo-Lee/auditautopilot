"""엑셀(.xlsx) 파일을 XML 수준에서 고친다.

openpyxl로 열고 다시 저장하면 조서 안의 도형(화살표 연결선 등)이 사라지므로,
필요한 XML 부분(workbook.xml, 관계 파일, 새 시트)만 바꾸고 나머지는 그대로 복사한다.

표준 라이브러리만 쓴다. xml.etree로 다시 쓰면 엑셀이 요구하는 네임스페이스 접두어
(mc:Ignorable 이 가리키는 x15, xr 등)가 바뀌어 엑셀이 '복구'를 요구하므로,
읽기는 정규식·etree로 하고 쓰기는 원래 문자열을 필요한 곳만 바꾼다.
"""

from __future__ import annotations

import html
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
WORKSHEET_TYPE = NS_REL + "/worksheet"
WORKSHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

_TOKEN_RE = re.compile(r"[A-Za-z_\\\u0080-￿][\w.\\\u0080-￿]*")
_FORMULA_RE = re.compile(rb"<(?:\w+:)?f\b[^>]*>(.*?)</(?:\w+:)?f>", re.S)
_FORMULA1_RE = re.compile(rb"<(?:\w+:)?formula\d?\b[^>]*>(.*?)</", re.S)
_SHEET_TAG_RE = re.compile(r"<sheet\b[^>]*?/>")
_DEFINED_NAME_RE = re.compile(r"<definedName\b([^>]*?)(?:/>|>(.*?)</definedName>)", re.S)
_ATTR_RE = re.compile(r'([\w:]+)="([^"]*)"')


def _attrs(tag_text: str) -> dict[str, str]:
    return {k: html.unescape(v) for k, v in _ATTR_RE.findall(tag_text)}


@dataclass
class SheetInfo:
    name: str
    rel_id: str
    state: str  # visible | hidden | veryHidden
    path: str   # zip 안 경로, 예: xl/worksheets/sheet3.xml


@dataclass
class DefinedName:
    name: str
    value: str
    hidden: bool
    local_sheet_id: int | None


class XlsxPackage:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path) as z:
            self._infos = z.infolist()
            self._data = {i.filename: z.read(i.filename) for i in self._infos}
        self._wb = self._data["xl/workbook.xml"].decode("utf-8")
        self._rels = self._data["xl/_rels/workbook.xml.rels"].decode("utf-8")
        self._ct = self._data["[Content_Types].xml"].decode("utf-8")
        self._new_files: dict[str, bytes] = {}

    # ---- 시트 ----
    def _rel_targets(self) -> dict[str, str]:
        out = {}
        root = ET.fromstring(self._rels)
        for rel in root.findall(f"{{{NS_PKG_REL}}}Relationship"):
            target = rel.get("Target")
            if target.startswith("/"):
                target = target[1:]
            elif not target.startswith("xl/"):
                target = "xl/" + target
            out[rel.get("Id")] = target
        return out

    def _sheet_tags(self) -> list[re.Match]:
        sheets_block = re.search(r"<sheets>(.*?)</sheets>", self._wb, re.S)
        if not sheets_block:
            return []
        return list(_SHEET_TAG_RE.finditer(self._wb, sheets_block.start(1), sheets_block.end(1)))

    def sheets(self) -> list[SheetInfo]:
        targets = self._rel_targets()
        result = []
        for m in self._sheet_tags():
            a = _attrs(m.group(0))
            rid = a.get("r:id") or next((v for k, v in a.items() if k.endswith(":id")), "")
            result.append(SheetInfo(a["name"], rid, a.get("state", "visible"), targets.get(rid, "")))
        return result

    def set_state(self, name: str, state: str) -> None:
        for m in self._sheet_tags():
            tag = m.group(0)
            if _attrs(tag).get("name") != name:
                continue
            new = re.sub(r'\s+state="[^"]*"', "", tag)
            if state != "visible":
                new = new[:-2].rstrip() + f' state="{state}"/>'
            self._wb = self._wb[:m.start()] + new + self._wb[m.end():]
            return
        raise KeyError(name)

    # ---- 이름 정의 ----
    def defined_names(self) -> list[DefinedName]:
        out = []
        for m in _DEFINED_NAME_RE.finditer(self._wb):
            a = _attrs(m.group(1))
            local = a.get("localSheetId")
            out.append(DefinedName(
                name=a.get("name", ""),
                value=html.unescape(m.group(2) or ""),
                hidden=a.get("hidden") in ("1", "true"),
                local_sheet_id=int(local) if local is not None else None,
            ))
        return out

    def _formula_tokens(self) -> set[str]:
        """시트·차트 수식에 쓰인 식별자(소문자). 이름이 실제로 쓰이는지 판단할 때 쓴다."""
        tokens: set[str] = set()
        for name, data in self._data.items():
            if not name.endswith(".xml"):
                continue
            if not (name.startswith("xl/worksheets/") or name.startswith("xl/charts/")):
                continue
            for regex in (_FORMULA_RE, _FORMULA1_RE):
                for m in regex.finditer(data):
                    text = html.unescape(m.group(1).decode("utf-8", "ignore"))
                    tokens.update(t.lower() for t in _TOKEN_RE.findall(text))
        return tokens

    def clean_defined_names(self, aggressive: bool = False) -> dict[str, int]:
        """깨진 이름 정의(#REF!)를 지운다. 시트 수식에서 쓰이는 이름은 남긴다.

        aggressive=True 이면 숨김 처리된 이름 중 어디에서도 쓰이지 않는 이름도 지운다.
        엑셀 기본 이름(_xlnm.Print_Area 등)은 건드리지 않는다.
        """
        names = self.defined_names()
        used = self._formula_tokens()
        for n in names:
            if "#REF!" not in n.value:
                used.update(t.lower() for t in _TOKEN_RE.findall(n.value))
        counts = {"ref": 0, "unused": 0}

        def keep_or_drop(m: re.Match) -> str:
            a = _attrs(m.group(1))
            name = a.get("name", "")
            value = html.unescape(m.group(2) or "")
            if name.startswith("_xlnm.") or name.lower() in used:
                return m.group(0)
            if "#REF!" in value:
                counts["ref"] += 1
                return ""
            if aggressive and a.get("hidden") in ("1", "true"):
                counts["unused"] += 1
                return ""
            return m.group(0)

        self._wb = _DEFINED_NAME_RE.sub(keep_or_drop, self._wb)
        self._wb = re.sub(r"<definedNames>\s*</definedNames>|<definedNames/>", "", self._wb)
        return {"before": len(names), "removed_ref": counts["ref"], "removed_unused": counts["unused"],
                "after": len(names) - counts["ref"] - counts["unused"]}

    # ---- 시트 추가 ----
    def add_sheet_first(self, name: str, sheet_xml: bytes) -> None:
        """맨 앞에 새 시트를 넣는다. 시트 순서가 바뀌므로 localSheetId와 활성 탭도 맞춘다."""
        if name in {s.name for s in self.sheets()}:
            raise ValueError(f"이미 '{name}' 시트가 있습니다.")
        used_paths = set(self._data) | set(self._new_files)
        n = 1
        while f"xl/worksheets/sheet{n}.xml" in used_paths:
            n += 1
        part = f"xl/worksheets/sheet{n}.xml"
        rel_ids = set(re.findall(r'Id="([^"]+)"', self._rels))
        k = 1
        while f"rId{k}" in rel_ids:
            k += 1
        rid = f"rId{k}"
        rel = f'<Relationship Id="{rid}" Type="{WORKSHEET_TYPE}" Target="worksheets/sheet{n}.xml"/>'
        self._rels = self._rels.replace("</Relationships>", rel + "</Relationships>")
        override = f'<Override PartName="/{part}" ContentType="{WORKSHEET_CT}"/>'
        self._ct = self._ct.replace("</Types>", override + "</Types>")

        # r:id 접두어는 루트에 선언된 이름을 쓰고, 루트에 없으면 새 태그에 직접 선언한다
        # (openpyxl이 만든 파일은 sheet 태그마다 따로 선언한다)
        root_tag = re.search(r"<workbook\b[^>]*>", self._wb).group(0)
        prefix = next((p for p, uri in re.findall(r'xmlns:(\w+)="([^"]+)"', root_tag) if uri == NS_REL), None)
        declare = ""
        if prefix is None:
            prefix, declare = "r", f' xmlns:r="{NS_REL}"'
        max_id = max((int(_attrs(m.group(0))["sheetId"]) for m in self._sheet_tags()), default=0)
        tag = f'<sheet{declare} name={quoteattr(name)} sheetId="{max_id + 1}" {prefix}:id="{rid}"/>'
        self._wb = re.sub(r"<sheets>", "<sheets>" + tag.replace("\\", "\\\\"), self._wb, count=1)

        self._wb = re.sub(r'localSheetId="(\d+)"',
                          lambda m: f'localSheetId="{int(m.group(1)) + 1}"', self._wb)
        self._new_files[part] = sheet_xml
        self.set_active_sheet(0)

    def set_active_sheet(self, index: int) -> None:
        def fix_view(m: re.Match) -> str:
            tag = re.sub(r'\s+(activeTab|firstSheet)="\d+"', "", m.group(0))
            end = "/>" if tag.endswith("/>") else ">"
            return tag[: -len(end)].rstrip() + f' firstSheet="0" activeTab="{index}"' + end
        self._wb = re.sub(r"<workbookView\b[^>]*?/?>", fix_view, self._wb)
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
        self._data["xl/workbook.xml"] = self._wb.encode("utf-8")
        self._data["xl/_rels/workbook.xml.rels"] = self._rels.encode("utf-8")
        self._data["[Content_Types].xml"] = self._ct.encode("utf-8")
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
                    widths: list[float]) -> bytes:
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
