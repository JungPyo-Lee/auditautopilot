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
        self._sheet_edits: dict[str, SheetXml] = {}

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
        self.add_sheet(name, sheet_xml, first=True)

    def add_sheet(self, name: str, sheet_xml: bytes, first: bool = False) -> None:
        """새 시트를 맨 뒤(first=True면 맨 앞)에 넣는다."""
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
        self._new_files[part] = sheet_xml
        if not first:
            i = self._wb.index("</sheets>")
            self._wb = self._wb[:i] + tag + self._wb[i:]
            return
        self._wb = re.sub(r"<sheets>", "<sheets>" + tag.replace("\\", "\\\\"), self._wb, count=1)

        self._wb = re.sub(r'localSheetId="(\d+)"',
                          lambda m: f'localSheetId="{int(m.group(1)) + 1}"', self._wb)
        self.set_active_sheet(0)

    def remove_sheet(self, name: str) -> None:
        """이 도구가 전에 넣은 시트를 다시 만들 때 지운다 (다른 시트의 수식이 가리키지 않는 시트만)."""
        info = next((x for x in self.sheets() if x.name == name), None)
        if info is None:
            return
        self._wb = _SHEET_TAG_RE.sub(lambda m: "" if _attrs(m.group(0)).get("name") == name else m.group(0),
                                     self._wb)
        self._rels = re.sub(rf'<Relationship\b[^>]*Id="{re.escape(info.rel_id)}"[^>]*/>', "", self._rels)
        part = info.path
        self._ct = re.sub(rf'<Override\b[^>]*PartName="/{re.escape(part)}"[^>]*/>', "", self._ct)
        self._data.pop(part, None)
        self._new_files.pop(part, None)
        self._infos = [i for i in self._infos if i.filename != part]
        self._sheet_edits.pop(name, None)
        rel_part = part.replace("worksheets/", "worksheets/_rels/") + ".rels"
        self._data.pop(rel_part, None)
        self._infos = [i for i in self._infos if i.filename != rel_part]

    # ---- 서식 ----
    def add_basic_styles(self) -> dict[str, str]:
        """새 시트에 쓸 서식(숫자 #,##0, 굵게, 제목, 머리글 바탕색)을 styles.xml 끝에 더한다. 반환: 이름 -> s 번호"""
        xml = self._data["xl/styles.xml"].decode("utf-8")

        def append(tag: str, items: list[str]) -> int:
            nonlocal xml
            m = re.search(rf"<{tag}\b([^>]*)>(.*?)</{tag}>", xml, re.S)
            child = {"fonts": "font", "fills": "fill", "cellXfs": "xf"}[tag]
            first = len(re.findall(rf"<{child}\b", m.group(2)))
            body = m.group(2) + "".join(items)
            head = re.sub(r'\s+count="\d+"', "", m.group(1)) + f' count="{first + len(items)}"'
            xml = xml[:m.start()] + f"<{tag}{head}>{body}</{tag}>" + xml[m.end():]
            return first

        f0 = append("fonts", ['<font><b/><sz val="10"/><name val="맑은 고딕"/></font>',
                              '<font><b/><sz val="13"/><name val="맑은 고딕"/></font>',
                              '<font><sz val="10"/><name val="맑은 고딕"/></font>',
                              '<font><b/><sz val="10"/><color rgb="FFC00000"/><name val="맑은 고딕"/></font>'])
        fill = append("fills", ['<fill><patternFill patternType="solid"><fgColor rgb="FFDDE6F4"/>'
                                '<bgColor indexed="64"/></patternFill></fill>',
                                '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF8DC"/>'
                                '<bgColor indexed="64"/></patternFill></fill>'])
        bold, title, plain, red = f0, f0 + 1, f0 + 2, f0 + 3
        specs = {
            "text": f'<xf numFmtId="0" fontId="{plain}" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
            "num": f'<xf numFmtId="3" fontId="{plain}" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>',
            "bold": f'<xf numFmtId="0" fontId="{bold}" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
            "boldnum": f'<xf numFmtId="3" fontId="{bold}" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>',
            "title": f'<xf numFmtId="0" fontId="{title}" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
            "head": (f'<xf numFmtId="0" fontId="{bold}" fillId="{fill}" borderId="0" xfId="0" applyFont="1" '
                     f'applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'),
            "input": f'<xf numFmtId="3" fontId="{plain}" fillId="{fill + 1}" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>',
            "check": f'<xf numFmtId="3" fontId="{red}" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>',
            "pct": f'<xf numFmtId="10" fontId="{plain}" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>',
        }
        x0 = append("cellXfs", list(specs.values()))
        self._data["xl/styles.xml"] = xml.encode("utf-8")
        return {k: str(x0 + i) for i, k in enumerate(specs)}

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

    # ---- 셀 쓰기 ----
    def sheet(self, name: str) -> "SheetXml":
        """시트 내용을 고칠 수 있는 편집기. 저장할 때 반영된다."""
        if name not in self._sheet_edits:
            info = next((s for s in self.sheets() if s.name == name), None)
            if info is None:
                raise KeyError(name)
            self._sheet_edits[name] = SheetXml(self._data[info.path].decode("utf-8"), name)
        return self._sheet_edits[name]

    def has_sheet(self, name: str) -> bool:
        return any(s.name == name for s in self.sheets())

    def insert_rows(self, name: str, before: int, count: int) -> None:
        """name 시트의 before 행 앞에 빈 행 count개를 넣는다. 다른 시트·이름 정의의 참조도 맞춘다."""
        if count <= 0:
            return
        self.sheet(name).insert_rows(before, count)
        for info in self.sheets():
            if info.name == name or info.path not in self._data:
                continue
            ed = self._sheet_edits.get(info.name)
            if ed is not None:
                ed.shift_foreign_refs(name, before, count)
                continue
            raw = self._data[info.path]
            if _quoted_sheet(name).encode() in raw or (name + "!").encode() in raw:
                ed = self.sheet(info.name)
                ed.shift_foreign_refs(name, before, count)

        def fix(m: re.Match) -> str:
            return m.group(1) + shift_formula(m.group(2), name, before, count, own=False) + m.group(3)
        self._wb = re.sub(r"(<definedName\b[^>]*>)(.*?)(</definedName>)", fix, self._wb, flags=re.S)

    def recalc_on_open(self) -> None:
        """수식 값을 바꿨으므로 엑셀이 열 때 다시 계산하게 한다. calcChain은 지운다(엑셀이 다시 만든다)."""
        if "xl/calcChain.xml" in self._data:
            del self._data["xl/calcChain.xml"]
            self._infos = [i for i in self._infos if i.filename != "xl/calcChain.xml"]
            self._rels = re.sub(r'<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*/>', "", self._rels)
            self._ct = re.sub(r'<Override\b[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', "", self._ct)
        m = re.search(r"<calcPr\b[^>]*?/?>", self._wb)
        if m:
            tag = re.sub(r'\s+fullCalcOnLoad="[^"]*"', "", m.group(0))
            end = "/>" if tag.endswith("/>") else ">"
            tag = tag[: -len(end)].rstrip() + ' fullCalcOnLoad="1"' + end
            self._wb = self._wb[:m.start()] + tag + self._wb[m.end():]
        else:
            self._wb = self._wb.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')

    def drop_cached_values(self) -> int:
        """모든 시트 수식 셀의 저장된 결과(<v>)를 지운다. 미리보기 프로그램이 양식의 옛 예시 값을
        보여 주지 않게 한다. 엑셀은 열 때 다시 계산한다."""
        n = 0
        for info in self.sheets():
            if info.path not in self._data:
                continue
            ed = self.sheet(info.name)
            before = ed.xml
            count = [0]

            def strip(m: re.Match) -> str:
                c = m.group(0)
                if "<f" not in c or "<v>" not in c:
                    return c
                count[0] += 1
                head, rest = c.split(">", 1)
                head = re.sub(r'\s+t="(?:str|b|e|n)"', "", head)
                return head + ">" + re.sub(r"<v>[^<]*</v>", "", rest)
            ed.xml = _CELL_RE.sub(strip, ed.xml)
            n += count[0]
            if ed.xml == before:
                del self._sheet_edits[info.name]
        return n

    # ---- 저장 ----
    def save(self, path: str | Path) -> Path:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        targets = {s.name: s.path for s in self.sheets()}
        for name, ed in self._sheet_edits.items():
            self._data[targets[name]] = ed.xml.encode("utf-8")
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


def grid_xml(cells: dict[str, tuple], widths: list[float], freeze: str | None = None,
             merges: list[str] | None = None) -> bytes:
    """셀 하나하나를 정해 시트 XML을 만든다. cells: 'B3' -> (값, s번호). 값은 글자·숫자·Formula."""
    rows: dict[int, list[tuple[int, str]]] = {}
    for ref, (value, style) in cells.items():
        col, row = split_ref(ref)
        attrs = f' r="{ref}"' + (f' s="{style}"' if style else "")
        if value is None or value == "":
            x = f"<c{attrs}/>"
        elif isinstance(value, Formula):
            x = f"<c{attrs}><f>{escape(value[1:] if value.startswith('=') else value)}</f></c>"
        elif isinstance(value, (int, float)):
            x = f"<c{attrs}><v>{_num(value)}</v></c>"
        else:
            x = f'<c{attrs} t="inlineStr"><is><t xml:space="preserve">{escape(str(value))}</t></is></c>'
        rows.setdefault(row, []).append((column_index(col), x))
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">', '<sheetViews><sheetView workbookViewId="0">']
    if freeze:
        col, row = split_ref(freeze)
        xs, ys = column_index(col) - 1, row - 1
        attrs = (f' xSplit="{xs}"' if xs else "") + (f' ySplit="{ys}"' if ys else "")
        pane = "bottomRight" if xs and ys else ("bottomLeft" if ys else "topRight")
        out.append(f'<pane{attrs} topLeftCell="{freeze}" activePane="{pane}" state="frozen"/>')
    out.append('</sheetView></sheetViews><sheetFormatPr defaultRowHeight="16.5"/><cols>')
    for i, w in enumerate(widths, start=1):
        out.append(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>')
    out.append("</cols><sheetData>")
    for r in sorted(rows):
        out.append(f'<row r="{r}">' + "".join(x for _, x in sorted(rows[r])) + "</row>")
    out.append("</sheetData>")
    if merges:
        out.append(f'<mergeCells count="{len(merges)}">' + "".join(f'<mergeCell ref="{m}"/>' for m in merges)
                   + "</mergeCells>")
    out.append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
    out.append("</worksheet>")
    return "".join(out).encode("utf-8")


def sheet_link(name: str, cell: str = "A1") -> str:
    return "'" + name.replace("'", "''") + "'!" + cell


# ---------------------------------------------------------------- 셀 편집

def column_index(letters: str) -> int:
    """'A' -> 0, 'AA' -> 26"""
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def split_ref(ref: str) -> tuple[str, int]:
    m = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)", ref)
    if not m:
        raise ValueError(ref)
    return m.group(1), int(m.group(2))


class Formula(str):
    """셀에 수식으로 넣을 글자 ('=' 없이 또는 있게)."""


_ROW_RE = re.compile(r'<row\b[^>]*?\br="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.S)
_CELL_RE = re.compile(r'<c\b[^>]*?\br="([A-Z]{1,3})(\d+)"[^>]*?(?:/>|>.*?</c>)', re.S)
_REF_IN_FORMULA = re.compile(
    r"(?<![A-Za-z0-9_.!$'\u0080-￿])"
    r"((?:'(?:[^']|'')+'|[A-Za-z0-9_.\u0080-￿]+)!)?"
    r"(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])")


def _quoted_sheet(name: str) -> str:
    return "'" + name.replace("'", "''") + "'!"


def _sheet_of_prefix(prefix: str) -> str:
    p = prefix[:-1]
    if p.startswith("'"):
        p = p[1:-1].replace("''", "'")
    return p


def shift_formula(text: str, sheet: str, before: int, count: int, own: bool) -> str:
    """수식 안의 참조 중 sheet 시트의 before 행 이상을 count만큼 내린다.
    own=True 면 시트 이름이 없는 참조도 그 시트로 본다. 따옴표 안 글자는 건드리지 않는다."""
    def repl(m: re.Match) -> str:
        prefix = m.group(1)
        target = _sheet_of_prefix(prefix) if prefix else (sheet if own else None)
        row = int(m.group(5))
        if target != sheet or row < before:
            return m.group(0)
        return f"{prefix or ''}{m.group(2)}{m.group(3)}{m.group(4)}{row + count}"
    parts = re.split(r'("(?:[^"]|"")*")', text)
    return "".join(p if i % 2 else _REF_IN_FORMULA.sub(repl, p) for i, p in enumerate(parts))


def _shift_sqref(text: str, before: int, count: int) -> str:
    def repl(m: re.Match) -> str:
        row = int(m.group(4))
        return m.group(0) if row < before else f"{m.group(1)}{m.group(2)}{m.group(3)}{row + count}"
    return re.sub(r"(\$?)([A-Z]{1,3})(\$?)(\d+)", repl, text)


def _num(v: float) -> str:
    if isinstance(v, bool):
        return "1" if v else "0"
    if float(v).is_integer():
        return str(int(v))
    return repr(float(v))


class SheetXml:
    """시트 XML 글자를 필요한 곳만 바꾼다 (나머지 서식·도형 연결은 그대로)."""

    def __init__(self, xml: str, name: str = ""):
        self.xml = xml
        self.name = name

    # ---- 읽기 ----
    def _rows(self) -> list[re.Match]:
        m = re.search(r"<sheetData>(.*?)</sheetData>", self.xml, re.S)
        if not m:
            return []
        return list(_ROW_RE.finditer(self.xml, m.start(1), m.end(1)))

    def max_row(self) -> int:
        return max((int(m.group(1)) for m in self._rows()), default=0)

    def cell_xml(self, ref: str) -> str | None:
        col, row = split_ref(ref)
        for m in self._rows():
            if int(m.group(1)) == row:
                for c in _CELL_RE.finditer(m.group(0)):
                    if c.group(1) == col:
                        return c.group(0)
        return None

    def style_of(self, ref: str) -> str | None:
        x = self.cell_xml(ref)
        m = re.search(r'\bs="(\d+)"', x or "")
        return m.group(1) if m else None

    def has_drawing(self) -> bool:
        return bool(re.search(r"<(drawing|legacyDrawing|tableParts)\b", self.xml))

    # ---- 쓰기 ----
    def _ensure_row(self, row: int) -> tuple[int, int]:
        rows = self._rows()
        for m in rows:
            if int(m.group(1)) == row:
                return m.start(), m.end()
        new = f'<row r="{row}"></row>'
        after = next((m for m in rows if int(m.group(1)) > row), None)
        if after is not None:
            pos = after.start()
        else:
            if "<sheetData/>" in self.xml:
                self.xml = self.xml.replace("<sheetData/>", "<sheetData></sheetData>", 1)
            pos = self.xml.index("</sheetData>")
        self.xml = self.xml[:pos] + new + self.xml[pos:]
        return pos, pos + len(new)

    def set(self, ref: str, value, style: str | None = None) -> None:
        """value: None(비움), 숫자, 글자, Formula. style을 안 주면 원래 셀의 서식을 그대로 둔다."""
        col, row = split_ref(ref)
        start, end = self._ensure_row(row)
        row_xml = self.xml[start:end]
        old = next((c for c in _CELL_RE.finditer(row_xml) if c.group(1) == col), None)
        if old is not None and re.search(r'<f\b[^>]*\bt="shared"[^>]*\bref="', old.group(0)):
            self.unshare(re.search(r'\bsi="(\d+)"', old.group(0)).group(1))
            start, end = self._ensure_row(row)
            row_xml = self.xml[start:end]
            old = next((c for c in _CELL_RE.finditer(row_xml) if c.group(1) == col), None)
        if old is not None:
            if style is None:
                m = re.search(r'\bs="(\d+)"', old.group(0).split(">", 1)[0])
                style = m.group(1) if m else None
        attrs = f' r="{col}{row}"' + (f' s="{style}"' if style else "")
        if value is None or value == "":
            new = f"<c{attrs}/>"
        elif isinstance(value, Formula):
            f = value[1:] if value.startswith("=") else value
            new = f"<c{attrs}><f>{escape(f)}</f></c>"
        elif isinstance(value, (int, float)):
            new = f"<c{attrs}><v>{_num(value)}</v></c>"
        else:
            new = f'<c{attrs} t="inlineStr"><is><t xml:space="preserve">{escape(str(value))}</t></is></c>'
        if old is not None:
            row_xml = row_xml[:old.start()] + new + row_xml[old.end():]
        else:
            if row_xml.endswith("/>"):
                row_xml = row_xml[:-2] + "></row>"
            idx = column_index(col)
            after = next((c for c in _CELL_RE.finditer(row_xml) if column_index(c.group(1)) > idx), None)
            pos = after.start() if after is not None else row_xml.rindex("</row>")
            row_xml = row_xml[:pos] + new + row_xml[pos:]
        row_xml = re.sub(r'\s+spans="[^"]*"', "", row_xml, count=1)
        self.xml = self.xml[:start] + row_xml + self.xml[end:]

    def unshare(self, si: str) -> None:
        """공유 수식 묶음(si)을 셀마다 따로 쓴 수식으로 푼다. 원본 셀을 바꾸기 전에 쓴다."""
        master = re.search(r'<c\b[^>]*?\br="([A-Z]{1,3})(\d+)"[^>]*>(?:(?!</c>).)*?'
                           r'<f\b([^>]*\bt="shared"[^>]*\bsi="%s"[^>]*)>(.*?)</f>' % si, self.xml, re.S)
        if not master or "ref=" not in master.group(3):
            return
        mc, mr, text = column_index(master.group(1)), int(master.group(2)), html.unescape(master.group(4))

        def moved(dc: int, dr: int) -> str:
            def repl(m: re.Match) -> str:
                col = m.group(3) if m.group(2) else column_letter(column_index(m.group(3)) + dc)
                row = m.group(5) if m.group(4) else str(int(m.group(5)) + dr)
                return f"{m.group(1) or ''}{m.group(2)}{col}{m.group(4)}{row}"
            parts = re.split(r'("(?:[^"]|"")*")', text)
            return "".join(p if i % 2 else _REF_IN_FORMULA.sub(repl, p) for i, p in enumerate(parts))

        def fix(m: re.Match) -> str:
            c = m.group(0)
            if not re.search(r'<f\b[^>]*\bt="shared"[^>]*\bsi="%s"' % si, c):
                return c
            col, row = column_index(m.group(1)), int(m.group(2))
            c = re.sub(r'<f\b[^>]*?(?:/>|>.*?</f>)', lambda _: f"<f>{escape(moved(col - mc, row - mr))}</f>", c,
                       count=1, flags=re.S)
            return c
        self.xml = _CELL_RE.sub(fix, self.xml)

    # ---- 행 넣기 ----
    def insert_rows(self, before: int, count: int) -> None:
        if self.has_drawing():
            raise ValueError(f"{self.name}: 도형·표가 있는 시트에는 행을 넣지 않습니다")

        def fix_cell(m: re.Match) -> str:
            c = m.group(0)
            row = int(m.group(2))
            if row >= before:
                c = c.replace(f'r="{m.group(1)}{row}"', f'r="{m.group(1)}{row + count}"', 1)
            return self._shift_formulas_in(c, before, count)

        def fix_row(m: re.Match) -> str:
            r = m.group(0)
            row = int(m.group(1))
            if row >= before:
                r = re.sub(r'(<row\b[^>]*?\br=")\d+"', lambda x: f'{x.group(1)}{row + count}"', r, count=1)
            return _CELL_RE.sub(fix_cell, r)

        data = re.search(r"<sheetData>(.*?)</sheetData>", self.xml, re.S)
        if data:
            body = _ROW_RE.sub(fix_row, data.group(1))
            self.xml = self.xml[:data.start(1)] + body + self.xml[data.end(1):]
        for tag, attr in (("mergeCell", "ref"), ("dimension", "ref"), ("conditionalFormatting", "sqref"),
                          ("dataValidation", "sqref"), ("hyperlink", "ref"), ("selection", "sqref"),
                          ("selection", "activeCell"), ("pane", "topLeftCell")):
            self.xml = re.sub(rf'(<{tag}\b[^>]*?\b{attr}=")([^"]*)(")',
                              lambda m: m.group(1) + _shift_sqref(m.group(2), before, count) + m.group(3), self.xml)
        self.xml = re.sub(r'(<conditionalFormatting\b.*?</conditionalFormatting>)',
                          lambda m: re.sub(r"(<formula>)(.*?)(</formula>)",
                                           lambda f: f.group(1) + shift_formula(f.group(2), self.name, before, count, True)
                                           + f.group(3), m.group(1), flags=re.S),
                          self.xml, flags=re.S)
        self.xml = re.sub(r'(<brk\b[^>]*?\bid=")(\d+)(")',
                          lambda m: m.group(1) + str(int(m.group(2)) + (count if int(m.group(2)) >= before else 0))
                          + m.group(3), self.xml)

    def _shift_formulas_in(self, cell_xml: str, before: int, count: int, sheet: str | None = None,
                           own: bool = True) -> str:
        sheet = sheet or self.name

        def fix_f(m: re.Match) -> str:
            open_tag = m.group(1)
            if own:
                open_tag = re.sub(r'(\bref=")([^"]*)(")',
                                  lambda x: x.group(1) + _shift_sqref(x.group(2), before, count) + x.group(3), open_tag)
            body = html.unescape(m.group(2))
            new = shift_formula(body, sheet, before, count, own)
            return open_tag + (escape(new) if new != body else m.group(2)) + m.group(3)
        return re.sub(r"(<f\b[^>]*>)(.*?)(</f>)", fix_f, cell_xml, flags=re.S)

    def shift_foreign_refs(self, sheet: str, before: int, count: int) -> None:
        """다른 시트(sheet)의 행이 밀렸을 때 이 시트 수식의 그 시트 참조를 맞춘다."""
        data = re.search(r"<sheetData>(.*?)</sheetData>", self.xml, re.S)
        if not data:
            return
        body = re.sub(r"<f\b[^>]*>.*?</f>",
                      lambda m: self._shift_formulas_in(m.group(0), before, count, sheet, own=False),
                      data.group(1), flags=re.S)
        self.xml = self.xml[:data.start(1)] + body + self.xml[data.end(1):]
