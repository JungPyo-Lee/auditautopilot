"""자료요청(PBC) 목록 템플릿.

기본 템플릿 두 가지(일반, 건설)는 인일회계법인 '감사준비자료' 양식의 항목을 옮겼다.
(건설 양식의 '19,20,21년 세무조정계산서'만 일반 양식처럼 '최근 3년간'으로 바꿨다.)
같은 4열 구조(기본자료 | 중간감사 | 기말감사 | 재고실사/금융자산실사)의 엑셀(.xlsx)을
불러와 새 템플릿으로 쓸 수도 있다.
"""

from __future__ import annotations

from pathlib import Path

SECTIONS = ("기본자료", "중간감사", "기말감사", "실사")
SECTION_PREFIX = {"기본자료": "기본", "중간감사": "중간", "기말감사": "기말", "실사": "실사"}

# 감사구분별로 기본으로 넣는 구분. 화면에서 감사 등록 때 바꿀 수 있다.
DEFAULT_SECTIONS = {"중간감사": ("기본자료", "중간감사"), "기말감사": ("기말감사", "실사")}

_COMMON_MID = [
    "회계데이타", "급여대장", "부가가치세 신고서", "어음할인한도", "당좌차월약정사항",
    "관계회사 및 특수관계자와의 거래내역", "담보제공자산내역", "제공받은 담보내역", "지급보증한 내역",
    "제공받은 지급보증내역", "계류중인 소송사건 내역", "주주총회의사록", "이사회의사록", "내부감사자료",
    "공문발송철", "공문접수철", "인감사용대장", "기타중요사항",
]
_COMMON_YE = [
    "회계데이타", "급여대장", "부가가치세 신고서", "어음할인한도", "당좌차월약정사항", "공시지가확인원",
    "관계회사 및 특수관계자와의 거래내역", "담보제공자산내역", "제공받은 담보내역", "지급보증한 내역",
    "제공받은 지급보증내역", "계류중인 소송사건 내역", "보험가입내역", "종업원에 대한 복지, 사회에대한 기여내역",
    "환경관련투자액,폐기물등의 처리내역", "할인어음내역", "견질어음내역", "배서/양도한 어음내역",
    "전직원퇴직금추계액", "주주총회의사록", "이사회의사록", "인감사용대장", "어음수표수불내역",
    "사용제한된 자산내역(질권,근저당등..)", "파생금융상품거래내역", "외화채권채무환산내역",
    "장기차입금의 상환일정(향후5년)", "주식.채권거래,고정자산취득.처분내역", "관계회사 재무제표",
    "부동산등기부등본,산업재산권등록원부", "기타재무제표에 영향을 미친 근거서류",
]
_COMMON_INSPECT = [
    "현금시재실사", "유가증권실사", "받을어음실사", "어음수표용지실사", "조회서발송", "미회수어음수표조회표",
    "재고자산 실사리스트", "재고실사계획서", "조회서컨트롤시트", "현금출납장", "Cut-off 자료",
]

TEMPLATES: dict[str, dict[str, list[str]]] = {
    "일반": {
        "기본자료": [
            "법인등기부등본", "정관", "사업자등록증", "건물/토지등기부등본", "임대차계약서", "차입약정서",
            "조직도", "업무분장표", "업무흐름도(Flowchart):구매,생산,판매", "최근 3년간 세무조정계산서",
            "관계회사및 특수관계자현황", "주주현황", "회사의연혁", "제품설명서", "인원현황",
            "주요계약서(구매,판매,재무,연구개발)", "유가증권신고서등 등록서류", "주요거래처현황(구매,판매,외주등)",
            "거래금융기관현황", "규정집", "기업의 환경기준과 정책", "사외이사현황", "적용받는 법률,규정,지침 등",
            "고문(기술,법률,회계.세무 등)현황", "산업재산권보유현황", "각종인증서(연구소,벤쳐,ISO 등)", "기타중요사항",
        ],
        "중간감사": _COMMON_MID,
        "기말감사": _COMMON_YE,
        "실사": _COMMON_INSPECT,
    },
    "건설": {
        "기본자료": [
            "법인등기부등본", "정관", "사업자등록증", "건물/토지등기부등본", "임대차계약서", "차입약정서",
            "조직도", "업무분장표", "업무흐름도(Flowchart): 도급,자재,외주(내부회계관리제도 자료로 대체가능)", "최근 3년간 세무조정계산서",
            "관계회사및 특수관계자현황", "주주현황", "회사의연혁", "인원현황", "주요계약서(도급, 자재,하도급)",
            "주요거래처현황(발주처, 자재, 하도급등)", "거래금융기관현황", "규정집", "기업의 환경기준과 정책",
            "사외이사현황", "적용받는 법률,규정,지침 등", "고문(기술,법률,회계.세무 등)현황",
            "건설면허및산업재산권보유현황", "각종인증서(연구소,벤쳐,ISO 등)", "기타중요사항",
        ],
        "중간감사": _COMMON_MID,
        "기말감사": _COMMON_YE,
        "실사": _COMMON_INSPECT,
    },
}


def build_items(template: dict[str, list[str]], sections) -> list[dict]:
    """템플릿에서 요청 항목을 만든다. 요청번호는 '중간-03'처럼 구분 약어와 순번."""
    items = []
    for section in sections:
        for i, title in enumerate(template.get(section, []), start=1):
            items.append({"code": f"{SECTION_PREFIX[section]}-{i:02d}", "section": section, "title": title})
    return items


def _section_of(header: str) -> str | None:
    h = (header or "").replace(" ", "")
    if h.startswith("기본자료"):
        return "기본자료"
    if "중간" in h:
        return "중간감사"
    if "기말" in h:
        return "기말감사"
    if "실사" in h:
        return "실사"
    return None


def load_template(path: str | Path) -> dict[str, list[str]]:
    """감사준비자료와 같은 4열 양식의 엑셀에서 템플릿을 읽는다 (.xlsx만. .xls는 엑셀에서 다른 이름으로 저장)."""
    import openpyxl

    path = Path(path)
    if path.suffix.lower() == ".xls":
        raise ValueError(".xls 파일은 엑셀에서 '다른 이름으로 저장 > Excel 통합 문서(.xlsx)'로 바꿔서 불러와 주세요.")
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    header_idx = None
    columns: dict[int, str] = {}
    for idx, row in enumerate(rows[:10]):
        found = {i: s for i, v in enumerate(row) if (s := _section_of(str(v or "")))}
        if len(found) >= 2:
            header_idx, columns = idx, found
            break
    if header_idx is None:
        raise ValueError("머리글(기본자료 | 중간감사 | 기말감사 | 실사)을 찾지 못했습니다.")
    template: dict[str, list[str]] = {s: [] for s in SECTIONS}
    for row in rows[header_idx + 1:]:
        for i, section in columns.items():
            v = str(row[i]).strip() if i < len(row) and row[i] is not None else ""
            if v and v not in template[section]:
                template[section].append(v)
    # 양식 맨 아래 법인명 한 칸(예: '인일회계법인')은 항목이 아니다
    for section in SECTIONS:
        template[section] = [v for v in template[section] if not v.endswith("회계법인")]
    return template


def load_list(path: str | Path, default_section: str) -> list[tuple[str, str]]:
    """새 요청 리스트 엑셀을 읽는다. 반환: [(구분, 자료명)]

    1) 감사준비자료와 같은 4열 양식(기본자료 | 중간감사 | 기말감사 | 실사)
    2) '자료명'(또는 요청자료·항목) 열이 있는 표. '구분' 열이 있으면 그 구분, 없으면 default_section
    3) 둘 다 아니면 첫 열의 글자를 자료명으로 본다
    """
    from ..workpaper.xlsx_xml import open_workbook

    try:
        tpl = load_template(path)
        return [(sec, t) for sec in SECTIONS for t in tpl.get(sec, [])]
    except ValueError as exc:
        if ".xls" in str(exc):
            raise
    wb = open_workbook(path, read_only=True, data_only=True)
    try:
        rows = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()
    title_col = sec_col = None
    start = 0
    for i, r in enumerate(rows[:10]):
        heads = [str(v or "").replace(" ", "") for v in r]
        t = next((j for j, h in enumerate(heads) if h in ("자료명", "요청자료", "요청자료명", "항목", "자료")), None)
        if t is not None:
            title_col, start = t, i + 1
            sec_col = next((j for j, h in enumerate(heads) if h in ("구분", "단계", "시기")), None)
            break
    out: list[tuple[str, str]] = []
    for r in rows[start:]:
        if title_col is None:
            v = next((x for x in r if isinstance(x, str) and x.strip()), None)
        else:
            v = r[title_col] if title_col < len(r) else None
        title = str(v or "").strip()
        if not title or title.endswith("회계법인"):
            continue
        sec = default_section
        if sec_col is not None and sec_col < len(r):
            sec = _section_of(str(r[sec_col] or "")) or default_section
        if (sec, title) not in out:
            out.append((sec, title))
    if not out:
        raise ValueError("요청 리스트에서 자료명을 찾지 못했습니다. '자료명' 열이 있는 표나 감사준비자료 4열 양식으로 올려 주세요.")
    return out
