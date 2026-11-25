"""한공회 채권채무 전자조회(NICE 감사지원 서비스)에서 회신현황 엑셀을 직접 받아 온다.

사이트 동작 (2026-11 확인):
- 로그인: POST /cm/CM0100M003PR.nice  loginid, loginpwd, captchaCd  → XML <LoginResult code="SUCCESS" nextUrl=…>
  여러 번 틀리면 자동입력방지 문자(캡차)가 필요해지는데, 이때는 사람이 브라우저에서 직접 로그인해야 한다.
- 프로젝트 목록: /rp/RP0100M001GE.nice 화면의 프로젝트 목록 (getUpche('프로젝트번호'))
- 엑셀: POST /rp/RPCM00M002PR.nice  formId=ED0011(회신현황) / ED0012(미회신현황) / ED0010(발송내역), adtprjrgstno=프로젝트번호

이용약관 제9조: ID 공유 금지, 감사 목적 외 사용 금지 → 사람마다 자기 ID로 쓴다.
표준 라이브러리(urllib, http.cookiejar)만 쓴다.
"""

from __future__ import annotations

import http.cookiejar
import re
import ssl
import urllib.parse
import urllib.request
from dataclasses import dataclass
from html.parser import HTMLParser
from xml.etree import ElementTree

BASE = "https://kicpa.nicebizline.com"
LOGIN_PAGE = "/cm/CM0200M001GE.nice"
LOGIN = "/cm/CM0100M003PR.nice"
LIST_PAGE = "/rp/RP0100M001GE.nice"
EXCEL = "/rp/RPCM00M002PR.nice"
FORMS = {"reply": "ED0011", "noreply": "ED0012", "sent": "ED0010"}
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/131.0.0.0 Safari/537.36")
CRED_KEY = "nicebiz:"          # 비밀번호 저장 이름: nicebiz:아이디

LOGIN_MESSAGES = {
    "FailNumber": "유효하지 않은 회계사 등록번호입니다 (휴업 회계사는 이용할 수 없음).",
    "EPC700": "비밀번호를 바꿔야 합니다. 사이트에 직접 로그인해서 비밀번호를 바꿔 주세요.",
}


class LoginError(Exception):
    pass


@dataclass
class Project:
    id: str
    year: str
    name: str
    kind: str
    created: str
    in_progress: bool = True

    @property
    def label(self) -> str:
        return f"{self.year} | {self.name} | {self.kind} | {self.created}"


class _ProjectParser(HTMLParser):
    """<div class="pj_list"> 안의 <a href="javascript:getUpche('…')">글자</a> 를 모은다."""

    def __init__(self):
        super().__init__()
        self.depth = 0          # pj_list div 안의 깊이 (0이면 밖)
        self.cur: str | None = None
        self.text: list[str] = []
        self.found: list[tuple[str, str]] = []

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if self.depth:
            if tag == "div":
                self.depth += 1
        elif tag == "div" and "pj_list" in (a.get("class") or "").split():
            self.depth = 1
            return
        if tag == "a":
            m = re.search(r"getUpche\(\s*['\"]?([^'\")]+)", (a.get("href") or "") + (a.get("onclick") or ""))
            if m:
                self.cur, self.text = m.group(1).strip(), []

    def handle_endtag(self, tag):
        if tag == "a" and self.cur:
            self.found.append((self.cur, " ".join("".join(self.text).split())))
            self.cur = None
        elif tag == "div" and self.depth:
            self.depth -= 1

    def handle_data(self, data):
        if self.cur:
            self.text.append(data)


def parse_projects(html: str, in_progress: bool = True) -> list[Project]:
    p = _ProjectParser()
    p.feed(html)
    out, seen = [], set()
    for pid, text in p.found:
        if pid in seen:
            continue
        seen.add(pid)
        parts = [x.strip() for x in re.split(r"\s*[|｜ㅣ│]\s*", text) if x.strip()]
        year = parts[0] if parts and re.fullmatch(r"\d{4}", parts[0]) else ""
        rest = parts[1:] if year else parts
        out.append(Project(pid, year, rest[0] if rest else text, rest[1] if len(rest) > 1 else "",
                           rest[2] if len(rest) > 2 else "", in_progress))
    return out


class Client:
    def __init__(self, base: str = BASE, opener=None, timeout: int = 30):
        self.base = base.rstrip("/")
        self.timeout = timeout
        self.jar = http.cookiejar.CookieJar()
        self.opener = opener or urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(self.jar),
            urllib.request.HTTPSHandler(context=ssl.create_default_context()))
        self.logged_in = False

    def _req(self, path: str, data: dict | None = None, referer: str = LIST_PAGE):
        body = urllib.parse.urlencode(data).encode() if data is not None else None
        req = urllib.request.Request(self.base + path, data=body, headers={
            "User-Agent": UA, "Referer": self.base + referer, "Accept-Language": "ko-KR,ko;q=0.9"})
        if body is not None:
            req.add_header("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8")
        return self.opener.open(req, timeout=self.timeout)

    def _text(self, resp) -> str:
        raw = resp.read()
        m = re.search(r"charset=([\w-]+)", resp.headers.get("Content-Type", ""))
        return raw.decode(m.group(1) if m else "utf-8", errors="replace")

    def login(self, user: str, password: str) -> None:
        self._text(self._req(LOGIN_PAGE, referer=LOGIN_PAGE))      # 세션 쿠키 받기
        xml = self._text(self._req(LOGIN, {"loginid": user, "loginpwd": password, "captchaCd": ""},
                                   referer=LOGIN_PAGE))
        try:
            root = ElementTree.fromstring(xml.strip().encode("utf-8"))
        except ElementTree.ParseError as exc:
            raise LoginError("로그인 응답을 읽지 못했습니다. 사이트 점검 중인지 확인해 주세요.") from exc
        node = root if root.tag == "LoginResult" else root.find(".//LoginResult")
        code = node.get("code") if node is not None else None
        if code != "SUCCESS":
            msg = LOGIN_MESSAGES.get(code or "", "")
            if not msg:
                msg = (f"로그인 실패 (응답 코드 {code or '없음'}). 아이디·비밀번호를 확인해 주세요. "
                       "여러 번 틀려서 자동입력방지 문자가 필요해졌을 수 있으니, 브라우저에서 한 번 직접 로그인한 뒤 다시 해 주세요.")
            raise LoginError(msg)
        nxt = node.get("nextUrl")
        if nxt and nxt.startswith("/"):
            self._text(self._req(nxt, referer=LOGIN_PAGE))
        self.logged_in = True

    def _list_form(self, project_id: str = "", in_progress: bool = True, form_id: str = "") -> dict:
        return {"formId": form_id, "adtprjrgstno": project_id, "sortkey": "", "sortorder": "", "entrnm": "",
                "pgs_yn": "Y" if in_progress else "N", "st_snd_date": "", "end_snd_date": "",
                "st_rpl_date": "", "end_rpl_date": "", "searchSelect": "bfsl_cus_entrnm", "searchValue": ""}

    def projects(self, include_ended: bool = False) -> list[Project]:
        out = parse_projects(self._text(self._req(LIST_PAGE, self._list_form())), True)
        if not out and not self.logged_in:
            raise LoginError("로그인이 필요합니다.")
        if include_ended:
            ids = {p.id for p in out}
            out += [p for p in parse_projects(self._text(self._req(LIST_PAGE, self._list_form(in_progress=False))), False)
                    if p.id not in ids]
        return out

    def excel(self, project_id: str, kind: str = "reply", in_progress: bool = True) -> bytes:
        resp = self._req(EXCEL, self._list_form(project_id, in_progress, FORMS[kind]))
        data = resp.read()
        if not data.startswith(b"PK"):
            raise LoginError("엑셀 대신 다른 화면이 왔습니다. 로그인이 풀렸거나 프로젝트 번호가 맞지 않습니다.")
        return data


def match_project(projects: list[Project], client_name: str, year: int | str | None = None) -> Project | None:
    """회사 이름(괄호·(주) 표기 무시)과 연도로 프로젝트 하나를 고른다. 여럿이면 진행 중·최근 것."""
    from ..confirm import norm_name

    key = norm_name(client_name)
    if not key:
        return None
    hits = [p for p in projects if key in norm_name(p.name) or norm_name(p.name) in key]
    if year:
        same_year = [p for p in hits if p.year == str(year)]
        hits = same_year or hits
    hits.sort(key=lambda p: (p.in_progress, p.year, p.created), reverse=True)
    return hits[0] if hits else None
