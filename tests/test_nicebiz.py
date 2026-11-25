"""전자조회 사이트 가져오기. 실제 사이트 대신 같은 모양으로 답하는 가짜 서버를 띄운다 (가짜 거래처만)."""

import io
import threading
import urllib.parse
from datetime import date
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import openpyxl
import pytest

from auditautopilot.manage import config, confirmations, mailbox, nicebiz
from auditautopilot.manage.db import Store

LIST_HTML = """<html><body><div class="pj_list"><ul>
<li class="fs on"><a href="javascript:getUpche('20261102000001');"><span>2026 | 가나상사 | 감사(외부감사법) | 2026.11.02</span></a></li>
<li class="fs"><a href="javascript:getUpche('20260511000002');"><span>2026 ｜ 다라건설 ｜ 감사 ｜ 2026.05.11</span></a></li>
</ul></div><table><tr><th>기업명</th></tr></table></body></html>"""
ENDED_HTML = """<div class="pj_list"><ul><li><a href="javascript:getUpche('20250301000009')">2025 | 가나상사 | 감사 | 2025.03.01</a></li></ul></div>"""


def _xlsx(rows):
    wb = openpyxl.Workbook()
    for r in rows:
        wb.active.append(r)
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


REPLY = _xlsx([
    ["기업명", "조회차수", "사업자등록번호", "담당자명", "담당자이메일", "일치여부", "통화", "구분", "발송금액", "회신금액", "차이금액", "최초발송일시", "회신일시"],
    ["마바상사", 1, "1112233333", "김", "a@x.test", "Y", "KRW", "외상매출금(커머스)", 1000, 1000, "0.00", "2026-11-03 10:00:00", "2026-11-05 10:00:00"],
    ["사아물산", 1, "2223344444", "이", "b@x.test", "N", "KRW", "외상매입금", 500, 450, "-50.00", "2026-11-03 10:00:00", "2026-11-06 10:00:00"],
])
NOREPLY = _xlsx([
    ["기업명", "조회차수", "사업자등록번호", "담당자명", "담당자이메일", "통화", "구분", "발송금액", "최초발송일시", "최종발송일시", "발송횟수"],
    ["자차상사", 1, "3334455555", "박", "c@x.test", "KRW", "미지급금", 300, "2026-11-03 10:00:00", "2026-11-20 10:00:00", 3],
])


class FakeNice(BaseHTTPRequestHandler):
    users = {"cpa01": "pw!"}
    log: list = []

    def log_message(self, *a):
        pass

    def _send(self, body: bytes, ctype="text/html; charset=UTF-8", cookie=None):
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        if cookie:
            self.send_header("Set-Cookie", cookie)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _authed(self):
        return "SID=ok" in (self.headers.get("Cookie") or "")

    def do_GET(self):
        FakeNice.log.append(("GET", self.path))
        if self.path.startswith(nicebiz.LOGIN_PAGE):
            return self._send(b"<html>login</html>", cookie="JSESSIONID=anon; Path=/")
        if self.path.startswith("/main"):
            return self._send(b"main")
        self.send_error(404)

    def do_POST(self):
        n = int(self.headers.get("Content-Length") or 0)
        f = dict(urllib.parse.parse_qsl(self.rfile.read(n).decode()))
        FakeNice.log.append(("POST", self.path, f))
        if self.path == nicebiz.LOGIN:
            ok = FakeNice.users.get(f.get("loginid")) == f.get("loginpwd") and "JSESSIONID" in (self.headers.get("Cookie") or "")
            code = "SUCCESS" if ok else "EPC100"
            xml = f'<?xml version="1.0" encoding="UTF-8"?><root><CM0100M003PR><LoginResult code="{code}" nextUrl="/main"/></CM0100M003PR></root>'
            return self._send(xml.encode(), "text/xml; charset=UTF-8", cookie="SID=ok; Path=/" if ok else None)
        if not self._authed():
            return self._send("<html>로그인이 필요합니다</html>".encode())
        if self.path == nicebiz.LIST_PAGE:
            return self._send((LIST_HTML if f.get("pgs_yn") == "Y" else ENDED_HTML).encode())
        if self.path == nicebiz.EXCEL:
            data = {"ED0011": REPLY, "ED0012": NOREPLY}[f["formId"]]
            assert f["adtprjrgstno"] == "20261102000001"
            return self._send(data, "binary/octet-stream")
        self.send_error(404)


@pytest.fixture
def site():
    srv = ThreadingHTTPServer(("127.0.0.1", 0), FakeNice)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    FakeNice.log = []
    yield f"http://127.0.0.1:{srv.server_address[1]}"
    srv.shutdown()
    srv.server_close()


@pytest.fixture
def store(tmp_path):
    mailbox._PASSWORD.clear()
    s = Store(tmp_path / "a.db")
    s.set_setting("root", str(tmp_path / "감사"))
    cid = s.add_client("C001", "(주)가나상사", "김담당", "kim@ganada.test")
    s.add_engagement(cid, 2026, "기말감사", "2027-01-11", "2027-01-15")
    return s


def test_parse_projects():
    ps = nicebiz.parse_projects(LIST_HTML)
    assert [(p.id, p.year, p.name, p.created) for p in ps] == [
        ("20261102000001", "2026", "가나상사", "2026.11.02"), ("20260511000002", "2026", "다라건설", "2026.05.11")]
    assert nicebiz.match_project(ps, "주식회사 가나상사", 2026).id == "20261102000001"
    assert nicebiz.match_project(ps, "없는회사", 2026) is None


def test_login_failure_message(site):
    c = nicebiz.Client(base=site)
    with pytest.raises(nicebiz.LoginError, match="자동입력방지"):
        c.login("cpa01", "wrong")


def test_fetch_saves_and_reconciles(site, store):
    store.set_setting("nice_user", "cpa01")
    confirmations.set_password("cpa01", "pw!")
    cfg = config.load(store)
    assert confirmations.ready(cfg)
    e = store.engagements()[0]
    factory = lambda: nicebiz.Client(base=site)  # noqa: E731
    rec = confirmations.fetch(store, e, cfg, date(2026, 11, 25), client_factory=factory)
    assert rec["counts"] == {"일치": 1, "불일치": 1, "미회신": 1}
    assert confirmations.project_of(store, e["id"]) == "20261102000001"     # 회사 이름으로 찾아 연결
    assert len(confirmations.cached_projects(store)) == 3                    # 진행 중 2 + 종료 1
    paths = confirmations.stored_paths(store, e["id"])
    assert len(paths) == 2 and all("채권채무조회" in str(p) for p in paths)
    assert "pw!" not in str([store.get_setting(k) for k in ("nice_user", "nice_projects", f"confirm:{e['id']}")])

    # 같은 내용이면 파일을 다시 저장하지 않는다
    before = len(store.files(e["id"]))
    confirmations.fetch(store, e, cfg, date(2026, 11, 26), client_factory=factory)
    assert len(store.files(e["id"])) == before

    log = confirmations.fetch_all(store, cfg, date(2026, 11, 26), client_factory=factory)
    assert log == ["전자조회 (주)가나상사: 일치 1, 불일치 1, 미회신 1"]


def test_noreply_file_reads_as_unanswered():
    from auditautopilot import confirm
    wb = openpyxl.load_workbook(io.BytesIO(NOREPLY))
    [r] = confirm.read_status_rows([tuple(x) for x in wb.active.iter_rows(values_only=True)])
    assert r["status"] == "미회신" and r["send_count"] == 3 and r["sent"] == 300


def test_web_fetch(site, store, monkeypatch):
    from auditautopilot.manage.web import make_server

    monkeypatch.setattr(nicebiz, "BASE", site)
    monkeypatch.setattr(nicebiz.Client.__init__, "__defaults__", (site, None, 30))
    srv = make_server(store, "127.0.0.1", 0, today=lambda: date(2026, 11, 25))
    port = srv.server_address[1]
    srv.RequestHandlerClass.allowed_hosts = {f"127.0.0.1:{port}"}
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{port}"
    import urllib.request

    class NoRedirect(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, *a, **k):
            return None
    op = urllib.request.build_opener(NoRedirect)

    def post(path, data):
        try:
            r = op.open(base + path, urllib.parse.urlencode(data).encode())
            return r.status, r.headers.get("Location", "")
        except urllib.error.HTTPError as err:
            return err.code, urllib.parse.unquote(err.headers.get("Location", ""))
    try:
        e = store.engagements()[0]
        assert "전자조회 계정" in op.open(base + f"/e/{e['id']}").read().decode()
        assert post("/settings/nice", {"nice_user": "cpa01", "password": "pw!"})[0] == 303
        st, loc = post("/settings/nice-test", {})
        assert st == 303 and "로그인 성공" in urllib.parse.unquote(loc)
        page = op.open(base + f"/e/{e['id']}").read().decode()
        assert "2026 | 다라건설" in page and "전자조회에서 지금 가져오기" in page
        st, loc = post(f"/e/{e['id']}/confirm/fetch", {})
        assert st == 303 and "미회신 1" in urllib.parse.unquote(loc)
        x = op.open(base + f"/e/{e['id']}/confirm.xlsx").read()
        wb = openpyxl.load_workbook(io.BytesIO(x))
        assert [r[0] for r in wb["미회신_대체절차"].iter_rows(values_only=True)][1] == "자차상사"
        mailbox._PASSWORD.clear()
        store.set_setting("nice_user", "cpa01")
        confirmations.set_password("cpa01", "bad")
        st, loc = post(f"/e/{e['id']}/confirm/fetch", {})
        assert st == 303 and "err=" in loc and "자동입력방지" in urllib.parse.unquote(loc)
    finally:
        srv.shutdown()
        srv.server_close()
