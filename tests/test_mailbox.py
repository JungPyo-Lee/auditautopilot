"""메일 보내기·받기. 실제 서버 대신 가짜 SMTP·IMAP 객체를 쓴다."""

from datetime import date, datetime
from email.message import EmailMessage
from email.utils import make_msgid

import pytest

from auditautopilot.manage import config, mailbox, mailer, schedule
from auditautopilot.manage.db import Store


class FakeSMTP:
    sent: list = []
    blocked: set = set()          # 이 포트로 오면 연결을 끊는다 (백신·방화벽 흉내)

    def __init__(self, host, port, context=None, timeout=None):
        self.host, self.port, self.tls = host, port, port == 465
        if port in FakeSMTP.blocked:
            raise mailbox.smtplib.SMTPServerDisconnected("Connection unexpectedly closed")

    def ehlo(self):
        pass

    def close(self):
        pass

    def starttls(self, context=None):
        self.tls = True

    def login(self, user, pw):
        assert self.tls, "암호화 없이 로그인하면 안 된다"
        if pw != "app-pw":
            raise mailbox.smtplib.SMTPAuthenticationError(535, b"bad")

    def send_message(self, msg):
        FakeSMTP.sent.append(msg)

    def quit(self):
        pass


class FakeIMAP:
    messages: dict = {}          # 계정마다 다른 받은편지함: {user: {uid: raw}} 또는 모든 계정 공통 {uid: raw}
    fail: set = set()

    def __init__(self, host, port, ssl_context=None):
        self.box = {}

    def login(self, user, pw):
        if user in FakeIMAP.fail:
            raise mailbox.imaplib.IMAP4.error("AUTHENTICATIONFAILED")
        m = FakeIMAP.messages
        self.box = m.get(user, {}) if m and isinstance(next(iter(m)), str) else m

    def select(self, box, readonly=False):
        assert readonly
        return "OK", [b"1"]

    def uid(self, cmd, *args):
        if cmd == "SEARCH":
            assert args[1] == "SINCE" and args[2][3:6].isascii()
            return "OK", [b" ".join(self.box)]
        uid = args[0]
        assert "PEEK" in args[1]  # 읽음 표시를 바꾸지 않는다
        return "OK", [(b"1 (BODY[] {10}", self.box[uid]), b")"]

    def logout(self):
        pass


@pytest.fixture
def env(tmp_path, monkeypatch):
    monkeypatch.setattr(mailbox.smtplib, "SMTP_SSL", FakeSMTP)
    monkeypatch.setattr(mailbox.smtplib, "SMTP", FakeSMTP)
    monkeypatch.setattr(mailbox.imaplib, "IMAP4_SSL", FakeIMAP)
    FakeSMTP.sent = []
    FakeSMTP.blocked = set()
    FakeIMAP.messages = {}
    FakeIMAP.fail = set()
    mailbox._PASSWORD.clear()
    s = Store(tmp_path / "a.db")
    s.set_setting("root", str(tmp_path / "감사"))
    s.set_setting("firm_name", "가나회계법인")
    s.set_setting("mail_accounts", [
        {"user": "me@gmail.com", "provider": "Gmail", **mailbox.PRESETS["Gmail"]},
        {"user": "me@hanmail.net", "provider": "다음(한메일)", **mailbox.PRESETS["다음(한메일)"]}])
    s.set_setting("send_from", "me@gmail.com")
    mailbox.set_password("me@gmail.com", "app-pw")
    mailbox.set_password("me@hanmail.net", "app-pw")
    cid = s.add_client("C001", "한국상사", "김담당", "kim@hankook.co.kr", cc="lee@hankook.co.kr")
    eid = s.add_engagement(cid, 2026, "기말감사", "2027-01-11", "2027-01-15")
    return s, s.engagement(eid)


def _mail(sender, subject, attachments=(), body="자료 보내드립니다."):
    m = EmailMessage()
    m["From"] = sender
    m["To"] = "me@gmail.com"
    m["Subject"] = subject
    m["Message-ID"] = make_msgid()
    m.set_content(body)
    for name, data in attachments:
        m.add_attachment(data, maintype="application", subtype="octet-stream", filename=name)
    return m.as_bytes()


def test_send_uses_tls_and_drops_unsent_header(env):
    s, e = env
    cfg = config.load(s)
    d = schedule.due_drafts(s, e, date(2027, 1, 4), cfg)[0]
    mailbox.send(cfg, mailer.message_for(s, cfg, e, d))
    [msg] = FakeSMTP.sent
    assert msg["X-Unsent"] is None and msg["To"] == "kim@hankook.co.kr" and msg["From"] == "me@gmail.com"
    assert any(p.get_filename() for p in msg.iter_attachments())

    accs = cfg["mail_accounts"]
    accs[1]["smtp_port"] = 587  # 네이버처럼 587이면 STARTTLS
    s.set_setting("mail_accounts", accs)
    mailbox.send(config.load(s), mailer.message_for(s, cfg, e, d), "me@hanmail.net")
    assert len(FakeSMTP.sent) == 2 and FakeSMTP.sent[1]["From"] == "me@hanmail.net"


def test_sender_per_client(env):
    s, e = env
    cfg = config.load(s)
    assert mailbox.sender_for(cfg, e) == "me@gmail.com"
    s.update_client(s.clients()[0]["id"], send_from="me@hanmail.net")
    e = s.engagement(e["id"])
    assert mailbox.sender_for(cfg, e) == "me@hanmail.net"
    d = schedule.due_drafts(s, e, date(2027, 1, 4), cfg)[0]
    assert mailer.message_for(s, cfg, e, d)["From"] == "me@hanmail.net"
    assert mailbox.sender_for(cfg, {"send_from": "deleted@x.com"}) == "me@gmail.com"


def test_old_single_account_settings_are_migrated(tmp_path):
    s = Store(tmp_path / "b.db")
    for k, v in {"mail_user": "old@gmail.com", "mail_provider": "Gmail", "imap_host": "imap.gmail.com",
                 "imap_port": 993, "smtp_host": "smtp.gmail.com", "smtp_port": 465}.items():
        s.set_setting(k, v)
    cfg = config.load(s)
    assert cfg["mail_accounts"][0]["user"] == "old@gmail.com" and cfg["send_from"] == "old@gmail.com"


def test_fetch_checks_every_account_and_skips_duplicates(env):
    s, e = env
    same = _mail("kim@hankook.co.kr", "급여대장 [C001-2026-기말]", [("급여대장.xlsx", b"z")])
    FakeIMAP.messages = {
        "me@gmail.com": {b"1": same},
        "me@hanmail.net": {b"1": same,
                           b"2": _mail("kim@hankook.co.kr", "보험 [C001-2026-기말]", [("보험가입내역.pdf", b"p")])},
    }
    r = mailbox.fetch(s, config.load(s), date(2027, 1, 6))
    assert r.checked == 2 and len(r.saved) == 2 and not r.errors
    FakeIMAP.fail = {"me@gmail.com"}
    FakeIMAP.messages["me@hanmail.net"][b"3"] = _mail("kim@hankook.co.kr", "x [C001-2026-기말]", [("a.pdf", b"a")])
    r = mailbox.fetch(s, config.load(s), date(2027, 1, 6))
    assert len(r.saved) == 1 and r.errors and r.errors[0].startswith("me@gmail.com")


def test_missing_password(env):
    s, _ = env
    mailbox._PASSWORD.clear()
    with pytest.raises(ValueError):
        mailbox.account(config.load(s))


def test_fetch_saves_attachments_by_tag_and_sender(env):
    s, e = env
    cfg = config.load(s)
    FakeIMAP.messages = {
        b"1": _mail("김담당 <kim@hankook.co.kr>", "RE: 자료 요청 [C001-2026-기말]",
                    [("기말-12 퇴직금추계.xlsx", b"x"), ("기타.pdf", b"y")]),
        b"2": _mail("박대리 <park@hankook.co.kr>", "급여대장 보내드립니다", [("급여대장.xlsx", b"z")]),
        b"3": _mail("spam@other.com", "광고", [("a.zip", b"q")]),
        b"4": _mail("kim@hankook.co.kr", "대용량 첨부",
                    body="대용량 파일: https://bigfile.mail.daum.net/bigfile_down?key=abc 기한 30일"),
    }
    r = mailbox.fetch(s, cfg, date(2027, 1, 6))
    assert r.checked == 4 and r.skipped == 1 and len(r.saved) == 3 and r.links == 1
    items = {i["code"]: i for i in s.items(e["id"])}
    assert items["기말-12"]["status"] == "수령"            # 파일명의 요청번호
    assert items["기말-02"]["status"] == "수령"            # 급여대장: 같은 도메인 사람이 보냄
    assert s.links(e["id"])[0]["url"].startswith("https://bigfile.mail.daum.net/")
    # 같은 메일은 다시 처리하지 않는다
    again = mailbox.fetch(s, cfg, date(2027, 1, 6))
    assert again.checked == 0


def test_public_domain_needs_exact_address(env):
    s, e = env
    s.update_client(s.clients()[0]["id"], contact_email="ceo@naver.com", cc="")
    cfg = config.load(s)
    FakeIMAP.messages = {b"1": _mail("other@naver.com", "자료", [("급여대장.xlsx", b"z")]),
                         b"2": _mail("ceo@naver.com", "자료", [("급여대장.xlsx", b"z")])}
    r = mailbox.fetch(s, cfg, date(2027, 1, 6))
    assert r.skipped == 1 and len(r.saved) == 1


def test_run_cycle_auto_send_once_a_day(env):
    s, e = env
    s.set_setting("auto_send", True)
    s.set_setting("send_hour", 9)
    state: dict = {}

    def build(eng, d):
        return mailer.message_for(s, config.load(s), eng, d)

    log = mailbox.run_cycle(s, config.load(s), datetime(2027, 1, 4, 8, 0), build, state)
    assert FakeSMTP.sent == [] and not any("자동 발송" in x for x in log)
    log = mailbox.run_cycle(s, config.load(s), datetime(2027, 1, 4, 9, 5), build, state)
    assert len(FakeSMTP.sent) == 1 and any("자동 발송: 한국상사 최초 요청" in x for x in log)
    mailbox.run_cycle(s, config.load(s), datetime(2027, 1, 4, 15, 0), build, state)
    assert len(FakeSMTP.sent) == 1
    assert all(i["status"] == "요청" for i in s.items(e["id"]))


def test_big_file_links():
    text = "링크 https://drive.google.com/file/d/1/view, 그리고 https://example.com/x 와 https://1drv.ms/u/abc."
    assert mailbox.big_file_links(text) == ["https://drive.google.com/file/d/1/view", "https://1drv.ms/u/abc"]


def test_web_mail_settings_and_send(env, monkeypatch):
    import threading
    import urllib.parse
    import urllib.request

    from auditautopilot.manage.web import make_server

    s, e = env
    mailbox._PASSWORD.clear()
    srv = make_server(s, "127.0.0.1", 0, today=lambda: date(2027, 1, 4))
    port = srv.server_address[1]
    srv.RequestHandlerClass.allowed_hosts = {f"127.0.0.1:{port}"}
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{port}"

    class NoRedirect(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, *a, **k):
            return None

    opener = urllib.request.build_opener(NoRedirect)

    def post(path, data):
        try:
            return opener.open(base + path, urllib.parse.urlencode(data).encode()).status
        except urllib.error.HTTPError as err:
            return err.code

    try:
        assert post("/settings/mail", {"provider": "네이버", "user": "me@naver.com", "password": "app-pw"}) == 303
        assert post("/settings/mail", {"provider": "Gmail", "user": "me@gmail.com", "password": "app-pw"}) == 303
        assert post("/settings/mail", {"provider": "다음(한메일)", "user": "me@hanmail.net", "password": "app-pw"}) == 303
        cfg = config.load(s)
        naver = next(a for a in cfg["mail_accounts"] if a["user"] == "me@naver.com")
        assert naver["imap_host"] == "imap.naver.com" and naver["smtp_port"] == 587
        assert [a["user"] for a in cfg["mail_accounts"]] == ["me@gmail.com", "me@hanmail.net", "me@naver.com"]
        assert "app-pw" not in str([s.get_setting(k) for k in config.DEFAULTS])   # 비밀번호는 DB에 없다
        assert post("/settings/mail-options", {"send_from": "me@naver.com", "check_minutes": "5", "send_hour": "9"}) == 303
        assert config.load(s)["send_from"] == "me@naver.com" and not config.load(s)["auto_send"]
        assert post("/settings/mail-delete", {"user": "me@hanmail.net"}) == 303
        assert len(config.load(s)["mail_accounts"]) == 2
        assert "메일 계정" in opener.open(base + "/settings").read().decode()
        page = opener.open(base + f"/e/{e['id']}/mail/request/2027-01-04").read().decode()
        assert "지금 보내기" in page
        assert post(f"/e/{e['id']}/mail/request/2027-01-04/send", {}) == 303
        assert len(FakeSMTP.sent) == 1 and s.request_sent(e["id"]) and FakeSMTP.sent[0]["From"] == "me@naver.com"
        assert post("/settings/mail-test", {"user": "me@gmail.com"}) == 303
        FakeSMTP.blocked = {465}       # 465가 막혀 있으면 587로 되고, 그 포트를 기억한다
        assert post("/settings/mail-test", {"user": "me@gmail.com"}) == 303
        gm = next(a for a in config.load(s)["mail_accounts"] if a["user"] == "me@gmail.com")
        assert gm["smtp_port"] == 587
        FakeSMTP.blocked = {465, 587}
        assert post("/settings/mail-test", {"user": "me@gmail.com"}) == 303   # 실패는 설명과 함께 되돌아간다
        FakeSMTP.blocked = set()
        assert post("/mail/fetch", {}) == 303
    finally:
        srv.shutdown()
        srv.server_close()


def test_connection_falls_back_to_587_and_explains(env):
    s, e = env
    cfg = config.load(s)
    r = mailbox.test_connection(cfg, "me@gmail.com")
    assert r.ok and r.smtp_port == 465

    FakeSMTP.blocked = {465}
    r = mailbox.test_connection(cfg, "me@gmail.com")
    assert r.ok and r.smtp_port == 587 and "587" in r.smtp_msg
    d = schedule.due_drafts(s, e, date(2027, 1, 4), cfg)[0]
    mailbox.send(cfg, mailer.message_for(s, cfg, e, d))       # 보내기도 587로 넘어간다
    assert len(FakeSMTP.sent) == 1

    FakeSMTP.blocked = {465, 587}
    r = mailbox.test_connection(cfg, "me@gmail.com")
    assert not r.smtp_ok and r.imap_ok                         # 받기는 되고 보내기만 막힘
    assert "백신" in r.smtp_msg and "Connection unexpectedly closed" in r.smtp_msg

    FakeSMTP.blocked = set()
    mailbox.set_password("me@gmail.com", "wrong")
    r = mailbox.test_connection(cfg, "me@gmail.com")
    assert not r.smtp_ok and "앱 비밀번호" in r.smtp_msg and r.smtp_port is None
