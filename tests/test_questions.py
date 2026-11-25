"""질문서: 작성 → 승인 → 발송 → 답변. 가짜 회사·거래처만 쓴다."""

import io
import threading
import urllib.parse
import urllib.request
from datetime import date

import openpyxl

from auditautopilot.manage import config, confirmations, mailbox, nicebiz, questions
from test_mailbox import FakeSMTP
from test_nicebiz import site, store  # noqa: F401  (픽스처 재사용)
from test_manage import server  # noqa: F401

TODAY = date(2027, 1, 12)


def test_flow_add_approve_send_answer(store):  # noqa: F811
    e = store.engagements()[0]
    a = questions.add(store, e["id"], "매출 컷오프 확인을 위해 12월 마지막 주 출고 내역을 주세요.", "매출", None, today=TODAY)
    b = questions.add(store, e["id"], "가지급금 증가 사유", "가지급금", 1_000_000, today=TODAY)
    assert [q["code"] for q in questions.all_(store, e["id"])] == ["Q-001", "Q-002"]

    cfg = config.load(store)
    try:
        questions.draft(store, e, cfg, TODAY)
        raise AssertionError("승인 전에는 보낼 수 없어야 한다")
    except ValueError as exc:
        assert "승인" in str(exc)

    assert questions.act(store, [a], "submit") == (1, [])
    assert questions.act(store, [a, b], "approve", "JP", TODAY) == (2, [])
    q = questions.get(store, a)
    assert q["status"] == "승인" and q["approved_by"] == "JP" and q["approved_on"] == "2027-01-12"
    try:
        questions.edit(store, a, question="바꾼 질문")
        raise AssertionError("승인한 질문은 고칠 수 없어야 한다")
    except ValueError:
        pass
    assert questions.act(store, [b], "delete") == (0, ["Q-002"])     # 승인한 질문은 지울 수 없다

    d = questions.draft(store, e, cfg, TODAY)
    assert d.to == "kim@ganada.test" and "질문 드립니다 (2건)" in d.subject and "[C001-2026-기말]" in d.subject
    assert "Q-002. [가지급금] 가지급금 증가 사유 (1,000,000원)" in d.body and "1월 17일" in d.body
    msg = questions.message(store, e, cfg, d, "me@gmail.com")
    [att] = list(msg.iter_attachments())
    wb = openpyxl.load_workbook(io.BytesIO(att.get_content()))
    rows = [r for r in wb.active.iter_rows(values_only=True)]
    assert rows[3][:5] == ("번호", "계정", "질문", "금액", "답변") and rows[5][0] == "Q-002"

    questions.mark_sent(store, d.ids, TODAY)
    assert {q["status"] for q in questions.all_(store, e["id"])} == {"발송"}

    # 회사가 답변 칸을 채워 보낸 엑셀
    ws = wb.active
    ws.cell(row=5, column=5, value="첨부한 출고 내역 참고 바랍니다.")
    ws.append(["Q-099", "", "없는 질문", None, "답"])
    buf = io.BytesIO()
    wb.save(buf)
    n, missing = questions.import_answers(store, e["id"], buf.getvalue(), date(2027, 1, 14))
    assert (n, missing) == (1, ["Q-099"])
    q = questions.get(store, a)
    assert q["status"] == "답변받음" and q["answer"].startswith("첨부한") and q["answered_on"] == "2027-01-14"
    assert questions.act(store, [a, b], "close") == (2, [])


def test_drafts_from_confirmations(site, store):  # noqa: F811
    store.set_setting("nice_user", "cpa01")
    confirmations.set_password("cpa01", "pw!")
    e = store.engagements()[0]
    confirmations.fetch(store, e, config.load(store), TODAY, client_factory=lambda: nicebiz.Client(base=site))
    assert questions.from_confirmations(store, e, TODAY) == 2          # 불일치 1 + 미회신 1
    assert questions.from_confirmations(store, e, TODAY) == 0          # 같은 건은 다시 만들지 않는다
    qs = {q["account"]: q for q in questions.all_(store, e["id"])}
    assert "-50원 차이" in qs["외상매입금"]["question"] and qs["외상매입금"]["amount"] == 500
    assert "회신이 없습니다" in qs["미지급금"]["question"] and qs["미지급금"]["status"] == "작성"


def test_web_questions(store, monkeypatch, tmp_path):  # noqa: F811
    from auditautopilot.manage.web import make_server

    monkeypatch.setattr(mailbox.smtplib, "SMTP_SSL", FakeSMTP)
    monkeypatch.setattr(mailbox.smtplib, "SMTP", FakeSMTP)
    FakeSMTP.sent, FakeSMTP.blocked = [], set()
    store.set_setting("mail_accounts", [{"user": "me@gmail.com", "provider": "Gmail", **mailbox.PRESETS["Gmail"]}])
    store.set_setting("send_from", "me@gmail.com")
    mailbox.set_password("me@gmail.com", "app-pw")
    e = store.engagements()[0]
    srv = make_server(store, "127.0.0.1", 0, today=lambda: TODAY)
    port = srv.server_address[1]
    srv.RequestHandlerClass.allowed_hosts = {f"127.0.0.1:{port}"}
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{port}"

    class NoRedirect(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, *a, **k):
            return None
    op = urllib.request.build_opener(NoRedirect)

    def post(path, data):
        body = urllib.parse.urlencode(data, doseq=True).encode()
        try:
            r = op.open(base + path, body)
            return r.status, urllib.parse.unquote(r.headers.get("Location", ""))
        except urllib.error.HTTPError as err:
            return err.code, urllib.parse.unquote(err.headers.get("Location", ""))
    try:
        eid = e["id"]
        assert post(f"/e/{eid}/questions", {"account": "재고", "question": "재고 실사 차이 사유", "amount": "2,500,000"})[0] == 303
        assert post(f"/e/{eid}/questions", {"question": "특수관계자 거래 목록"})[0] == 303
        ids = [q["id"] for q in questions.all_(store, eid)]
        st, loc = post(f"/e/{eid}/questions/act", {"ids": ids, "action": "approve"})
        assert "작성자" in loc                                            # 승인자 이름이 없으면 막는다
        store.set_setting("author", "JP")
        assert post(f"/e/{eid}/questions/{ids[1]}", {"question": "특수관계자 거래 목록과 잔액", "account": "", "amount": ""})[0] == 303
        st, loc = post(f"/e/{eid}/questions/act", {"ids": [ids[0]], "action": "approve"})
        assert "1건 처리" in loc
        page = op.open(base + f"/e/{eid}").read().decode()
        assert "승인한 질문 1건 메일 미리보기" in page and "특수관계자 거래 목록과 잔액" in page
        prev = op.open(base + f"/e/{eid}/questions/mail").read().decode()
        assert "Q-001. [재고] 재고 실사 차이 사유 (2,500,000원)" in prev and "Q-002" not in prev
        st, loc = post(f"/e/{eid}/questions/mail/send", {})
        assert st == 303 and "1건을 보냈습니다" in loc
        [sent] = FakeSMTP.sent
        assert sent["To"] == "kim@ganada.test" and any(p.get_filename().endswith("질문서.xlsx") for p in sent.iter_attachments())
        assert questions.get(store, ids[0])["status"] == "발송" and questions.get(store, ids[1])["status"] == "작성"
        x = op.open(base + f"/e/{eid}/questions.xlsx").read()
        assert openpyxl.load_workbook(io.BytesIO(x)).active.cell(row=5, column=1).value == "Q-001"
    finally:
        srv.shutdown()
        srv.server_close()


def test_answer_by_mail(store):  # noqa: F811
    from email.message import EmailMessage
    from email.utils import make_msgid

    e = store.engagements()[0]
    a = questions.add(store, e["id"], "가지급금 증가 사유", "가지급금", today=TODAY)
    questions.act(store, [a], "approve", "JP", TODAY)
    questions.mark_sent(store, [a], TODAY)
    wb = openpyxl.load_workbook(io.BytesIO(questions.xlsx_bytes(e, questions.all_(store, e["id"]))))
    wb.active.cell(row=5, column=5, value="대표이사 단기 대여, 1월 중 상환 예정")
    buf = io.BytesIO()
    wb.save(buf)
    m = EmailMessage()
    m["From"], m["To"], m["Message-ID"] = "kim@ganada.test", "me@gmail.com", make_msgid()
    m["Subject"] = "RE: 질문 드립니다 [C001-2026-기말]"
    m.set_content("답변 드립니다.")
    m.add_attachment(buf.getvalue(), maintype="application", subtype="octet-stream", filename="(주)가나상사_기말감사_질문서.xlsx")
    r = mailbox.process_message(store, config.load(store), m.as_bytes(), TODAY)
    assert r["engagement"]["id"] == e["id"] and "질문서" in str(r["files"][0]["path"])
    assert questions.get(store, a)["status"] == "답변받음"


def test_from_confirm_without_status_returns_to_card(server, store):  # noqa: F811
    from test_manage import _open
    base, _ = server
    cid = store.add_client("C010", "가나상사", "김담당", "kim@example.com")
    eid = store.add_engagement(cid, 2026, "기말감사", "2027-01-11", "2027-01-15")
    st, h, _ = _open(base + f"/e/{eid}/questions/from-confirm", {}, {"Referer": base + f"/e/{eid}"})
    assert st == 303 and h["Location"].startswith(f"/e/{eid}?err=") and h["Location"].endswith("#questions")
