import threading
import urllib.error
import urllib.parse
import urllib.request
from datetime import date
from email import message_from_bytes

import openpyxl
import pytest

from auditautopilot.manage import config, filing, pbc, schedule
from auditautopilot.manage.db import Store
from auditautopilot.manage.web import make_server


@pytest.fixture
def store(tmp_path):
    s = Store(tmp_path / "audit.db")
    s.set_setting("root", str(tmp_path / "감사"))
    s.set_setting("firm_name", "가나회계법인")
    s.set_setting("author", "이회계")
    return s


@pytest.fixture
def eng(store):
    cid = store.add_client("C001", "한국상사", "김담당", "kim@example.com")
    # 2027-01-11 은 월요일
    eid = store.add_engagement(cid, 2026, "기말감사", "2027-01-11", "2027-01-15")
    return store.engagement(eid)


def test_templates_and_codes():
    items = pbc.build_items(pbc.TEMPLATES["일반"], ("기본자료", "중간감사"))
    assert items[0] == {"code": "기본-01", "section": "기본자료", "title": "법인등기부등본"}
    assert len([i for i in items if i["section"] == "중간감사"]) == 18
    assert "건설면허및산업재산권보유현황" in pbc.TEMPLATES["건설"]["기본자료"]


def test_load_template_from_xlsx(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["회계감사준비자료"])
    ws.append([])
    ws.append(["기본자료(작년에 제출했을 경우 변동사항만)", "중간감사", "기말감사", "재고실사/금융자산실사"])
    ws.append(["정관", "급여대장", "보험가입내역", "현금시재실사"])
    ws.append([None, "이사회의사록", None, None])
    ws.append([None, None, None, "가나회계법인"])
    p = tmp_path / "t.xlsx"
    wb.save(p)
    t = pbc.load_template(p)
    assert t == {"기본자료": ["정관"], "중간감사": ["급여대장", "이사회의사록"],
                 "기말감사": ["보험가입내역"], "실사": ["현금시재실사"]}
    with pytest.raises(ValueError):
        pbc.load_template(tmp_path / "x.xls")


def test_plan_skips_weekend_and_moves_request_earlier(eng):
    cfg = dict(config.DEFAULTS)
    steps = schedule.plan(eng, cfg)
    # D-7 = 2027-01-04(월). D-3 = 01-08(금), D-2·D-1 은 주말이라 빠진다.
    assert [(p.kind, p.on.isoformat(), p.d_n) for p in steps] == [
        ("request", "2027-01-04", 7), ("reminder", "2027-01-08", 3), ("dday", "2027-01-11", 0)]
    cfg["holidays"] = ["2027-01-04"]
    assert schedule.plan(eng, cfg)[0].on == date(2027, 1, 1)  # 공휴일이면 그 전 영업일(금)
    cfg["skip_weekends"] = False
    cfg["holidays"] = []
    assert [p.d_n for p in schedule.plan(eng, cfg)] == [7, 3, 2, 1, 0]


def test_due_drafts_flow(store, eng):
    cfg = config.load(store)
    eid = eng["id"]
    assert schedule.due_drafts(store, eng, date(2027, 1, 3), cfg) == []
    [d] = schedule.due_drafts(store, eng, date(2027, 1, 4), cfg)
    assert d.kind == "request" and d.to == "kim@example.com"
    assert "[가나회계법인] 한국상사 2026년 기말감사 자료 요청" in d.subject
    assert "[C001-2026-기말]" in d.subject
    assert "김담당님" in d.body and "[기말-01] 회계데이타" in d.body
    assert "2027년 1월 7일(목)" in d.body  # 제출기한 D-4
    # 놓친 요청은 다음 날에도 뜬다
    late = schedule.due_drafts(store, eng, date(2027, 1, 5), cfg)
    assert late[0].overdue
    schedule.record_sent(store, d, date(2027, 1, 4))
    assert all(i["status"] == "요청" for i in store.items(eid))
    assert schedule.due_drafts(store, eng, date(2027, 1, 5), cfg) == []
    # 두 항목만 남기고 수령
    items = store.items(eid)
    for i in items[2:]:
        store.update_item(i["id"], status="수령", on=date(2027, 1, 6))
    [r] = schedule.due_drafts(store, eng, date(2027, 1, 8), cfg)
    assert r.kind == "reminder" and r.item_codes == [items[0]["code"], items[1]["code"]]
    assert "미제출 2건 (현장감사 D-3)" in r.subject
    schedule.record_sent(store, r, date(2027, 1, 8))
    assert schedule.due_drafts(store, eng, date(2027, 1, 8), cfg) == []
    # 다 받으면 당일 메일도 없다
    for i in items[:2]:
        store.update_item(i["id"], status="해당없음")
    assert schedule.due_drafts(store, eng, date(2027, 1, 11), cfg) == []


def test_dday_uses_dday_body(store, eng):
    cfg = config.load(store)
    d = schedule.due_drafts(store, eng, date(2027, 1, 4), cfg)[0]
    schedule.record_sent(store, d, date(2027, 1, 4))
    [dd] = schedule.due_drafts(store, eng, date(2027, 1, 11), cfg)
    assert dd.kind == "dday" and "오늘부터 현장감사가 시작됩니다" in dd.body
    assert "(현장감사 당일)" in dd.subject


def test_folder_rules(eng):
    cfg = dict(config.DEFAULTS, root="감사")
    item = {"code": "기말-12", "title": "전직원퇴직금추계액"}
    ctx = filing.context(eng, "PBC", cfg, "퇴직금.xlsx", item, date(2027, 1, 8))
    assert filing.build_path(config.JP_RULE, ctx).as_posix() == "감사/01.한국상사/01.2026/02.기말감사/01.PBC/퇴직금.xlsx"
    assert filing.build_path(config.PROPOSED_RULE, ctx).as_posix() == (
        "감사/C001_한국상사/FY2026/2_기말/01_PBC/20270108_기말-12_전직원퇴직금추계액_퇴직금.xlsx")
    # 항목을 못 찾으면 빈 칸이 정리되고 99.미분류로 간다
    ctx2 = filing.context(eng, "미분류", cfg, "a:b?.pdf", None, date(2027, 1, 8))
    assert filing.build_path(config.PROPOSED_RULE, ctx2).as_posix() == (
        "감사/C001_한국상사/FY2026/2_기말/99_미분류/20270108_a_b_.pdf")
    with pytest.raises(ValueError):
        filing.build_path("{루트}/{없는칸}/{원본파일명}", ctx)


def test_match_item():
    items = [{"id": 1, "code": "기말-01", "title": "회계데이타"},
             {"id": 2, "code": "기말-12", "title": "전직원퇴직금추계액"},
             {"id": 3, "code": "중간-03", "title": "부가가치세 신고서"}]
    assert filing.match_item(items, "기말_12 퇴직금.xlsx")["id"] == 2
    assert filing.match_item(items, "[기말-1] 원장.xlsx")["id"] == 1
    assert filing.match_item(items, "2026 부가가치세신고서(2기).pdf")["id"] == 3
    assert filing.match_item(items, "기타.pdf") is None


def test_save_and_sweep(store, eng, tmp_path):
    cfg = config.load(store)
    r = filing.save_file(store, eng, cfg, "기말-12 퇴직금.xlsx", data=b"x", received_on=date(2027, 1, 6))
    assert r["matched"] and r["path"].exists()
    again = filing.save_file(store, eng, cfg, "기말-12 퇴직금.xlsx", data=b"y", received_on=date(2027, 1, 6))
    assert again["path"].name == "기말-12 퇴직금_v2.xlsx"
    assert next(i for i in store.items(eng["id"]) if i["code"] == "기말-12")["status"] == "수령"
    box = filing.inbox_dir(eng, cfg)
    box.mkdir(parents=True)
    (box / "모르는파일.pdf").write_bytes(b"z")
    (box / "~$임시.xlsx").write_bytes(b"z")
    [res] = filing.sweep_inbox(store, eng, cfg, date(2027, 1, 7))
    assert not res["matched"] and "99.미분류" in str(res["path"])
    assert not (box / "모르는파일.pdf").exists()
    rec = store.files(eng["id"])[0]
    item = next(i for i in store.items(eng["id"]) if i["code"] == "기말-01")
    dest = filing.refile(store, eng, cfg, rec["id"], item["id"])
    assert "01.PBC" in str(dest) and dest.exists()
    assert next(i for i in store.items(eng["id"]) if i["id"] == item["id"])["status"] == "수령"


# ── 화면 ────────────────────────────────────────────
class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *a, **k):
        return None


@pytest.fixture
def server(store):
    today = {"d": date(2027, 1, 4)}
    srv = make_server(store, "127.0.0.1", 0, today=lambda: today["d"])
    port = srv.server_address[1]
    srv.RequestHandlerClass.allowed_hosts = {f"127.0.0.1:{port}"}
    t = threading.Thread(target=srv.serve_forever, daemon=True)
    t.start()
    yield f"http://127.0.0.1:{port}", today
    srv.shutdown()
    srv.server_close()


def _open(url, data=None, headers=None):
    opener = urllib.request.build_opener(NoRedirect)
    body = urllib.parse.urlencode(data, doseq=True).encode() if isinstance(data, dict) else data
    req = urllib.request.Request(url, data=body, headers=headers or {})
    try:
        r = opener.open(req)
        return r.status, r.headers, r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.headers, e.read()


def _multipart(fields, files):
    b = "----aapboundary"
    out = b""
    for k, v in fields:
        out += f'--{b}\r\nContent-Disposition: form-data; name="{k}"\r\n\r\n{v}\r\n'.encode()
    for k, name, data in files:
        out += (f'--{b}\r\nContent-Disposition: form-data; name="{k}"; filename="{name}"\r\n'
                f"Content-Type: application/octet-stream\r\n\r\n").encode() + data + b"\r\n"
    out += f"--{b}--\r\n".encode()
    return out, {"Content-Type": f"multipart/form-data; boundary={b}"}


def test_web_flow(server, store):
    base, today = server
    st, _, body = _open(base + "/")
    assert st == 200 and "등록된 감사가 없습니다" in body.decode()
    st, h, _ = _open(base + "/clients", {"code": "C002", "name": "다라건설", "industry": "건설",
                                         "contact_name": "박담당", "contact_email": "park@example.com"})
    assert st == 303
    cid = store.clients()[0]["id"]
    data, hdr = _multipart([("client_id", cid), ("fiscal_year", 2026), ("kind", "기말감사"),
                            ("fieldwork_start", "2027-01-11"), ("fieldwork_end", "2027-01-15"),
                            ("sections", "기말감사"), ("sections", "실사")], [])
    st, h, _ = _open(base + "/engagements", data, hdr)
    assert st == 303 and h["Location"].startswith("/e/")
    eid = store.engagements()[0]["id"]
    assert {i["section"] for i in store.items(eid)} == {"기말감사", "실사"}

    st, _, body = _open(base + "/")
    page = body.decode()
    assert "다라건설" in page and "최초 요청 (D-7)" in page
    st, _, body = _open(base + f"/e/{eid}/mail/request/2027-01-04")
    assert "박담당님" in body.decode()
    st, h, eml = _open(base + f"/e/{eid}/mail/request/2027-01-04.eml")
    assert st == 200 and h["Content-Type"] == "message/rfc822"
    msg = message_from_bytes(eml)
    assert msg["X-Unsent"] == "1" and msg["To"] == "park@example.com"
    assert any(p.get_filename() and p.get_filename().endswith(".xlsx") for p in msg.walk())
    st, _, _ = _open(base + f"/e/{eid}/mail/request/2027-01-04/sent", {})
    assert st == 303
    assert all(i["status"] == "요청" for i in store.items(eid))
    assert "오늘 보낼 메일이 없습니다" in _open(base + "/")[2].decode()

    item = store.items(eid)[0]
    st, _, _ = _open(base + f"/e/{eid}/items", {f"status_{item['id']}": "해당없음", f"note_{item['id']}": "없음"})
    assert store.items(eid)[0]["status"] == "해당없음" and store.items(eid)[0]["note"] == "없음"

    second = store.items(eid)[1]
    data, hdr = _multipart([], [("files", f"{second['code']} 자료.xlsx", b"abc"), ("files", "기타.pdf", b"d")])
    st, h, _ = _open(base + f"/e/{eid}/upload", data, hdr)
    assert st == 303 and "%ED%95%AD%EB%AA%A9%20%EC%97%B0%EA%B2%B0%201" in h["Location"]  # '항목 연결 1'
    assert len(store.files(eid)) == 2

    # 전자조회 회신현황 → 대사표 내려받기 (가짜 거래처)
    import io

    import openpyxl
    wb = openpyxl.Workbook()
    for r in [["기업명", "구분", "발송금액", "회신금액", "일치여부", "회신일시"],
              ["가나상사", "외상매출금", 1000, 900, "N", "2027-01-05"], ["다라상사", "미지급금", 500, None, "", None]]:
        wb.active.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    data, hdr = _multipart([], [("files", "회신현황.xlsx", buf.getvalue())])
    st, h, x = _open(base + f"/e/{eid}/confirm", data, hdr)
    assert st == 200 and "attachment" in h["Content-Disposition"]
    out = openpyxl.load_workbook(io.BytesIO(x))
    assert [r[0] for r in out["불일치_차이조정"].iter_rows(values_only=True)][1] == "가나상사"
    assert [r[0] for r in out["미회신_대체절차"].iter_rows(values_only=True)][1] == "다라상사"
    data, hdr = _multipart([], [("files", "메모.txt", b"x")])
    assert _open(base + f"/e/{eid}/confirm", data, hdr)[0] == 303     # 엑셀이 아니면 설명과 함께 되돌아감

    st, h, x = _open(base + f"/e/{eid}/pbc.xlsx")
    assert st == 200 and x[:2] == b"PK"

    st, _, _ = _open(base + "/settings", {"firm_name": "가나회계법인", "author": "이회계", "folder_rule": "{루트}/{없는칸}"})
    assert st == 303 and config.load(store)["folder_rule"] == config.JP_RULE  # 틀린 규칙은 저장 안 됨
    st, _, _ = _open(base + "/settings", {"folder_rule": config.PROPOSED_RULE, "skip_weekends": "1",
                                          "holidays": "2027-01-08", "request_offset": "7"})
    assert store.get_setting("folder_rule") == config.PROPOSED_RULE
    assert store.get_setting("holidays") == ["2027-01-08"]


def test_web_rejects_cross_site_post(server, store):
    base, _ = server
    st, _, _ = _open(base + "/clients", {"code": "X", "name": "Y"}, {"Origin": "http://evil.example"})
    assert st == 403 and store.clients() == []
    st, _, _ = _open(base + "/", headers={"Host": "evil.example"})
    assert st == 403
