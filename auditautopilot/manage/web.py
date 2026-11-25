"""관리 화면 (파이썬 표준 라이브러리 http.server).

  python -m auditautopilot serve

브라우저에서 http://127.0.0.1:8765 를 연다. 기본으로 이 PC에서만 접속된다.
"""

from __future__ import annotations

import html
import json
import re
import threading
import urllib.parse
import webbrowser
from datetime import date
from email import policy
from email.parser import BytesParser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from . import config, filing, mailbox, mailer, pbc, schedule
from .db import KINDS, STATUSES, Store

MAX_UPLOAD = 300 * 1024 * 1024
STATUS_COLOR = {"미요청": "#8a8f98", "요청": "#2f6fdf", "일부수령": "#c07a00", "수령": "#1d8a4e",
                "검토완료": "#11613a", "보완요청": "#c2362b", "해당없음": "#b5b9c0"}

CSS = """
:root{--bg:#f6f7f9;--card:#fff;--ink:#1d2330;--mute:#677083;--line:#e2e5ea;--accent:#2f6fdf;--warn:#c2362b}
*{box-sizing:border-box}body{margin:0;font:14px/1.5 -apple-system,"Malgun Gothic","맑은 고딕",sans-serif;
background:var(--bg);color:var(--ink)}a{color:var(--accent);text-decoration:none}a:hover{text-decoration:underline}
header{background:#1d2330;color:#fff;padding:10px 20px;display:flex;gap:20px;align-items:center}
header a{color:#cfd6e4}header b{color:#fff;margin-right:12px}main{max-width:1200px;margin:0 auto;padding:20px}
h1{font-size:20px;margin:4px 0 16px}h2{font-size:16px;margin:24px 0 8px}
.card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:16px;margin-bottom:16px}
table{border-collapse:collapse;width:100%}th,td{text-align:left;padding:6px 8px;border-bottom:1px solid var(--line);
vertical-align:top}th{font-weight:600;color:var(--mute);font-size:12px;background:#fafbfc}
input,select,textarea{font:inherit;padding:5px 7px;border:1px solid #c9ced6;border-radius:5px;background:#fff}
textarea{width:100%}button,.btn{font:inherit;padding:6px 12px;border:1px solid var(--accent);background:var(--accent);
color:#fff;border-radius:5px;cursor:pointer;display:inline-block}button.ghost,.btn.ghost{background:#fff;color:var(--accent)}
.pill{display:inline-block;padding:1px 8px;border-radius:10px;color:#fff;font-size:12px;white-space:nowrap}
.mute{color:var(--mute)}.warn{color:var(--warn)}.row{display:flex;gap:12px;flex-wrap:wrap;align-items:end}
.row label{display:flex;flex-direction:column;gap:3px;font-size:12px;color:var(--mute)}
.msg{padding:10px 14px;border-radius:6px;background:#e8f1ff;margin-bottom:16px}.msg.err{background:#fdecea}
pre{white-space:pre-wrap;font-family:inherit;background:#fafbfc;border:1px solid var(--line);padding:10px;border-radius:6px;margin:6px 0}
.bar{height:8px;background:#e9ecf1;border-radius:4px;overflow:hidden;min-width:120px}
.bar i{display:block;height:100%;background:#1d8a4e}
button,.btn{white-space:nowrap}.card{overflow-x:auto}td input:not([type=checkbox]):not([type=hidden]){width:100%;min-width:56px}
"""


def esc(v) -> str:
    return html.escape("" if v is None else str(v))


def pill(status: str) -> str:
    return f'<span class="pill" style="background:{STATUS_COLOR.get(status, "#888")}">{esc(status)}</span>'


def page(title: str, body: str, flash: str = "", error: bool = False) -> bytes:
    msg = f'<div class="msg{" err" if error else ""}">{esc(flash)}</div>' if flash else ""
    return f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>{esc(title)} · auditautopilot</title>
<style>{CSS}</style></head><body><header><b>auditautopilot</b><a href="/">현황</a><a href="/clients">회사</a>
<a href="/engagements/new">감사 등록</a><a href="/settings">설정</a></header>
<main>{msg}<h1>{esc(title)}</h1>{body}</main></body></html>""".encode("utf-8")


def options(values, selected=None) -> str:
    return "".join(f'<option value="{esc(v)}"{" selected" if str(v) == str(selected) else ""}>{esc(lbl)}</option>'
                   for v, lbl in values)


class App:
    """화면이 쓰는 상태. 테스트에서는 today 를 바꿔 날짜를 흉내 낸다."""

    def __init__(self, store: Store, today=date.today):
        self.store = store
        self.today = today
        self.lock = threading.Lock()
        self.log: list[str] = []   # 백그라운드 메일 확인·발송 기록 (최근 것부터)

    def note(self, lines) -> None:
        from datetime import datetime
        stamp = datetime.now().strftime("%m-%d %H:%M")
        for line in lines:
            self.log.insert(0, f"{stamp} {line}")
        del self.log[50:]

    def cfg(self) -> dict:
        return config.load(self.store)


class Handler(BaseHTTPRequestHandler):
    app: App
    allowed_hosts: set[str] = set()
    server_version = "auditautopilot"

    # ── 공통 ──────────────────────────────────────────
    def log_message(self, fmt, *args):  # 콘솔을 조용히
        pass

    def _send(self, body: bytes, ctype="text/html; charset=utf-8", status=HTTPStatus.OK, headers=None):
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _redirect(self, to: str, flash: str = "", error: bool = False):
        if flash:
            to, hash_, frag = to.partition("#")      # 메시지는 #앞에 붙여야 화면에 보인다
            sep = "&" if "?" in to else "?"
            to += f"{sep}{'err' if error else 'msg'}={urllib.parse.quote(flash)}" + hash_ + frag
        self.send_response(HTTPStatus.SEE_OTHER)
        self.send_header("Location", to)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def _download(self, data: bytes, filename: str, ctype: str):
        q = urllib.parse.quote(filename)
        self._send(data, ctype, headers={"Content-Disposition": f"attachment; filename*=UTF-8''{q}"})

    def _host_ok(self) -> bool:
        host = (self.headers.get("Host") or "").lower()
        if self.allowed_hosts and host not in self.allowed_hosts:
            return False
        origin = self.headers.get("Origin")
        if self.command == "POST" and origin and origin.lower() != f"http://{host}":
            return False
        return True

    def _query(self) -> dict:
        return {k: v[-1] for k, v in urllib.parse.parse_qs(urllib.parse.urlsplit(self.path).query).items()}

    def _flash(self) -> tuple[str, bool]:
        q = self._query()
        return (q.get("err") or q.get("msg") or "", "err" in q)

    def _form(self) -> tuple[dict, list[tuple[str, bytes]]]:
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_UPLOAD:
            raise ValueError("파일이 너무 큽니다 (최대 300MB). 수신함 폴더에 직접 넣어 주세요.")
        raw = self.rfile.read(length)
        ctype = self.headers.get("Content-Type", "")
        if ctype.startswith("multipart/form-data"):
            msg = BytesParser(policy=policy.default).parsebytes(
                b"Content-Type: " + ctype.encode("latin-1") + b"\r\n\r\n" + raw)
            fields, files = {}, []
            lists: dict[str, list[str]] = {}
            for part in msg.iter_parts():
                name = part.get_param("name", header="content-disposition")
                fname = part.get_filename()
                data = part.get_payload(decode=True) or b""
                if fname:
                    if data:
                        files.append((Path(fname.replace("\\", "/")).name, data))
                elif name:
                    value = data.decode(part.get_content_charset() or "utf-8")
                    fields[name] = value
                    lists.setdefault(f"{name}[]", []).append(value)
            return fields | lists, files
        parsed = urllib.parse.parse_qs(raw.decode("utf-8"), keep_blank_values=True)
        return {k: v[-1] for k, v in parsed.items()} | {f"{k}[]": v for k, v in parsed.items()}, []

    def do_GET(self):
        self._dispatch("GET")

    def do_POST(self):
        self._dispatch("POST")

    ROUTES = [
        ("GET", r"/", "home"),
        ("GET", r"/clients", "clients"),
        ("POST", r"/clients", "client_add"),
        ("POST", r"/clients/(\d+)", "client_update"),
        ("GET", r"/engagements/new", "engagement_new"),
        ("POST", r"/engagements", "engagement_add"),
        ("GET", r"/e/(\d+)", "engagement_view"),
        ("POST", r"/e/(\d+)/dates", "engagement_dates"),
        ("POST", r"/e/(\d+)/delete", "engagement_delete"),
        ("POST", r"/e/(\d+)/items", "items_save"),
        ("POST", r"/e/(\d+)/items/add", "item_add"),
        ("POST", r"/e/(\d+)/upload", "upload"),
        ("POST", r"/e/(\d+)/sweep", "sweep"),
        ("POST", r"/e/(\d+)/files/(\d+)", "refile"),
        ("GET", r"/e/(\d+)/pbc\.xlsx", "pbc_xlsx"),
        ("GET", r"/e/(\d+)/mail/(request|reminder|dday)/(\d{4}-\d\d-\d\d)", "mail_preview"),
        ("GET", r"/e/(\d+)/mail/(request|reminder|dday)/(\d{4}-\d\d-\d\d)\.eml", "mail_eml"),
        ("POST", r"/e/(\d+)/mail/(request|reminder|dday)/(\d{4}-\d\d-\d\d)/sent", "mail_sent"),
        ("POST", r"/e/(\d+)/mail/(request|reminder|dday)/(\d{4}-\d\d-\d\d)/send", "mail_send"),
        ("POST", r"/e/(\d+)/links/(\d+)/done", "link_done"),
        ("POST", r"/mail/fetch", "mail_fetch"),
        ("GET", r"/settings", "settings"),
        ("POST", r"/settings", "settings_save"),
        ("POST", r"/settings/mail", "mail_account_save"),
        ("POST", r"/settings/mail-test", "mail_test"),
        ("POST", r"/settings/mail-delete", "mail_account_delete"),
        ("POST", r"/settings/mail-options", "mail_options_save"),
    ]

    def _dispatch(self, method: str):
        if not self._host_ok():
            self._send("허용되지 않은 접속입니다.".encode(), "text/plain; charset=utf-8", HTTPStatus.FORBIDDEN)
            return
        path = urllib.parse.unquote(urllib.parse.urlsplit(self.path).path)
        for m, pattern, name in self.ROUTES:
            match = re.fullmatch(pattern, path)
            if m == method and match:
                try:
                    with self.app.lock:
                        getattr(self, name)(*match.groups())
                except (ValueError, KeyError) as exc:
                    back = self.headers.get("Referer") or "/"
                    if method == "POST":
                        self._redirect(urllib.parse.urlsplit(back).path or "/", str(exc), error=True)
                    else:
                        self._send(page("오류", f"<p>{esc(exc)}</p>"), status=HTTPStatus.BAD_REQUEST)
                return
        self._send(page("없는 페이지", '<p><a href="/">처음으로</a></p>'), status=HTTPStatus.NOT_FOUND)

    @property
    def store(self) -> Store:
        return self.app.store

    def _eng(self, eid) -> dict:
        e = self.store.engagement(int(eid))
        if e is None:
            raise ValueError("감사 건을 찾지 못했습니다.")
        return e

    # ── 현황 ──────────────────────────────────────────
    def home(self):
        cfg, today = self.app.cfg(), self.app.today()
        due = schedule.all_due(self.store, today, cfg)
        warn = ""
        if not cfg["firm_name"] or not cfg["author"]:
            warn = '<p class="warn">설정에서 회계법인 이름과 작성자를 먼저 넣어 주세요. 메일 문구에 들어갑니다.</p>'
        rows = []
        for e, d in due:
            base = f"/e/{e['id']}/mail/{d.kind}/{d.planned_on.isoformat()}"
            late = ' <span class="warn">(늦음)</span>' if d.overdue else ""
            rows.append(f"<tr><td><a href='/e/{e['id']}'>{esc(e['client_name'])}</a> {esc(e['kind'])}</td>"
                        f"<td>{esc(d.label)}{late}</td><td>{len(d.item_codes)}건</td><td>{esc(d.to) or '<span class=warn>수신자 없음</span>'}</td>"
                        f"<td><a class='btn ghost' href='{base}'>미리보기</a></td></tr>")
        mail_html = ("<table><tr><th>감사</th><th>메일</th><th>항목</th><th>받는 사람</th><th></th></tr>"
                     + "".join(rows) + "</table>") if rows else '<p class="mute">오늘 보낼 메일이 없습니다.</p>'
        erows = []
        for e in self.store.engagements():
            nxt = next((p for p in schedule.plan(e, cfg) if p.on >= today
                        and not self.store.mail_sent(e["id"], p.kind, p.on)), None)
            done = e["total"] - e["outstanding"]
            pct = int(done / e["total"] * 100) if e["total"] else 0
            erows.append(
                f"<tr><td><a href='/e/{e['id']}'>{esc(e['client_code'])} {esc(e['client_name'])}</a></td>"
                f"<td>{esc(e['fiscal_year'])} {esc(e['kind'])}</td>"
                f"<td>{esc(e['fieldwork_start'])} ~ {esc(e['fieldwork_end'])}</td>"
                f"<td><div class='bar'><i style='width:{pct}%'></i></div><span class='mute'>{done}/{e['total']}</span></td>"
                f"<td>{e['outstanding']}건</td><td>{esc(f'{nxt.on} {nxt.label}') if nxt else '-'}</td></tr>")
        eng_html = ("<table><tr><th>회사</th><th>감사</th><th>현장감사</th><th>수령</th><th>미제출</th><th>다음 메일</th></tr>"
                    + "".join(erows) + "</table>") if erows else (
            '<p class="mute">등록된 감사가 없습니다. <a href="/clients">회사</a>를 먼저 등록하고 '
            '<a href="/engagements/new">감사를 등록</a>하세요.</p>')
        if cfg["mail_accounts"]:
            auto = "켜짐" if cfg["auto_send"] else "꺼짐(미리보기에서 직접 보내기)"
            users = ", ".join(a["user"] for a in cfg["mail_accounts"])
            mail_html += (f"<form method='post' action='/mail/fetch' class='row' style='margin-top:10px'>"
                          f"<button class='ghost'>받은 메일 지금 확인</button>"
                          f"<span class='mute'>{esc(users)} · {cfg['check_minutes']}분마다 확인 · 자동 발송 {auto}</span></form>")
        else:
            mail_html += "<p class='mute'><a href='/settings#mail'>메일 계정을 연결</a>하면 여기서 바로 보내고, 받은 첨부를 자동으로 저장합니다.</p>"
        if self.app.log:
            mail_html += "<h2>최근 메일 작업</h2><pre>" + esc("\n".join(self.app.log[:15])) + "</pre>"
        body = (f"{warn}<div class='card'><h2 style='margin-top:0'>오늘({today.isoformat()}) 보낼 메일</h2>{mail_html}</div>"
                f"<div class='card'><h2 style='margin-top:0'>감사 진행 현황</h2>{eng_html}</div>")
        self._send(page("현황", body, *self._flash()))

    # ── 회사 ──────────────────────────────────────────
    def clients(self):
        cfg = self.app.cfg()
        senders = [("", "기본 계정")] + [(a["user"], a["user"]) for a in cfg["mail_accounts"]]
        rows = []
        for c in self.store.clients():
            fid = f"c{c['id']}"
            rows.append(
                f"<tr><td><form id='{fid}' method='post' action='/clients/{c['id']}'></form>"
                f"<input form='{fid}' name='seq' value='{esc(c['seq'])}' style='width:48px;min-width:0'></td>"
                f"<td>{esc(c['code'])}</td>"
                f"<td><input form='{fid}' name='name' value='{esc(c['name'])}'></td>"
                f"<td><select form='{fid}' name='industry'>{options([(k, k) for k in pbc.TEMPLATES], c['industry'])}</select></td>"
                f"<td><input form='{fid}' name='contact_name' value='{esc(c['contact_name'])}'></td>"
                f"<td><input form='{fid}' name='contact_email' value='{esc(c['contact_email'])}'></td>"
                f"<td><input form='{fid}' name='cc' value='{esc(c['cc'])}' placeholder='쉼표로 구분'></td>"
                f"<td><select form='{fid}' name='send_from'>{options(senders, c['send_from'])}</select></td>"
                f"<td><button form='{fid}' class='ghost'>저장</button></td></tr>")
        body = f"""<div class="card"><h2 style="margin-top:0">새 회사</h2>
<form method="post" action="/clients" class="row">
<label>회사코드<input name="code" required placeholder="C001" size="6"></label>
<label>회사명<input name="name" required></label>
<label>업종 템플릿<select name="industry">{options([(k, k) for k in pbc.TEMPLATES])}</select></label>
<label>담당자<input name="contact_name" size="8"></label>
<label>이메일<input name="contact_email" type="email"></label>
<label>참조(CC)<input name="cc" placeholder="쉼표로 구분"></label>
<button>등록</button></form>
<p class="mute">순번은 폴더 규칙의 {{회사순번}}(예: 01.회사명)에 쓰입니다.</p></div>
<div class="card"><table><tr><th>순번</th><th>코드</th><th>회사명</th><th>업종</th><th>담당자</th><th>이메일</th><th>참조</th><th>보내는 계정</th><th></th></tr>
{''.join(rows) or '<tr><td colspan=9 class=mute>아직 없습니다.</td></tr>'}</table></div>"""
        self._send(page("회사", body, *self._flash()))

    def client_add(self):
        f, _ = self._form()
        self.store.add_client(f.get("code", ""), f.get("name", ""), f.get("contact_name", ""),
                              f.get("contact_email", ""), f.get("cc", ""), f.get("industry", "일반"))
        self._redirect("/clients", "회사를 등록했습니다.")

    def client_update(self, cid):
        f, _ = self._form()
        self.store.update_client(int(cid), name=f.get("name", "").strip(), seq=int(f.get("seq") or 1),
                                 contact_name=f.get("contact_name", ""), contact_email=f.get("contact_email", ""),
                                 cc=f.get("cc", ""), industry=f.get("industry", "일반"),
                                 send_from=f.get("send_from", ""))
        self._redirect("/clients", "저장했습니다.")

    # ── 감사 등록 ─────────────────────────────────────
    def engagement_new(self):
        cs = self.store.clients()
        if not cs:
            self._redirect("/clients", "회사를 먼저 등록하세요.", error=True)
            return
        year = self.app.today().year
        secs = "".join(f"<label style='flex-direction:row;gap:4px'><input type='checkbox' name='sections' value='{s}' "
                       f"data-s='{s}'>{s}</label>" for s in pbc.SECTIONS)
        defaults = json.dumps({k: list(v) for k, v in pbc.DEFAULT_SECTIONS.items()}, ensure_ascii=False)
        body = f"""<div class="card"><form method="post" action="/engagements" enctype="multipart/form-data">
<div class="row">
<label>회사<select name="client_id">{options([(c['id'], f"{c['code']} {c['name']}") for c in cs])}</select></label>
<label>사업연도<input name="fiscal_year" type="number" value="{year}" style="width:90px"></label>
<label>감사구분<select name="kind" id="kind">{options([(k, k) for k in KINDS])}</select></label>
<label>현장감사 시작일<input name="fieldwork_start" type="date" required></label>
<label>종료일<input name="fieldwork_end" type="date" required></label>
<label>연도 순번<input name="year_seq" type="number" value="1" style="width:70px"></label>
</div><h2>요청 목록에 넣을 구분</h2><div class="row">{secs}</div>
<h2>요청 목록 양식 (선택)</h2>
<p class="mute">비워 두면 회사 업종의 기본 템플릿(감사준비자료 일반·건설)을 씁니다.
같은 4열 양식(기본자료 | 중간감사 | 기말감사 | 재고실사/금융자산실사)의 .xlsx 를 올리면 그 목록을 씁니다.</p>
<input type="file" name="template" accept=".xlsx"><p><button>등록하고 요청 목록 만들기</button></p></form></div>
<script>const D={defaults};const k=document.getElementById('kind');function f(){{document.querySelectorAll('[data-s]')
.forEach(c=>c.checked=D[k.value].includes(c.dataset.s))}}k.onchange=f;f();</script>"""
        self._send(page("감사 등록", body, *self._flash()))

    def engagement_add(self):
        f, files = self._form()
        sections = [s for s in pbc.SECTIONS if s in (f.get("sections[]") or [])] or None
        template = None
        if files:
            import tempfile
            with tempfile.TemporaryDirectory() as d:
                p = Path(d) / "template.xlsx"
                p.write_bytes(files[0][1])
                template = pbc.load_template(p)
        eid = self.store.add_engagement(int(f["client_id"]), int(f["fiscal_year"]), f["kind"],
                                        f["fieldwork_start"], f["fieldwork_end"], sections=sections,
                                        template=template, year_seq=int(f.get("year_seq") or 1))
        e = self._eng(eid)
        filing.inbox_dir(e, self.app.cfg()).mkdir(parents=True, exist_ok=True)
        self._redirect(f"/e/{eid}", f"감사를 등록하고 요청 항목 {len(self.store.items(eid))}개를 만들었습니다.")

    # ── 감사 상세 ─────────────────────────────────────
    def engagement_view(self, eid):
        e, cfg, today = self._eng(eid), self.app.cfg(), self.app.today()
        items = self.store.items(e["id"])
        steps = []
        for p in schedule.plan(e, cfg):
            sent = self.store.mail_sent(e["id"], p.kind, p.on) or (p.kind == "request" and self.store.request_sent(e["id"]))
            state = "보냄" if sent else ("오늘" if p.on == today else ("지남" if p.on < today else "예정"))
            link = f"/e/{e['id']}/mail/{p.kind}/{p.on.isoformat()}"
            steps.append(f"<tr><td>{p.on.isoformat()}</td><td>{esc(p.label)}</td><td>{state}</td>"
                         f"<td><a href='{link}'>문구 보기</a></td></tr>")
        counts = {s: sum(1 for i in items if i["status"] == s) for s in STATUSES}
        summary = " ".join(f"{pill(s)} {n}" for s, n in counts.items() if n)
        status_opts = [(s, s) for s in STATUSES]
        irows = []
        for i in items:
            irows.append(
                f"<tr><td>{esc(i['code'])}</td><td class='mute'>{esc(i['section'])}</td><td>{esc(i['title'])}</td>"
                f"<td><select name='status_{i['id']}'>{options(status_opts, i['status'])}</select></td>"
                f"<td class='mute'>{esc(i['requested_on'] or '')}</td><td class='mute'>{esc(i['received_on'] or '')}</td>"
                f"<td><input name='note_{i['id']}' value='{esc(i['note'])}' style='width:100%'></td></tr>")
        item_opts = [("", "(미분류)")] + [(i["id"], f"{i['code']} {i['title'][:20]}") for i in items]
        frows = []
        for fr in self.store.files(e["id"]):
            frows.append(
                f"<tr><td>{esc(fr['received_on'])}</td><td>{esc(fr['original_name'])}</td>"
                f"<td class='mute' style='word-break:break-all'>{esc(fr['path'])}</td>"
                f"<td><form method='post' action='/e/{e['id']}/files/{fr['id']}' class='row'>"
                f"<select name='item_id'>{options(item_opts, fr['item_id'] or '')}</select>"
                f"<button class='ghost'>옮기기</button></form></td></tr>")
        sec_opts = [(s, s) for s in pbc.SECTIONS]
        inbox = filing.inbox_dir(e, cfg)
        lrows = "".join(
            f"<tr><td>{esc(l['received_on'])}</td><td>{esc(l['sender'])}<br><span class='mute'>{esc(l['subject'])}</span></td>"
            f"<td style='word-break:break-all'><a href='{esc(l['url'])}' target='_blank' rel='noopener noreferrer'>{esc(l['url'])}</a></td>"
            f"<td><form method='post' action='/e/{e['id']}/links/{l['id']}/done'><button class='ghost'>받아서 수신함에 넣음</button></form></td></tr>"
            for l in self.store.links(e["id"]))
        links_html = (f"<h2>대용량 링크 (직접 내려받기 필요)</h2><p class='mute'>링크에서 내려받아 위 수신함 폴더에 넣고 "
                      f"[수신함 정리]를 누르세요.</p><table><tr><th>받은 날</th><th>보낸 사람·제목</th><th>링크</th><th></th></tr>"
                      f"{lrows}</table>") if lrows else ""
        body = f"""<p class="mute">{esc(e['client_code'])} · 현장감사 {esc(e['fieldwork_start'])} ~ {esc(e['fieldwork_end'])}
· 담당자 {esc(e['contact_name'] or '-')} &lt;{esc(e['contact_email'] or '이메일 없음')}&gt;</p>
<div class="card"><h2 style="margin-top:0">메일 일정</h2>
<table><tr><th>날짜</th><th>메일</th><th>상태</th><th></th></tr>{''.join(steps)}</table>
<form method="post" action="/e/{e['id']}/dates" class="row" style="margin-top:10px">
<label>현장감사 시작일<input type="date" name="fieldwork_start" value="{esc(e['fieldwork_start'])}"></label>
<label>종료일<input type="date" name="fieldwork_end" value="{esc(e['fieldwork_end'])}"></label>
<button class="ghost">일정 바꾸기</button></form></div>
<div class="card"><h2 style="margin-top:0">자료요청 목록</h2><p>{summary}
<a class="btn ghost" href="/e/{e['id']}/pbc.xlsx">엑셀로 받기</a></p>
<form method="post" action="/e/{e['id']}/items"><table>
<tr><th>요청번호</th><th>구분</th><th>자료명</th><th>상태</th><th>요청일</th><th>수령일</th><th>비고(메일에 함께 나감)</th></tr>
{''.join(irows)}</table><p><button>상태 저장</button></p></form>
<form method="post" action="/e/{e['id']}/items/add" class="row">
<label>구분<select name="section">{options(sec_opts)}</select></label>
<label>자료명<input name="title" required size="40"></label><button class="ghost">항목 추가</button></form></div>
<div class="card"><h2 style="margin-top:0">받은 자료</h2>
<form method="post" action="/e/{e['id']}/upload" enctype="multipart/form-data" class="row">
<input type="file" name="files" multiple><button>올려서 정리</button></form>
<p class="mute">메일로 받은 첨부나 대용량 링크로 받은 파일을 여기에 올리거나, 아래 수신함 폴더에 넣고
[수신함 정리]를 누르면 폴더 규칙대로 저장됩니다. 파일 이름에 요청번호(예: 중간-03)나 자료명이 있으면 해당 항목이 '수령'으로 바뀝니다.<br>
수신함: <code>{esc(inbox)}</code></p>
<form method="post" action="/e/{e['id']}/sweep"><button class="ghost">수신함 정리</button></form>
{links_html}
<table style="margin-top:10px"><tr><th>받은 날</th><th>원본 파일명</th><th>저장 위치</th><th>요청 항목</th></tr>
{''.join(frows) or '<tr><td colspan=4 class=mute>아직 없습니다.</td></tr>'}</table></div>
<form method="post" action="/e/{e['id']}/delete" onsubmit="return confirm('이 감사 건과 요청 목록을 지웁니다. 저장된 파일은 지우지 않습니다.')">
<button class="ghost" style="color:var(--warn);border-color:var(--warn)">감사 건 삭제</button></form>"""
        self._send(page(f"{e['client_name']} {e['fiscal_year']} {e['kind']}", body, *self._flash()))

    def engagement_dates(self, eid):
        f, _ = self._form()
        if f.get("fieldwork_end", "") < f.get("fieldwork_start", ""):
            raise ValueError("종료일이 시작일보다 빠릅니다.")
        self.store.update_engagement(int(eid), fieldwork_start=f["fieldwork_start"], fieldwork_end=f["fieldwork_end"])
        self._redirect(f"/e/{eid}", "일정을 바꿨습니다.")

    def engagement_delete(self, eid):
        self.store.delete_engagement(int(eid))
        self._redirect("/", "감사 건을 지웠습니다.")

    def items_save(self, eid):
        f, _ = self._form()
        today = self.app.today()
        n = 0
        for i in self.store.items(int(eid)):
            st, note = f.get(f"status_{i['id']}"), f.get(f"note_{i['id']}")
            if (st is not None and st != i["status"]) or (note is not None and note != i["note"]):
                self.store.update_item(i["id"], status=st, note=note, on=today)
                n += 1
        self._redirect(f"/e/{eid}", f"{n}개 항목을 저장했습니다.")

    def item_add(self, eid):
        f, _ = self._form()
        if not f.get("title", "").strip():
            raise ValueError("자료명을 넣어 주세요.")
        self.store.add_item(int(eid), f.get("section", "기말감사"), f["title"])
        self._redirect(f"/e/{eid}", "항목을 추가했습니다.")

    def upload(self, eid):
        e, cfg = self._eng(eid), self.app.cfg()
        _, files = self._form()
        if not files:
            raise ValueError("올릴 파일을 고르세요.")
        res = [filing.save_file(self.store, e, cfg, name, data=data, received_on=self.app.today(), source="업로드")
               for name, data in files]
        matched = sum(r["matched"] for r in res)
        self._redirect(f"/e/{eid}", f"{len(res)}개 저장 (항목 연결 {matched}개, 미분류 {len(res) - matched}개)")

    def sweep(self, eid):
        e = self._eng(eid)
        res = filing.sweep_inbox(self.store, e, self.app.cfg(), self.app.today())
        matched = sum(r["matched"] for r in res)
        self._redirect(f"/e/{eid}", f"수신함에서 {len(res)}개 정리 (항목 연결 {matched}개)")

    def refile(self, eid, fid):
        f, _ = self._form()
        item_id = int(f["item_id"]) if f.get("item_id") else None
        dest = filing.refile(self.store, self._eng(eid), self.app.cfg(), int(fid), item_id)
        self._redirect(f"/e/{eid}", f"옮김: {dest}")

    def pbc_xlsx(self, eid):
        e = self._eng(eid)
        data = mailer.pbc_xlsx_bytes(self.store.items(e["id"]))
        self._download(data, f"{e['client_name']}_{e['fiscal_year']}_{e['kind']}_자료요청목록.xlsx", mailer.XLSX[0] + "/" + mailer.XLSX[1])

    # ── 메일 ──────────────────────────────────────────
    def _draft(self, eid, kind, on):
        e, cfg = self._eng(eid), self.app.cfg()
        planned = next((p for p in schedule.plan(e, cfg) if p.kind == kind and p.on.isoformat() == on), None)
        if planned is None:
            raise ValueError("일정에 없는 메일입니다. 현장감사 일정이 바뀌었는지 확인하세요.")
        if kind == "request":
            items = [i for i in self.store.items(e["id"]) if i["status"] == "미요청"] or self.store.outstanding(e["id"])
        else:
            items = self.store.outstanding(e["id"])
        return e, cfg, items, schedule.build_draft(e, items, planned, cfg)

    def mail_preview(self, eid, kind, on):
        e, cfg, items, d = self._draft(eid, kind, on)
        sent = self.store.mail_sent(e["id"], kind, date.fromisoformat(on))
        base = f"/e/{e['id']}/mail/{kind}/{on}"
        attach = "첨부: 자료요청목록.xlsx" if kind == "request" else ""
        after = "요청 항목이 '요청' 상태로 바뀌고 " if kind == "request" else ""
        sender = mailbox.sender_for(cfg, e)
        send_btn = (f'<form method="post" action="{base}/send" style="display:inline" '
                    f'onsubmit="return confirm({esc(json.dumps(f"{sender} 계정으로 {d.to} 에게 지금 보냅니다.", ensure_ascii=False))})">'
                    f'<button>지금 보내기</button></form> '
                    if cfg["mail_accounts"] and d.to else "")
        actions = ('<p class="mute">이미 보낸 메일입니다.</p>' if sent else
                   f"""<p>{send_btn}<a class="btn ghost" href="{base}.eml">아웃룩으로 열기 (.eml)</a>
<form method="post" action="{base}/sent" style="display:inline"><button class="ghost">보냄으로 표시</button></form></p>
<p class="mute">내려받은 .eml 을 열면 아웃룩에 보낼 수 있는 새 메일로 열립니다. 보낸 뒤 [보냄으로 표시]를 누르면
{after}같은 메일이 다시 뜨지 않습니다.
</p>""")
        body = f"""<div class="card"><p><b>받는 사람</b> {esc(d.to) or '<span class=warn>회사 담당자 이메일이 없습니다</span>'}
{f' · <b>참조</b> {esc(d.cc)}' if d.cc else ''}</p><p><b>제목</b> {esc(d.subject)}</p><pre>{esc(d.body)}</pre>
<p class="mute">{attach}</p>{actions}</div><p><a href="/e/{e['id']}">← 감사 건으로</a></p>"""
        self._send(page(f"{d.label} 메일", body, *self._flash()))

    def mail_eml(self, eid, kind, on):
        e, cfg, items, d = self._draft(eid, kind, on)
        if not d.to:
            raise ValueError("회사 담당자 이메일이 없습니다. 회사 화면에서 넣어 주세요.")
        name = f"{e['client_name']}_{e['kind']}_자료요청목록.xlsx" if kind == "request" else None
        msg = mailer.build_message(d, cfg.get("sender_email", ""), items, name)
        self._download(mailer.eml_bytes(msg), f"{e['client_name']}_{d.label}.eml", "message/rfc822")

    def mail_sent(self, eid, kind, on):
        e, cfg, items, d = self._draft(eid, kind, on)
        if self.store.mail_sent(e["id"], kind, d.planned_on):
            raise ValueError("이미 보냄으로 표시한 메일입니다.")
        schedule.record_sent(self.store, d, self.app.today())
        self._redirect(f"/e/{eid}", f"{d.label} 메일을 보냄으로 기록했습니다.")

    def mail_send(self, eid, kind, on):
        e, cfg, items, d = self._draft(eid, kind, on)
        if self.store.mail_sent(e["id"], kind, d.planned_on):
            raise ValueError("이미 보낸 메일입니다.")
        if not d.to:
            raise ValueError("회사 담당자 이메일이 없습니다.")
        try:
            mailbox.send(cfg, mailer.message_for(self.store, cfg, e, d), mailbox.sender_for(cfg, e))
        except (OSError, mailbox.smtplib.SMTPException) as exc:
            raise ValueError(f"보내지 못했습니다: {exc}") from exc
        schedule.record_sent(self.store, d, self.app.today())
        self.app.note([f"보냄: {e['client_name']} {d.label} → {d.to}"])
        self._redirect(f"/e/{eid}", f"{d.label} 메일을 보냈습니다.")

    def link_done(self, eid, lid):
        self.store.set_link_done(int(lid))
        self._redirect(f"/e/{eid}", "링크를 처리함으로 표시했습니다.")

    def mail_fetch(self):
        cfg = self.app.cfg()
        try:
            r = mailbox.fetch(self.store, cfg, self.app.today())
        except (OSError, imaplib_error()) as exc:
            raise ValueError(f"메일을 확인하지 못했습니다: {exc}") from exc
        matched = sum(1 for f in r.saved if f["matched"])
        msg = (f"새 메일 {r.checked}통 확인: 첨부 {len(r.saved)}개 저장(항목 연결 {matched}개), "
               f"대용량 링크 {r.links}개, 감사와 관계없는 메일 {r.skipped}통")
        if r.errors:
            msg += " / 실패: " + "; ".join(r.errors)
        self.app.note([msg])
        self._redirect("/", msg, error=bool(r.errors) and not r.checked)

    def mail_account_save(self):
        f, _ = self._form()
        user = f.get("user", "").strip()
        if "@" not in user:
            raise ValueError("메일 주소를 넣어 주세요.")
        provider = f.get("provider", "직접 입력")
        preset = mailbox.PRESETS.get(provider, {})
        acc = {"user": user, "provider": provider}
        for k in ("imap_host", "smtp_host"):
            acc[k] = (f.get(k) or preset.get(k, "")).strip()
            if not acc[k]:
                raise ValueError("받기·보내기 서버 주소를 넣어 주세요.")
        for k in ("imap_port", "smtp_port"):
            acc[k] = int(f.get(k) or preset.get(k) or 0)
        cfg = self.app.cfg()
        accounts = [a for a in cfg["mail_accounts"] if a["user"] != user]
        old = next((i for i, a in enumerate(cfg["mail_accounts"]) if a["user"] == user), len(accounts))
        accounts.insert(old, acc)
        self.store.set_setting("mail_accounts", accounts)
        if not cfg["send_from"]:
            self.store.set_setting("send_from", user)
        note = f"{user} 계정을 저장했습니다."
        if f.get("password"):
            note += " " + mailbox.set_password(user, f["password"], remember=f.get("remember") == "1")
        self._redirect("/settings#mail", note)

    def mail_account_delete(self):
        f, _ = self._form()
        user = f.get("user", "")
        cfg = self.app.cfg()
        accounts = [a for a in cfg["mail_accounts"] if a["user"] != user]
        self.store.set_setting("mail_accounts", accounts)
        if cfg["send_from"] == user:
            self.store.set_setting("send_from", accounts[0]["user"] if accounts else "")
        mailbox._PASSWORD.pop(user, None)
        from . import credstore
        credstore.delete(user)
        self._redirect("/settings#mail", f"{user} 계정을 지웠습니다.")

    def mail_options_save(self):
        f, _ = self._form()
        new = {"send_from": f.get("send_from", ""), "auto_send": f.get("auto_send") == "1",
               "send_hour": int(f.get("send_hour") or 9), "check_minutes": int(f.get("check_minutes") or 10)}
        if not 0 <= new["send_hour"] <= 23 or new["check_minutes"] < 1:
            raise ValueError("발송 시각은 0~23시, 확인 주기는 1분 이상으로 넣어 주세요.")
        for k, v in new.items():
            self.store.set_setting(k, v)
        self._redirect("/settings#mail", "메일 설정을 저장했습니다.")

    def mail_test(self):
        f, _ = self._form()
        cfg = self.app.cfg()
        r = mailbox.test_connection(cfg, f.get("user") or None)
        if r.smtp_port:
            accs = cfg["mail_accounts"]
            for a in accs:
                if a["user"] == r.user and int(a["smtp_port"]) != r.smtp_port:
                    a["smtp_port"] = r.smtp_port    # 실제로 된 포트를 기억한다
                    self.store.set_setting("mail_accounts", accs)
        if r.ok:
            self._redirect("/settings#mail", r.summary())
        else:
            raise ValueError(r.summary())

    # ── 설정 ──────────────────────────────────────────
    def settings(self):
        cfg = self.app.cfg()
        try:
            pv = filing.preview(cfg["folder_rule"], cfg)
        except ValueError as exc:
            pv = f"오류: {exc}"
        presets = "".join(f"<option value='{esc(v)}'>{esc(k)}</option>" for k, v in config.RULE_PRESETS.items())

        def inp(key, label, **kw):
            attrs = " ".join(f'{k}="{esc(v)}"' for k, v in kw.items())
            return f'<label>{label}<input name="{key}" value="{esc(cfg[key])}" {attrs}></label>'

        def area(key, label, rows=6):
            return f'<label style="display:block;margin:8px 0">{label}<textarea name="{key}" rows="{rows}">{esc(cfg[key])}</textarea></label>'

        body = f"""<form method="post" action="/settings">
<div class="card"><h2 style="margin-top:0">보내는 사람</h2><div class="row">
{inp('firm_name', '회계법인')}{inp('author', '작성자')}{inp('contact', '연락처(전화·메일)')}{inp('sender_email', '보내는 메일 주소')}</div></div>
<div class="card"><h2 style="margin-top:0">일정</h2><div class="row">
{inp('request_offset', '최초 요청 (현장감사 D-n)', type='number', style='width:80px')}
{inp('reminder_from', '독촉 시작 (D-n부터 매일)', type='number', style='width:80px')}
{inp('due_offset', '제출기한 (D-n)', type='number', style='width:80px')}
<label style="flex-direction:row;gap:4px"><input type="checkbox" name="skip_weekends" value="1" {'checked' if cfg['skip_weekends'] else ''}>주말에는 보내지 않음</label></div>
<label style="display:block;margin-top:8px">공휴일 (YYYY-MM-DD, 줄마다 하나. 이날은 보내지 않음)
<textarea name="holidays" rows="3">{esc(chr(10).join(cfg['holidays']))}</textarea></label></div>
<div class="card"><h2 style="margin-top:0">저장 폴더 규칙</h2>
<div class="row">{inp('root', '루트 폴더', size=40)}
<label>미리 만든 규칙<select onchange="if(this.value)document.getElementById('rule').value=this.value">
<option value="">고르기…</option>{presets}</select></label></div>
<label style="display:block;margin:8px 0">규칙<input id="rule" name="folder_rule" value="{esc(cfg['folder_rule'])}" style="width:100%"></label>
<p class="mute">쓸 수 있는 칸: {esc(' '.join('{' + t + '}' for t in filing.TOKENS))}. 숫자 칸은 {{회사순번:02}}처럼 자릿수를 정합니다.</p>
<p>예시: <code>{esc(pv)}</code></p>
<label style="display:block">분류 폴더 (쉼표로 구분, 순서가 {{분류번호}})<input name="categories" value="{esc(', '.join(cfg['categories']))}" style="width:100%"></label>
<p class="mute">규칙을 바꾸면 이후 받는 파일부터 적용됩니다.</p></div>
<div class="card"><h2 style="margin-top:0">메일 문구</h2>
<p class="mute">중괄호 칸은 자동으로 채워집니다: {{회계법인}} {{작성자}} {{연락처}} {{담당자}} {{회사명}} {{사업연도}} {{감사구분}}
{{현장감사 시작일}} {{현장감사 종료일}} {{제출기한}} {{미제출 건수}} {{D-n}} {{감사코드}} {{요청 자료 표}} {{미제출 자료 표}}</p>
{inp('request_subject', '최초 요청 제목', style='width:100%')}{area('request_body', '최초 요청 본문', 12)}
{inp('reminder_subject', '독촉 제목', style='width:100%')}{area('reminder_body', '독촉 본문', 9)}{area('dday_body', '당일 본문', 7)}
</div><p><button>설정 저장</button></p></form>{self._mail_settings(cfg)}"""
        self._send(page("설정", body, *self._flash()))

    def _mail_settings(self, cfg) -> str:
        provs = [(k, k) for k in mailbox.PRESETS] + [("직접 입력", "직접 입력")]
        presets = json.dumps(mailbox.PRESETS, ensure_ascii=False)
        rows = []
        for a in cfg["mail_accounts"]:
            pw = "들어 있음" if mailbox.get_password(a["user"]) else "<span class='warn'>없음</span>"
            u = esc(a["user"])
            rows.append(
                f"<tr><td>{u}{' (기본)' if a['user'] == cfg['send_from'] else ''}</td><td>{esc(a['provider'])}</td>"
                f"<td class='mute'>{esc(a['imap_host'])}:{esc(a['imap_port'])} / {esc(a['smtp_host'])}:{esc(a['smtp_port'])}</td>"
                f"<td>{pw}</td><td class='row'>"
                f"<form method='post' action='/settings/mail-test'><input type='hidden' name='user' value='{u}'>"
                f"<button class='ghost'>연결 시험</button></form>"
                f"<button class='ghost' type='button' onclick='editAcc({esc(json.dumps(a, ensure_ascii=False))})'>고치기</button>"
                f"<form method='post' action='/settings/mail-delete' onsubmit='return confirm(\"이 계정을 지웁니다.\")'>"
                f"<input type='hidden' name='user' value='{u}'><button class='ghost'>지우기</button></form></td></tr>")
        table = ("<table><tr><th>계정</th><th>메일</th><th>서버 (받기 / 보내기)</th><th>비밀번호</th><th></th></tr>"
                 + "".join(rows) + "</table>") if rows else "<p class='mute'>아직 연결한 계정이 없습니다.</p>"
        senders = [(a["user"], a["user"]) for a in cfg["mail_accounts"]]
        return f"""<div class="card" id="mail"><h2 style="margin-top:0">메일 계정</h2>
<p class="mute">계정을 여러 개 넣을 수 있습니다. 받은 메일은 모든 계정에서 확인하고, 보낼 때는 기본 계정(또는 회사 화면에서
회사마다 정한 계정)을 씁니다. 대부분 일반 비밀번호 대신 메일 설정에서 만든 '앱 비밀번호'를 넣어야 하며,
비밀번호는 데이터베이스에 저장하지 않습니다.</p>{table}
<h2>계정 추가·고치기</h2><form method="post" action="/settings/mail"><div class="row">
<label>메일<select name="provider" id="prov">{options(provs, 'Gmail')}</select></label>
<label>메일 주소<input name="user" id="acc_user" type="email" required></label>
<label>비밀번호 (바꿀 때만)<input name="password" type="password" autocomplete="new-password"></label>
<label style="flex-direction:row;gap:4px"><input type="checkbox" name="remember" value="1">윈도우 자격 증명 관리자에 저장</label>
</div><div class="row" style="margin-top:8px">
<label>받기 서버<input name="imap_host" id="imap_host" value="imap.gmail.com"></label>
<label>포트<input name="imap_port" id="imap_port" value="993" size="4"></label>
<label>보내기 서버<input name="smtp_host" id="smtp_host" value="smtp.gmail.com"></label>
<label>포트<input name="smtp_port" id="smtp_port" value="465" size="4"></label>
<button>계정 저장</button></div></form>
<h2>보내기·확인 설정</h2><form method="post" action="/settings/mail-options"><div class="row">
<label>기본 보내는 계정<select name="send_from">{options(senders, cfg['send_from'])}</select></label>
<label>받은 메일 확인 (분마다)<input name="check_minutes" value="{esc(cfg['check_minutes'])}" size="3"></label>
<label style="flex-direction:row;gap:4px"><input type="checkbox" name="auto_send" value="1" {'checked' if cfg['auto_send'] else ''}>
요청·독촉 메일 자동 발송</label>
<label>발송 시각 (시)<input name="send_hour" value="{esc(cfg['send_hour'])}" size="3"></label>
<button class="ghost">저장</button></div></form>
<p class="mute">자동 발송을 켜면 매일 발송 시각 이후 '오늘 보낼 메일'을 회사 담당자에게 보냅니다.
처음에는 끈 채로 미리보기에서 [지금 보내기]로 몇 번 보내 보고 켜기를 권합니다. 프로그램이 켜져 있을 때만 동작합니다.</p></div>
<script>const P={presets};const $=id=>document.getElementById(id);
$('prov').onchange=function(){{const p=P[this.value];if(!p)return;for(const k in p)$(k).value=p[k];}};
function editAcc(a){{$('prov').value=a.provider;$('acc_user').value=a.user;
for(const k of ['imap_host','imap_port','smtp_host','smtp_port'])$(k).value=a[k];$('acc_user').scrollIntoView();}}</script>"""

    def settings_save(self):
        f, _ = self._form()
        cfg = self.app.cfg()
        new = {}
        for k in ("firm_name", "author", "contact", "sender_email", "root", "folder_rule", "request_subject",
                  "request_body", "reminder_subject", "reminder_body", "dday_body"):
            if k in f:
                new[k] = f[k].replace("\r\n", "\n").strip()
        for k in ("request_offset", "reminder_from", "due_offset"):
            if k in f:
                v = int(f[k])
                if not 0 <= v <= 60:
                    raise ValueError("일정 값은 0~60일 사이로 넣어 주세요.")
                new[k] = v
        new["skip_weekends"] = f.get("skip_weekends") == "1"
        hol = [h.strip() for h in f.get("holidays", "").splitlines() if h.strip()]
        for h in hol:
            date.fromisoformat(h)
        new["holidays"] = hol
        if "categories" in f:
            new["categories"] = [c.strip() for c in f["categories"].split(",") if c.strip()] or cfg["categories"]
        filing.preview(new.get("folder_rule", cfg["folder_rule"]), {**cfg, **new})  # 틀린 규칙이면 여기서 오류
        for k, v in new.items():
            self.store.set_setting(k, v)
        self._redirect("/settings", "설정을 저장했습니다.")


def imaplib_error():
    return mailbox.imaplib.IMAP4.error


def make_server(store: Store, host: str = "127.0.0.1", port: int = 8765, today=date.today) -> ThreadingHTTPServer:
    app = App(store, today)
    hosts = {f"{h}:{port}" for h in ("127.0.0.1", "localhost")}
    if host not in ("127.0.0.1", "localhost"):
        hosts = set()  # 팀 공유(0.0.0.0 등)일 때는 Host 검사를 끈다. Origin 검사는 그대로
    handler = type("BoundHandler", (Handler,), {"app": app, "allowed_hosts": hosts})
    return ThreadingHTTPServer((host, port), handler)


def serve(db_path: str, host: str = "127.0.0.1", port: int = 8765, open_browser: bool = True) -> None:
    store = Store(db_path)
    srv = make_server(store, host, port)
    url = f"http://{'127.0.0.1' if host in ('0.0.0.0', '') else host}:{port}"
    print(f"관리 화면: {url}  (끝내려면 Ctrl+C)")
    print(f"데이터베이스: {Path(db_path).resolve()}")
    if open_browser:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    start_background(srv)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        srv.server_close()


def start_background(srv) -> threading.Thread:
    """받은 메일 확인과 자동 발송을 주기적으로 돈다."""
    import time
    from datetime import datetime

    app = srv.RequestHandlerClass.app
    state: dict = {}

    def loop():
        while True:
            cfg = app.cfg()
            with app.lock:
                lines = mailbox.run_cycle(app.store, cfg, datetime.now(),
                                          lambda e, d: mailer.message_for(app.store, cfg, e, d), state)
            if lines:
                app.note(lines)
                for line in lines:
                    print(line)
            time.sleep(max(1, int(cfg.get("check_minutes") or 10)) * 60)

    t = threading.Thread(target=loop, daemon=True, name="mail")
    t.start()
    return t
