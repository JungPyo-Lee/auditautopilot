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

from . import auto_questions, config, confirmations, filing, mailbox, mailer, nicebiz, pbc, questions, schedule
from .db import KINDS, STATUSES, Store

MAX_UPLOAD = 300 * 1024 * 1024
STATUS_COLOR = {"미요청": "#8a8f98", "요청": "#2f6fdf", "일부수령": "#c07a00", "수령": "#1d8a4e",
                "검토완료": "#11613a", "보완요청": "#c2362b", "해당없음": "#b5b9c0",
                # 질문서
                "작성": "#8a8f98", "승인대기": "#c07a00", "승인": "#2f6fdf", "발송": "#6a4fc7",
                "답변받음": "#1d8a4e", "종결": "#11613a"}

CSS = """
:root{--bg:#f6f7f9;--card:#fff;--ink:#1d2330;--mute:#677083;--line:#e2e5ea;--accent:#2f6fdf;--warn:#c2362b}
*{box-sizing:border-box}body{margin:0;font:14px/1.5 -apple-system,"Malgun Gothic","맑은 고딕",sans-serif;
background:var(--bg);color:var(--ink)}a{color:var(--accent);text-decoration:none}a:hover{text-decoration:underline}
header{background:#1d2330;color:#fff;padding:10px 20px;display:flex;gap:20px;align-items:center}
header a{color:#cfd6e4}header b{color:#fff;margin-right:12px}main{max-width:1200px;margin:0 auto;padding:20px}
h1{font-size:20px;margin:4px 0 16px}h2{font-size:16px;margin:24px 0 8px}
.card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:16px;margin-bottom:16px}
.card.alert{border:2px solid var(--warn);box-shadow:0 0 0 3px rgba(194,54,43,.12)}.badge-bad{display:inline-block;border:2px solid var(--warn);color:var(--warn);border-radius:6px;padding:0 6px;font-weight:600}
.badge-ok{color:#1d7a3a}
table{border-collapse:collapse;width:100%}th,td{text-align:left;padding:6px 8px;border-bottom:1px solid var(--line);
vertical-align:top}th{font-weight:600;color:var(--mute);font-size:12px;background:#fafbfc}
input,select,textarea{font:inherit;padding:5px 7px;border:1px solid #c9ced6;border-radius:5px;background:#fff}
textarea{width:100%}button,.btn{font:inherit;padding:6px 12px;border:1px solid var(--accent);background:var(--accent);
color:#fff;border-radius:5px;cursor:pointer;display:inline-block}button.ghost,.btn.ghost{background:#fff;color:var(--accent)}
.pill{display:inline-block;padding:1px 8px;border-radius:10px;color:#fff;font-size:12px;white-space:nowrap}
.mute{color:var(--mute)}.warn{color:var(--warn)}.row{display:flex;gap:12px;flex-wrap:wrap;align-items:end}
.row label{display:flex;flex-direction:column;gap:3px;font-size:12px;color:var(--mute)}
.msg{padding:10px 14px;border-radius:6px;background:#e8f1ff;margin-bottom:16px;position:sticky;top:0;z-index:5;box-shadow:0 2px 6px rgba(0,0,0,.12)}.msg.err{background:#fdecea}
pre{white-space:pre-wrap;font-family:inherit;background:#fafbfc;border:1px solid var(--line);padding:10px;border-radius:6px;margin:6px 0}
.bar{height:8px;background:#e9ecf1;border-radius:4px;overflow:hidden;min-width:120px}
.bar i{display:block;height:100%;background:#1d8a4e}
button,.btn{white-space:nowrap}.card{overflow-x:auto}
.co{font-size:22px;margin:4px 0 2px}.steps{display:flex;gap:3px;margin:12px 0 4px}
.steps a{flex:1;color:#fff;font-size:12px;text-align:center;padding:6px 4px;border-radius:4px;text-decoration:none;
white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.steps a.ok{background:#1d8a4e}.steps a.no{background:#c2362b}
.ministeps{display:flex;gap:2px}.ministeps span{width:18px;height:12px;border-radius:2px;display:inline-block}
.ministeps .ok{background:#1d8a4e}.ministeps .no{background:#c2362b}
.tabs{display:flex;gap:2px;border-bottom:2px solid #9aa3b2;margin-top:14px;flex-wrap:wrap}
.tabs a{padding:7px 14px;background:#e4e7ec;border:1px solid #c9ced6;border-bottom:none;border-radius:6px 6px 0 0;
color:var(--ink);font-size:13px;text-decoration:none}.tabs a.on{background:#fff;font-weight:600;border-color:#9aa3b2;
position:relative;top:2px;border-bottom:2px solid #fff}.tabs a i{font-style:normal;margin-right:4px}
.js .tab{display:none}.js .tab.on{display:block}.tab{padding-top:14px}td input:not([type=checkbox]):not([type=hidden]){width:100%;min-width:56px}
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


def xlsx_sheet_names(data: bytes) -> list[str]:
    """엑셀 파일을 디스크에 쓰지 않고 시트 이름만 읽는다."""
    import io
    import zipfile

    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            xml = z.read("xl/workbook.xml").decode("utf-8", "replace")
    except (zipfile.BadZipFile, KeyError):
        return []
    return [html.unescape(n) for n in re.findall(r'<(?:\w+:)?sheet\b[^>]*\bname="([^"]*)"', xml)]


def stamped(filename: str, now=None) -> str:
    """내려받는 엑셀·워드 파일 이름 끝에 만든 시각을 붙인다.
    엑셀은 이름이 같은 파일을 동시에 열지 못하고, 받은 폴더에 같은 이름이 있으면 덮어쓰기에서 막히기 때문."""
    from datetime import datetime

    p = Path(filename)
    if p.suffix.lower() not in (".xlsx", ".docx"):
        return filename
    return f"{p.stem}_{(now or datetime.now()).strftime('%Y%m%d_%H%M%S')}{p.suffix}"

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
        filename = stamped(filename)
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
                        lists.setdefault(f"@{name}", []).append(len(files) - 1)   # 파일 칸 이름 -> 순번
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
        ("POST", r"/e/(\d+)/confirm", "confirm_reconcile"),
        ("POST", r"/e/(\d+)/confirm/fetch", "confirm_fetch"),
        ("POST", r"/e/(\d+)/confirm/project", "confirm_project"),
        ("POST", r"/e/(\d+)/confirm/materiality", "confirm_materiality"),
        ("GET", r"/e/(\d+)/confirm\.xlsx", "confirm_xlsx"),
        ("POST", r"/e/(\d+)/leadsheet", "leadsheet_make"),
        ("GET", r"/e/(\d+)/leadsheet\.xlsx", "leadsheet_xlsx"),
        ("GET", r"/e/(\d+)/aje-template\.xlsx", "aje_template"),
        ("POST", r"/e/(\d+)/adjust", "adjust_make"),
        ("GET", r"/e/(\d+)/adjusted\.xlsx", "adjusted_xlsx"),
        ("GET", r"/e/(\d+)/notes-request\.xlsx", "notes_request"),
        ("GET", r"/e/(\d+)/report\.docx", "report_docx"),
        ("POST", r"/e/(\d+)/notes-filled", "notes_filled"),
        ("GET", r"/e/(\d+)/fs\.docx", "fs_docx"),
        ("POST", r"/e/(\d+)/questions", "question_add"),
        ("POST", r"/e/(\d+)/questions/act", "question_act"),
        ("POST", r"/e/(\d+)/questions/(\d+)", "question_edit"),
        ("POST", r"/e/(\d+)/questions/from-confirm", "question_from_confirm"),
        ("POST", r"/e/(\d+)/questions/auto", "question_auto"),
        ("POST", r"/e/(\d+)/questions/answers", "question_answers"),
        ("GET", r"/e/(\d+)/questions\.xlsx", "question_xlsx"),
        ("GET", r"/e/(\d+)/questions/mail", "question_mail"),
        ("GET", r"/e/(\d+)/questions/mail\.eml", "question_eml"),
        ("POST", r"/e/(\d+)/questions/mail/send", "question_send"),
        ("POST", r"/e/(\d+)/questions/mail/sent", "question_mark_sent"),
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
        ("POST", r"/settings/nice", "nice_save"),
        ("POST", r"/settings/nice-test", "nice_test"),
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
                except (ValueError, KeyError, OSError) as exc:
                    if isinstance(exc, OSError) and not isinstance(exc, FileNotFoundError):
                        exc = ValueError(f"파일을 읽거나 쓰지 못했습니다 ({exc.strerror or exc}). "
                                         "같은 파일이 엑셀에서 열려 있으면 닫고 다시 해 주세요.")
                    back = self.headers.get("Referer") or "/"
                    if method == "POST":
                        to = urllib.parse.urlsplit(back).path or "/"
                        sec = re.match(r"/e/\d+/(questions|confirm|leadsheet|adjust|notes-filled)", self.path)
                        if sec and to.startswith("/e/"):        # 오류가 난 카드로 돌아간다
                            to += "#" + {"notes-filled": "closing"}.get(sec.group(1), sec.group(1))
                        self._redirect(to, str(exc), error=True)
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

    # ── 진행 단계 ─────────────────────────────────────
    STEPS = (("mail", "메일 일정"), ("pbc", "자료요청"), ("files", "받은자료"), ("questions", "질문서"),
             ("confirm", "채권채무조회"), ("lead", "리드시트"), ("close", "수정분개·재무제표"))

    def progress(self, e, cfg) -> list[tuple[str, str, bool, str]]:
        """[(탭, 이름, 완료, 설명)] 업무 순서대로."""
        eid = e["id"]
        items = self.store.items(eid)
        req = [i for i in items if i["status"] != "해당없음"]
        out_items = [i for i in req if i["status"] in ("미요청", "요청", "일부수령", "보완요청")]
        files = self.store.files(eid)
        unfiled = [f for f in files if not f["item_id"]]
        qs = questions.all_(self.store, eid)
        open_q = [q for q in qs if q["status"] != "종결"]
        a = confirmations.assess(confirmations.latest(self.store, eid), confirmations.materiality(self.store, eid),
                                 float(cfg.get("confirm_min_reply_rate", 80)))
        sent = self.store.request_sent(eid)
        lead = self.store.get_setting(f"leadsheet:{eid}")
        adj = self.store.get_setting(f"adjust:{eid}")
        nf = self.store.get_setting(f"notes_filled:{eid}")
        steps = [
            (sent, "자료요청 메일 보냄" if sent else "자료요청 메일을 아직 안 보냄"),
            (bool(req) and not out_items, f"미제출 {len(out_items)}건" if out_items else
             ("요청 항목 없음" if not req else "모두 받음")),
            (bool(files) and not unfiled, "받은 파일 없음" if not files else
             (f"항목 미연결 {len(unfiled)}개" if unfiled else f"{len(files)}개 모두 항목에 연결")),
            (bool(qs) and not open_q, "질문서 없음" if not qs else
             (f"종결 안 된 질문 {len(open_q)}건" if open_q else "모두 종결")),
            (bool(a and a["ok"]), "회신현황 없음" if a is None else
             ("기준 충족" if a["ok"] else " / ".join(a["reasons"]))),
            (bool(lead), "조서 채움" if lead else "조서를 아직 안 채움"),
            (bool(adj and nf), "수정분개 반영·주석 자료 받음" if adj and nf else
             ("주석 자료를 아직 안 받음" if adj else "수정분개를 아직 반영 안 함")),
        ]
        return [(k, lbl, ok, hint) for (k, lbl), (ok, hint) in zip(self.STEPS, steps)]

    def _tabs(self, e, cfg, panels: dict[str, str]) -> str:
        prog = self.progress(e, cfg)
        done = sum(1 for p in prog if p[2])
        bar = "".join(f"<a href='#tab-{k}' class='{'ok' if ok else 'no'}' title='{esc(lbl)}: {esc(hint)}'>"
                      f"{'✓' if ok else '✗'} {esc(lbl)}</a>" for k, lbl, ok, hint in prog)
        tabs = "".join(f"<a href='#tab-{k}' data-tab='{k}'><i style='color:{'#1d8a4e' if ok else '#c2362b'}'>●</i>"
                       f"{esc(lbl)}</a>" for k, lbl, ok, _ in prog)
        body = "".join(f"<section class='tab' id='tab-{k}'>{panels[k]}</section>" for k, *_ in prog)
        return f"""<div class="steps">{bar}</div><p class="mute" style="margin:0">진행 {done}/{len(prog)} 단계 완료 ·
막대에 마우스를 올리면 남은 일이 보입니다.</p>
<nav class="tabs">{tabs}</nav>{body}
<script>(function(){{document.body.classList.add('js');
var tabs=document.querySelectorAll('.tabs a'),secs=document.querySelectorAll('section.tab');
function show(k){{secs.forEach(function(s){{s.classList.toggle('on',s.id==='tab-'+k)}});
tabs.forEach(function(a){{a.classList.toggle('on',a.dataset.tab===k)}});try{{localStorage.setItem('tab{e['id']}',k)}}catch(x){{}}}}
function fromHash(){{var h=location.hash.slice(1);if(!h)return null;if(h.indexOf('tab-')===0)return h.slice(4);
var el=document.getElementById(h);var s=el&&el.closest('section.tab');return s?s.id.slice(4):null}}
var k=fromHash();if(!k){{try{{k=localStorage.getItem('tab{e['id']}')}}catch(x){{}}}}show(k||'mail');
var el=location.hash&&document.getElementById(location.hash.slice(1));if(el&&el.tagName!=='SECTION')setTimeout(function(){{el.scrollIntoView()}},0);
window.addEventListener('hashchange',function(){{var k=fromHash();if(k)show(k)}});
document.querySelectorAll('.steps a,.tabs a').forEach(function(a){{a.addEventListener('click',function(ev){{
ev.preventDefault();var k=a.getAttribute('href').slice(5);show(k);history.replaceState(null,'','#tab-'+k)}})}});}})();</script>"""

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
            a = confirmations.assess(confirmations.latest(self.store, e["id"]),
                                     confirmations.materiality(self.store, e["id"]),
                                     float(cfg.get("confirm_min_reply_rate", 80)))
            conf = ("-" if a is None else (f"<span class='badge-ok'>회신 {a['rate']}%</span>" if a["ok"]
                    else f"<a href='/e/{e['id']}#confirm' class='badge-bad' title='{esc(' / '.join(a['reasons']))}'>회신 {a['rate']}%</a>"))
            done = e["total"] - e["outstanding"]
            pct = int(done / e["total"] * 100) if e["total"] else 0
            prog = self.progress(e, cfg)
            mini = ("<div class='ministeps'>" + "".join(
                f"<a href='/e/{e['id']}#tab-{k}' title='{esc(lbl)}: {esc(hint)}'><span class='{'ok' if ok else 'no'}'></span></a>"
                for k, lbl, ok, hint in prog) + f"</div><span class='mute'>{sum(p[2] for p in prog)}/{len(prog)}</span>")
            erows.append(
                f"<tr><td><a href='/e/{e['id']}'>{esc(e['client_code'])} {esc(e['client_name'])}</a></td>"
                f"<td>{esc(e['fiscal_year'])} {esc(e['kind'])}</td>"
                f"<td>{esc(e['fieldwork_start'])} ~ {esc(e['fieldwork_end'])}</td><td>{mini}</td>"
                f"<td><div class='bar'><i style='width:{pct}%'></i></div><span class='mute'>{done}/{e['total']}</span></td>"
                f"<td>{e['outstanding']}건</td><td>{conf}</td><td>{esc(f'{nxt.on} {nxt.label}') if nxt else '-'}</td></tr>")
        eng_html = ("<table><tr><th>회사</th><th>감사</th><th>현장감사</th><th title=\"" + " → ".join(l for _, l in self.STEPS) + "\">업무 진행 (" + str(len(self.STEPS)) + "단계)</th><th>수령</th><th>미제출</th><th>채권채무 조회</th><th>다음 메일</th></tr>"
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
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as d:
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
        lead_rec = self.store.get_setting(f"leadsheet:{e['id']}")
        panels = {
            "mail": f"""<div class="card"><h2 style="margin-top:0">메일 일정</h2>
<table><tr><th>날짜</th><th>메일</th><th>상태</th><th></th></tr>{''.join(steps)}</table>
<form method="post" action="/e/{e['id']}/dates" class="row" style="margin-top:10px">
<label>현장감사 시작일<input type="date" name="fieldwork_start" value="{esc(e['fieldwork_start'])}"></label>
<label>종료일<input type="date" name="fieldwork_end" value="{esc(e['fieldwork_end'])}"></label>
<button class="ghost">일정 바꾸기</button></form></div>""",
            "pbc": f"""<div class="card"><h2 style="margin-top:0">자료요청 목록</h2><p>{summary}
<a class="btn ghost" href="/e/{e['id']}/pbc.xlsx">엑셀로 받기</a></p>
<form method="post" action="/e/{e['id']}/items"><table>
<tr><th>요청번호</th><th>구분</th><th>자료명</th><th>상태</th><th>요청일</th><th>수령일</th><th>비고(메일에 함께 나감)</th></tr>
{''.join(irows)}</table><p><button>상태 저장</button></p></form>
<form method="post" action="/e/{e['id']}/items/add" class="row">
<label>구분<select name="section">{options(sec_opts)}</select></label>
<label>자료명<input name="title" required size="40"></label><button class="ghost">항목 추가</button></form></div>""",
            "files": f"""<div class="card"><h2 style="margin-top:0">받은 자료</h2>
<form method="post" action="/e/{e['id']}/upload" enctype="multipart/form-data" class="row">
<input type="file" name="files" multiple><button>올려서 정리</button></form>
<p class="mute">메일로 받은 첨부나 대용량 링크로 받은 파일을 여기에 올리거나, 아래 수신함 폴더에 넣고
[수신함 정리]를 누르면 폴더 규칙대로 저장됩니다. 파일 이름에 요청번호(예: 중간-03)나 자료명이 있으면 해당 항목이 '수령'으로 바뀝니다.<br>
수신함: <code>{esc(inbox)}</code></p>
<form method="post" action="/e/{e['id']}/sweep"><button class="ghost">수신함 정리</button></form>
{links_html}
<table style="margin-top:10px"><tr><th>받은 날</th><th>원본 파일명</th><th>저장 위치</th><th>요청 항목</th></tr>
{''.join(frows) or '<tr><td colspan=4 class=mute>아직 없습니다.</td></tr>'}</table></div>""",
            "questions": self._question_card(e, cfg),
            "confirm": self._confirm_card(e, cfg),
            "lead": self._leadsheet_card(e),
            "close": f"<div class='card'>{self._adjust_part(e, lead_rec)}</div>",
        }
        body = f"""<p class="mute">{esc(e['client_code'])} · 현장감사 {esc(e['fieldwork_start'])} ~ {esc(e['fieldwork_end'])}
· 담당자 {esc(e['contact_name'] or '-')} &lt;{esc(e['contact_email'] or '이메일 없음')}&gt;</p>
{self._tabs(e, cfg, panels)}
<form method="post" action="/e/{e['id']}/delete" onsubmit="return confirm('이 감사 건과 요청 목록을 지웁니다. 저장된 파일은 지우지 않습니다.')">
<button class="ghost" style="color:var(--warn);border-color:var(--warn)">감사 건 삭제</button></form>"""
        self._send(page(f"{e['client_name']} {e['fiscal_year']} {e['kind']}", body, *self._flash()))

    def engagement_dates(self, eid):
        f, _ = self._form()
        if f.get("fieldwork_end", "") < f.get("fieldwork_start", ""):
            raise ValueError("종료일이 시작일보다 빠릅니다.")
        self.store.update_engagement(int(eid), fieldwork_start=f["fieldwork_start"], fieldwork_end=f["fieldwork_end"])
        self._redirect(f"/e/{eid}#tab-mail", "일정을 바꿨습니다.")

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
        self._redirect(f"/e/{eid}#tab-pbc", f"{n}개 항목을 저장했습니다.")

    def item_add(self, eid):
        f, _ = self._form()
        if not f.get("title", "").strip():
            raise ValueError("자료명을 넣어 주세요.")
        self.store.add_item(int(eid), f.get("section", "기말감사"), f["title"])
        self._redirect(f"/e/{eid}#tab-pbc", "항목을 추가했습니다.")

    def upload(self, eid):
        e, cfg = self._eng(eid), self.app.cfg()
        _, files = self._form()
        if not files:
            raise ValueError("올릴 파일을 고르세요.")
        res = [filing.save_file(self.store, e, cfg, name, data=data, received_on=self.app.today(), source="업로드")
               for name, data in files]
        matched = sum(r["matched"] for r in res)
        self._redirect(f"/e/{eid}#tab-files", f"{len(res)}개 저장 (항목 연결 {matched}개, 미분류 {len(res) - matched}개)")

    def sweep(self, eid):
        e = self._eng(eid)
        res = filing.sweep_inbox(self.store, e, self.app.cfg(), self.app.today())
        matched = sum(r["matched"] for r in res)
        self._redirect(f"/e/{eid}#tab-files", f"수신함에서 {len(res)}개 정리 (항목 연결 {matched}개)")

    def refile(self, eid, fid):
        f, _ = self._form()
        item_id = int(f["item_id"]) if f.get("item_id") else None
        dest = filing.refile(self.store, self._eng(eid), self.app.cfg(), int(fid), item_id)
        self._redirect(f"/e/{eid}#tab-files", f"옮김: {dest}")

    def _confirm_card(self, e, cfg) -> str:
        eid = e["id"]
        rec = confirmations.latest(self.store, eid)
        projects = confirmations.cached_projects(self.store)
        cur = confirmations.project_of(self.store, eid)
        if confirmations.ready(cfg):
            opts = "<option value=''>자동으로 찾기 (회사 이름·연도)</option>" + "".join(
                f"<option value='{esc(p['id'])}' {'selected' if p['id'] == cur else ''}>"
                f"{esc(p['label'])}{'' if p.get('in_progress', True) else ' (종료)'}</option>" for p in projects)
            top = (f"<form method='post' action='/e/{eid}/confirm/project' class='row'>"
                   f"<label>전자조회 프로젝트<select name='project_id' style='max-width:520px'>{opts}</select></label>"
                   f"<button class='ghost'>연결</button></form>"
                   f"<form method='post' action='/e/{eid}/confirm/fetch' style='margin-top:8px'>"
                   f"<button>전자조회에서 지금 가져오기</button></form>")
        else:
            top = ("<p>설정 &gt; <a href='/settings#nice'>전자조회 계정</a>에 아이디·비밀번호를 넣으면 "
                   "사이트에서 회신현황을 바로 가져옵니다.</p>")
        mat = confirmations.materiality(self.store, eid)
        a = confirmations.assess(rec, mat, float(cfg.get("confirm_min_reply_rate", 80)))
        mat_form = (f"<form method='post' action='/e/{eid}/confirm/materiality' class='row' style='margin-top:8px'>"
                    f"<label>중요성 금액 (미회신 금액 판정 기준)<input name='materiality' value='{'' if mat is None else f'{mat:,.0f}'}' "
                    f"placeholder='예: 354,929,000' style='width:160px'></label><button class='ghost'>저장</button>"
                    f"<span class='mute'>회신률 기준 {float(cfg.get('confirm_min_reply_rate', 80)):g}% (설정에서 바꿈)</span></form>")
        verdict = ""
        if a:
            if a["ok"]:
                verdict = (f"<p class='badge-ok'>기준 충족: 회신률 {a['rate']}% ({a['replied']}/{a['total']}건), "
                           f"미회신 금액 {a['noreply_amount']:,.0f}원</p>")
            else:
                verdict = "<p><span class='badge-bad'>확인 필요</span> " + " / ".join(esc(r) for r in a["reasons"]) + "</p>"
            if a["notes"]:
                verdict += "<p class='mute'>" + " · ".join(esc(n) for n in a["notes"]) + "</p>"
        result = "<p class='mute'>아직 가져온 결과가 없습니다.</p>"
        if rec:
            c = rec.get("counts", {})
            rows = "".join(
                f"<tr><td>{esc(s['계정'])}</td><td>{s['조회 건수']}</td><td>{s['일치']}</td><td>{s['불일치']}</td>"
                f"<td>{s['미회신']}</td><td style='text-align:right'>{s['발송금액(원화)']:,.0f}</td></tr>"
                for s in rec.get("summary", []))
            result = (f"<p>{esc(rec.get('fetched_at', ''))} {esc(rec.get('source', ''))}"
                      f"{' · ' + esc(rec['project']) if rec.get('project') else ''}<br>"
                      f"일치 {c.get('일치', 0)} · <b>불일치 {c.get('불일치', 0)}</b> · <b>미회신 {c.get('미회신', 0)}</b> "
                      f"<a class='btn ghost' href='/e/{eid}/confirm.xlsx'>대사표 내려받기</a></p>"
                      f"<table><tr><th>계정</th><th>건수</th><th>일치</th><th>불일치</th><th>미회신</th><th>발송금액(원화)</th></tr>"
                      f"{rows}</table>")
        alert = " alert" if a and not a["ok"] else ""
        return f"""<div class="card{alert}" id="confirm"><h2 style="margin-top:0">채권채무 조회 (전자조회)</h2>
{top}{mat_form}
<div style="margin-top:10px">{verdict}{result}</div>
<form method="post" action="/e/{eid}/confirm" enctype="multipart/form-data" class="row" style="margin-top:10px">
<input type="file" name="files" multiple accept=".xlsx"><button class="ghost">장부와 함께 대사표 만들기</button></form>
<p class="mute">더존 재무상태표·거래처원장을 고르면 위에서 가져온 회신현황과 합쳐 발송금액을 장부 잔액과 맞춰 보고 계정별 조회 비율을 낸 대사표를 내려받습니다.
사이트에서 직접 받은 회신현황 엑셀을 함께 골라도 됩니다.</p></div>"""

    # ── 리드시트 ──────────────────────────────────────
    def _leadsheet_card(self, e) -> str:
        eid = e["id"]
        rec = self.store.get_setting(f"leadsheet:{eid}")
        last = ""
        if rec:
            lines = "".join(f"<li>{esc(x)}</li>" for x in rec.get("lines", []))
            last = (f"<p>{esc(rec.get('made_at', ''))} 만든 조서: <a class='btn ghost' href='/e/{eid}/leadsheet.xlsx'>"
                    f"{esc(Path(rec.get('path', '')).name)} 내려받기</a></p><ul class='mute'>{lines}</ul>")
        return f"""<div class="card" id="leadsheet"><h2 style="margin-top:0">리드시트·거래처 명세 (조서 채우기)</h2>
<form method="post" action="/e/{eid}/leadsheet" enctype="multipart/form-data">
<div class="row"><label>조서 양식<input type="file" name="workpaper" accept=".xlsx" required></label>
<label>회사 자료 (여러 개)<input type="file" name="files" multiple accept=".xlsx"></label>
<button>조서 채우기</button></div></form>
<p class="mute">더존 합계잔액시산표는 꼭 넣고, 재무상태표·손익계산서(과목별, 전기 금액용)와 거래처원장(잔액)을 함께 고르면
TB·BS·IS 탭이 채워져 ○1000 리드시트가 저절로 계산되고, 거래처원장은 해당 ○1100 탭에 거래처별 명세로 들어갑니다.
거래처원장이 어느 계정인지는 잔액으로 찾고, 못 찾으면 파일 이름을 계정명으로 바꿔 올려 주세요 (예: 미지급금.xlsx).
회사가 준 시산표를 조서에 옮겨 적는 것이며 재무제표를 대신 작성하지 않습니다.</p>{last}</div>"""

    def leadsheet_make(self, eid):
        import tempfile
        from datetime import datetime

        from ..workpaper import leadsheet

        e, cfg = self._eng(eid), self.app.cfg()
        f, files = self._form()
        idx = (f.get("@workpaper") or [None])[0]
        if idx is None:
            raise ValueError("조서 양식 파일을 고르세요.")
        wp = files[idx]
        sources = [x for i, x in enumerate(files) if i != idx]
        if not sources:
            raise ValueError("합계잔액시산표 등 회사 자료를 함께 고르세요.")
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            paths = []
            for i, (name, data) in enumerate([wp] + sources):
                if not name.lower().endswith(".xlsx"):
                    raise ValueError(f"{name}: .xlsx 파일만 받습니다.")
                d = Path(tmp) / f"{i:02d}"
                d.mkdir()
                p = d / Path(name).name
                p.write_bytes(data)
                paths.append(p)
            out = Path(tmp) / "out" / f"{Path(wp[0]).stem}_리드시트.xlsx"
            res = leadsheet.build(paths[0], paths[1:], out)
            saved = filing.save_file(self.store, e, cfg, out.name, data=out.read_bytes(), received_on=self.app.today(),
                                     source="리드시트", category="조서")
        self.store.set_setting(f"leadsheet:{e['id']}", {
            "made_at": datetime.now().replace(microsecond=0).isoformat(" "), "path": str(saved["path"]),
            "lines": res.lines()})
        self._redirect(f"/e/{eid}#leadsheet", res.lines()[0])

    def _adjust_part(self, e, lead_rec) -> str:
        eid = e["id"]
        rec = self.store.get_setting(f"adjust:{eid}")
        last = ""
        if rec:
            lines = "".join(f"<li>{esc(x)}</li>" for x in rec.get("lines", []))
            last = (f"<p>{esc(rec.get('made_at', ''))} 반영: <a class='btn ghost' href='/e/{eid}/adjusted.xlsx'>"
                    f"{esc(Path(rec.get('path', '')).name)} 내려받기</a></p><ul class='mute'>{lines}</ul>")
        tpl_btn = (f"<a class='btn ghost' href='/e/{eid}/aje-template.xlsx'>수정분개 양식 내려받기</a>" if lead_rec
                   else "<span class='mute'>조서를 채우면 그 계정과목으로 수정분개 양식을 받을 수 있습니다.</span>")
        return f"""<h3 id="adjust">수정분개 반영 (정산표·수정후 BS·IS·현금흐름정산표)</h3>
<p>{tpl_btn}</p>
<form method="post" action="/e/{eid}/adjust" enctype="multipart/form-data" class="row">
<label>채운 수정분개 양식<input type="file" name="aje" accept=".xlsx" required></label>
<label>다른 조서에 반영할 때만: 조서 파일 (비우면 마지막으로 채운 조서)<input type="file" name="workpaper" accept=".xlsx"></label>
<button>수정분개 반영</button></form>
<p class="mute">조서에 AJE·정산표·수정후BS·수정후IS·현금흐름정산표 탭을 더하고, 리드시트 Audited(F/S) 칸을 'G/L + 수정분개' 수식으로 바꿉니다.
모든 금액이 수식으로 이어져 있고 탭마다 검증 칸(0이어야 함)이 있습니다. 다시 올리면 그 탭들을 새로 만듭니다.</p>{last}
{self._closing_part(e, rec)}"""

    def _closing_part(self, e, adj_rec) -> str:
        eid = e["id"]
        title = '<h3 id="closing">주석 자료 · 재무제표·주석 초안 (워드)</h3>'
        if not adj_rec:
            return title + """<p class="mute">위에서 [수정분개 반영]을 한 번 하면 여기에 주석 자료요청 양식과 재무제표·주석 초안 버튼이 생깁니다.
수정분개가 없으면 모든 칸이 비어 있는 수정분개 양식을 그대로 올려도 됩니다.</p>"""
        nrec = self.store.get_setting(f"notes_filled:{eid}") or {}
        got = (f"<p>{esc(nrec.get('made_at', ''))} 받은 주석 자료: <b>{esc(nrec.get('name', ''))}</b> "
               f"(초안에 이 파일을 씁니다)</p>" if nrec else
               "<p class='mute'>아직 받은 주석 자료가 없습니다. 지금 초안을 만들면 정산표 금액만 들어가고 회사 작성 주석은 [ ]로 비어 있습니다.</p>")
        return title + f"""
<p class="row"><a class="btn ghost" href="/e/{eid}/notes-request.xlsx">① 주석 자료요청 양식 내려받기</a></p>
<p class="mute">수정분개까지 반영한 금액을 미리 넣은 주석 시트(회색 칸)와 회사가 처음부터 채울 시트(특수관계자·담보·보험·토지 공시지가·약정·소송 등)로 나뉩니다.
계정이 없는 주석은 빠집니다. 회사에 보내 채워 받으세요.</p>
<form method="post" action="/e/{eid}/notes-filled" enctype="multipart/form-data" class="row">
<label>② 회사가 채운 주석 자료<input type="file" name="notes" accept=".xlsx" required></label>
<button class="ghost">올리기</button></form>
{got}
<p class="row"><a class="btn" href="/e/{eid}/fs.docx">③ 재무제표·주석 초안 (워드)</a></p>
<p class="mute">재무상태표·손익계산서·자본변동표·현금흐름표(정산표 금액)와 받은 주석 자료의 시트마다 주석 하나를 넣은 워드 파일입니다.
표는 워드 표라 그대로 복사해 붙여 넣을 수 있습니다. 회사 개요·회계정책처럼 자료에 없는 주석은 [ ]로 비워 둡니다.</p>"""

    def _adjusted_workpaper(self, e) -> Path:
        rec = self.store.get_setting(f"adjust:{e['id']}") or {}
        p = Path(rec.get("path") or "")
        if not rec or not p.exists():
            raise ValueError("먼저 [수정분개 반영]을 해 주세요.")
        return p

    def notes_request(self, eid):
        import tempfile

        from ..workpaper import notes

        e = self._eng(eid)
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            res = notes.request_workbook(self._adjusted_workpaper(e), Path(tmp) / "n.xlsx", e["client_name"])
            data = res.output.read_bytes()
        self._download(data, f"{e['client_name']}_주석자료요청.xlsx", mailer.XLSX[0] + "/" + mailer.XLSX[1])

    def report_docx(self, eid):
        import tempfile

        from ..workpaper import report

        e = self._eng(eid)
        opinion = (self._query().get("opinion") or "적정").strip()
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            out, _ = report.draft(self._adjusted_workpaper(e), Path(tmp) / "r.docx", opinion)
            data = out.read_bytes()
        self._download(data, f"{e['client_name']}_감사보고서_초안.docx",
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    def notes_filled(self, eid):
        from datetime import datetime

        from ..workpaper import fs_doc

        e, cfg = self._eng(eid), self.app.cfg()
        _, files = self._form()
        got = [(n, d) for n, d in files if d]
        if not got:
            raise ValueError("회사가 채운 주석 자료 파일을 골라 주세요.")
        name, data = got[0]
        if not name.lower().endswith(".xlsx"):
            raise ValueError(f"{name}: .xlsx 파일만 받습니다.")
        if "목차" not in xlsx_sheet_names(data):
            raise ValueError(f"{name}: 주석 자료요청 양식(목차 시트가 있는 파일)이 아닙니다.")
        saved = filing.save_file(self.store, e, cfg, name, data=data, received_on=self.app.today(),
                                 source="주석 자료", category="조서")
        items = fs_doc.read_filled_notes(saved["path"])
        self.store.set_setting(f"notes_filled:{e['id']}", {
            "made_at": datetime.now().replace(microsecond=0).isoformat(" "), "path": str(saved["path"]),
            "name": name})
        self._redirect(f"/e/{eid}#closing", f"주석 자료를 받았습니다: 주석 {len(items)}개")

    def fs_docx(self, eid):
        import tempfile

        from ..workpaper import fs_doc

        e = self._eng(eid)
        nrec = self.store.get_setting(f"notes_filled:{e['id']}") or {}
        npath = Path(nrec["path"]) if nrec.get("path") and Path(nrec["path"]).exists() else None
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            res = fs_doc.draft(self._adjusted_workpaper(e), Path(tmp) / "fs.docx", npath)
            data = res.output.read_bytes()
        self._download(data, f"{e['client_name']}_재무제표_주석_초안.docx",
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    def _last_workpaper(self, e) -> Path:
        rec = self.store.get_setting(f"leadsheet:{e['id']}") or {}
        p = Path(rec.get("path") or "")
        if not rec or not p.exists():
            raise ValueError("먼저 [조서 채우기]를 해 주세요.")
        return p

    def aje_template(self, eid):
        import tempfile

        from ..workpaper import adjust

        e = self._eng(eid)
        wp = self._last_workpaper(e)
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            out = adjust.aje_template(wp, Path(tmp) / "수정분개양식.xlsx")
            data = out.read_bytes()
        self._download(data, f"{e['client_name']}_수정분개양식.xlsx", mailer.XLSX[0] + "/" + mailer.XLSX[1])

    def adjust_make(self, eid):
        import tempfile
        from datetime import datetime

        from ..workpaper import adjust

        e, cfg = self._eng(eid), self.app.cfg()
        _, files = self._form()
        # 칸을 바꿔 올려도 되게 시트 이름으로 조서(TB 탭)와 수정분개(수정분개·AJE 탭)를 가린다
        wps, ajes = [], []
        for name, data in files:
            if not data:
                continue
            if not name.lower().endswith(".xlsx"):
                raise ValueError(f"{name}: .xlsx 파일만 받습니다.")
            sheets = xlsx_sheet_names(data)
            if "TB" in sheets:
                wps.append((name, data))
            elif not sheets:
                raise ValueError(f"{name}: 엑셀(.xlsx) 파일이 아니거나 손상되었습니다.")
            else:
                ajes.append((name, data))      # 회사·법인 자체 양식도 '계정과목·차변·대변' 머리글이 있으면 읽는다
        if not ajes:
            raise ValueError("채운 수정분개 양식을 골라 주세요. [수정분개 양식 내려받기]로 받은 파일입니다.")
        if len(ajes) > 1 or len(wps) > 1:
            raise ValueError("수정분개 양식과 조서는 하나씩만 올려 주세요.")
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            def put(item, sub):
                name, data = item
                d = Path(tmp) / sub
                d.mkdir()
                p = d / Path(name).name
                p.write_bytes(data)
                return p
            aje = put(ajes[0], "aje")
            wp = put(wps[0], "wp") if wps else self._last_workpaper(e)
            stem = wp.stem.removesuffix("_수정분개반영")
            out = Path(tmp) / "out" / f"{stem}_수정분개반영.xlsx"
            res = adjust.apply(wp, aje, out)
            saved = filing.save_file(self.store, e, cfg, out.name, data=out.read_bytes(), received_on=self.app.today(),
                                     source="수정분개", category="조서")
            filing.save_file(self.store, e, cfg, ajes[0][0], data=ajes[0][1], received_on=self.app.today(),
                             source="수정분개", category="조서")
        self.store.set_setting(f"adjust:{e['id']}", {
            "made_at": datetime.now().replace(microsecond=0).isoformat(" "), "path": str(saved["path"]),
            "lines": res.summary()})
        self._redirect(f"/e/{eid}#adjust", res.summary()[0])

    def adjusted_xlsx(self, eid):
        e = self._eng(eid)
        rec = self.store.get_setting(f"adjust:{e['id']}") or {}
        p = Path(rec.get("path") or "")
        if not rec or not p.exists():
            raise ValueError("먼저 [수정분개 반영]을 해 주세요.")
        self._download(p.read_bytes(), p.name, mailer.XLSX[0] + "/" + mailer.XLSX[1])

    def leadsheet_xlsx(self, eid):
        e = self._eng(eid)
        rec = self.store.get_setting(f"leadsheet:{e['id']}") or {}
        p = Path(rec.get("path") or "")
        if not rec or not p.exists():
            raise ValueError("먼저 [조서 채우기]를 해 주세요.")
        self._download(p.read_bytes(), p.name, mailer.XLSX[0] + "/" + mailer.XLSX[1])

    # ── 질문서 ────────────────────────────────────────
    def _question_card(self, e, cfg) -> str:
        eid = e["id"]
        qs = questions.all_(self.store, eid)
        counts = {s: sum(1 for q in qs if q["status"] == s) for s in questions.STATES}
        rows = []
        for q in qs:
            editable = q["status"] in questions.EDITABLE
            fid = f"q{q['id']}"
            if editable:
                qcell = (f"<form id='{fid}' method='post' action='/e/{eid}/questions/{q['id']}'></form>"
                         f"<textarea form='{fid}' name='question' rows='2' style='width:100%;min-width:280px'>{esc(q['question'])}</textarea>")
                acc = f"<input form='{fid}' name='account' value='{esc(q['account'])}' style='width:90px'>"
                amount_txt = "" if q["amount"] is None else f"{q['amount']:,.0f}"
                amt = (f"<input form='{fid}' name='amount' value='{amount_txt}' style='width:110px'>"
                       f"<button form='{fid}' class='ghost' style='margin-top:4px'>저장</button>")
            else:
                qcell = esc(q["question"])
                acc = esc(q["account"])
                amt = "" if q["amount"] is None else f"{q['amount']:,.0f}"
            who = f"<br><span class='mute'>{esc(q['approved_by'] or '')} {esc(q['approved_on'] or '')}</span>" if q["approved_on"] else ""
            ans = esc(q["answer"]) if q["answer"] else "<span class='mute'>-</span>"
            rows.append(f"<tr><td><input type='checkbox' name='ids' value='{q['id']}' form='qact'></td><td>{esc(q['code'])}</td>"
                        f"<td>{acc}</td><td>{qcell}</td><td style='text-align:right'>{amt}</td>"
                        f"<td>{pill(q['status'])}{who}</td><td>{ans}</td></tr>")
        table = ("<table><tr><th></th><th>번호</th><th>계정</th><th>질문</th><th>금액</th><th>상태</th><th>답변</th></tr>"
                 + "".join(rows) + "</table>") if rows else "<p class='mute'>아직 질문이 없습니다.</p>"
        summary = " · ".join(f"{s} {n}" for s, n in counts.items() if n)
        n_ok = counts["승인"]
        mail_btn = (f"<a class='btn' href='/e/{eid}/questions/mail'>승인한 질문 {n_ok}건 메일 미리보기</a>" if n_ok
                    else "<span class='mute'>승인한 질문이 있으면 메일로 보낼 수 있습니다.</span>")
        return f"""<div class="card" id="questions"><h2 style="margin-top:0">질문서</h2>
<p>{summary or ''}</p>
<form id="qact" method="post" action="/e/{eid}/questions/act" class="row">
<button type="button" class="ghost" onclick="document.querySelectorAll('input[name=ids][form=qact]').forEach(c=>c.checked=true)">모두 선택</button>
<button type="button" class="ghost" onclick="document.querySelectorAll('input[name=ids][form=qact]').forEach(c=>c.checked=false)">모두 해제</button>
<button name="action" value="submit" class="ghost">승인 요청</button>
<button name="action" value="approve">승인</button>
<button name="action" value="revert" class="ghost">작성으로 되돌리기</button>
<button name="action" value="close" class="ghost">종결</button>
<button name="action" value="delete" class="ghost" onclick="return confirm('고른 질문을 지웁니다 (작성·승인대기 상태만).')">지우기</button>
<span class="mute">표에서 고른 질문에 적용</span></form>
<div style="margin-top:8px">{table}</div>
<p class="row" style="margin-top:10px">{mail_btn}
<a class="btn ghost" href="/e/{eid}/questions.xlsx">질문서 엑셀</a></p>
<form method="post" action="/e/{eid}/questions" class="row" style="margin-top:10px">
<label>계정<input name="account" style="width:110px"></label>
<label style="flex:1">질문<textarea name="question" rows="2" style="width:100%" required></textarea></label>
<label>금액<input name="amount" style="width:120px"></label><button class="ghost">질문 추가</button></form>
<div class="row" style="margin-top:10px">
<form method="post" action="/e/{eid}/questions/from-confirm"><button class="ghost">채권채무 조회 불일치·미회신으로 초안 만들기</button></form>
<form method="post" action="/e/{eid}/questions/answers" enctype="multipart/form-data" class="row">
<input type="file" name="files" accept=".xlsx"><button class="ghost">답변 엑셀 반영</button></form></div>
{self._auto_question_form(e)}
<p class="mute">승인한 질문만 보낼 수 있습니다. 승인자는 설정의 작성자({esc(cfg.get('author') or '미입력')})로 기록됩니다. 보낸 뒤 회사가 답변 칸을 채운 엑셀을 올리면 답변이 들어갑니다.</p></div>"""

    @staticmethod
    def _amount(v: str | None):
        raw = (v or "").replace(",", "").replace("원", "").strip()
        if not raw:
            return None
        try:
            return float(raw)
        except ValueError as exc:
            raise ValueError("금액은 숫자로 넣어 주세요.") from exc

    def question_add(self, eid):
        f, _ = self._form()
        e = self._eng(eid)
        questions.add(self.store, e["id"], f.get("question", ""), f.get("account", ""), self._amount(f.get("amount")),
                      "직접 입력", self.app.today())
        self._redirect(f"/e/{eid}#questions", "질문을 추가했습니다.")

    def question_edit(self, eid, qid):
        f, _ = self._form()
        self._eng(eid)
        questions.edit(self.store, int(qid), question=f.get("question"), account=f.get("account"),
                       amount=self._amount(f.get("amount")))
        self._redirect(f"/e/{eid}#questions", "질문을 고쳤습니다.")

    def question_act(self, eid):
        f, _ = self._form()
        e = self._eng(eid)
        ids = [int(x) for x in f.get("ids[]", [])]
        if not ids:
            raise ValueError("표에서 질문을 먼저 고르세요.")
        action = f.get("action", "")
        by = self.app.cfg().get("author") or ""
        if action == "approve" and not by:
            raise ValueError("설정에서 작성자(승인자 이름)를 먼저 넣어 주세요.")
        done, skipped = questions.act(self.store, [i for i in ids if (questions.get(self.store, i) or {}).get("engagement_id") == e["id"]],
                                      action, by, self.app.today())
        msg = f"{done}건 처리했습니다."
        if skipped:
            msg += f" 상태가 맞지 않아 건너뜀: {', '.join(skipped)}"
        self._redirect(f"/e/{eid}#questions", msg)

    def _auto_question_form(self, e) -> str:
        eid = e["id"]
        mat = confirmations.materiality(self.store, eid)
        rec = self.store.get_setting(f"autoq:{eid}")
        last = ""
        if rec:
            lines = "".join(f"<li>{esc(x)}</li>" for x in rec.get("notes", []))
            last = f"<p class='mute'>{esc(rec.get('made_at', ''))} 만든 초안 (표본 시드 {esc(rec.get('seed'))})</p><ul class='mute'>{lines}</ul>"
        return f"""<form method="post" action="/e/{eid}/questions/auto" enctype="multipart/form-data" class="row" style="margin-top:10px">
<label>회사 자료 (여러 개)<input type="file" name="files" multiple accept=".xlsx" required></label>
<label>중요성 금액<input name="materiality" value="{'' if mat is None else f'{mat:,.0f}'}" style="width:140px"></label>
<label>표본 시드<input name="seed" style="width:90px" placeholder="자동"></label>
<button class="ghost">장부로 증감분석 질문 초안 만들기</button></form>
<p class="mute">재무상태표·손익계산서(전기 금액 포함), 분개장, 거래처원장(또는 세무사랑 묶음 파일)을 고르면
① 중요성 금액 이상 증감한 계정 ② 전기·당기 같은 금액인 매출채권·대손충당금 ③ 매출원가율 {auto_questions.COST_RATIO_PP:g}%p 이상 변동
④ 분개장의 특이분개(최대 {auto_questions.JE_MAX}건) ⑤ 매출·매입전표 표본 {auto_questions.SAMPLE_N}건씩 증빙 요청을 '작성' 상태로 넣습니다.
이미 만든 질문은 다시 넣지 않습니다.</p>{last}"""

    def question_auto(self, eid):
        import tempfile
        from datetime import datetime

        from ..importers import books as books_mod

        e = self._eng(eid)
        f, files = self._form()
        if not files:
            raise ValueError("회사 자료 파일을 고르세요.")
        mat = self._amount(f.get("materiality"))
        if mat is None:
            mat = confirmations.materiality(self.store, e["id"])
        if not mat:
            raise ValueError("중요성 금액을 넣어 주세요.")
        raw_seed = (f.get("seed") or "").strip()
        if raw_seed and not raw_seed.isdigit():
            raise ValueError("표본 시드는 숫자로 넣어 주세요.")
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            paths = []
            for i, (name, data) in enumerate(files):
                if not name.lower().endswith(".xlsx"):
                    raise ValueError(f"{name}: .xlsx 파일만 받습니다.")
                d = Path(tmp) / f"{i:02d}"
                d.mkdir()
                p = d / Path(name).name
                p.write_bytes(data)
                paths.append(p)
            bk = books_mod.load(paths)
        have = {q["source"] for q in questions.all_(self.store, e["id"])}
        plan = auto_questions.build(bk, mat, int(raw_seed) if raw_seed else None, have)
        added, skipped = auto_questions.apply(self.store, e["id"], plan, self.app.today())
        self.store.set_setting(f"autoq:{e['id']}", {
            "made_at": datetime.now().replace(microsecond=0).isoformat(" "), "seed": plan.seed,
            "notes": bk.notes + plan.notes})
        msg = f"질문 초안 {added}건을 넣었습니다" + (f" (이미 있는 {skipped}건은 건너뜀)" if skipped else "") + "."
        self._redirect(f"/e/{eid}#questions", msg)

    def question_from_confirm(self, eid):
        e = self._eng(eid)
        n = questions.from_confirmations(self.store, e, self.app.today())
        self._redirect(f"/e/{eid}#questions", f"질문 초안 {n}건을 만들었습니다." if n else "새로 만들 질문이 없습니다 (이미 만들었거나 모두 일치).")

    def question_answers(self, eid):
        e = self._eng(eid)
        _, files = self._form()
        if not files:
            raise ValueError("답변을 채운 질문서 엑셀을 고르세요.")
        total, missing = 0, []
        for name, data in files:
            n, miss = questions.import_answers(self.store, e["id"], data, self.app.today())
            total += n
            missing += miss
            filing.save_file(self.store, e, self.app.cfg(), name, data=data, received_on=self.app.today(),
                             source="질문서 답변", category="질문서")
        msg = f"답변 {total}건을 반영했습니다."
        if missing:
            msg += f" 없는 질문 번호: {', '.join(missing)}"
        self._redirect(f"/e/{eid}#questions", msg)

    def question_xlsx(self, eid):
        e = self._eng(eid)
        self._download(questions.xlsx_bytes(e, questions.all_(self.store, e["id"])),
                       f"{e['client_name']}_{e['kind']}_질문서.xlsx", mailer.XLSX[0] + "/" + mailer.XLSX[1])

    def question_mail(self, eid):
        e, cfg = self._eng(eid), self.app.cfg()
        d = questions.draft(self.store, e, cfg, self.app.today())
        sender = mailbox.sender_for(cfg, e)
        send_btn = (f'<form method="post" action="/e/{eid}/questions/mail/send" style="display:inline" '
                    f'onsubmit="return confirm({esc(json.dumps(f"{sender} 계정으로 {d.to} 에게 질문 {len(d.ids)}건을 지금 보냅니다.", ensure_ascii=False))})">'
                    f'<button>지금 보내기</button></form> ' if cfg["mail_accounts"] and d.to else "")
        body = f"""<div class="card"><p><b>받는 사람</b> {esc(d.to) or '<span class=warn>회사 담당자 이메일이 없습니다</span>'}
{f' · <b>참조</b> {esc(d.cc)}' if d.cc else ''}</p><p><b>제목</b> {esc(d.subject)}</p><pre>{esc(d.body)}</pre>
<p class="mute">첨부: {esc(e['client_name'])}_{esc(e['kind'])}_질문서.xlsx (승인한 질문 {len(d.ids)}건, 답변 칸 포함)</p>
<p>{send_btn}<a class="btn ghost" href="/e/{eid}/questions/mail.eml">아웃룩으로 열기 (.eml)</a>
<form method="post" action="/e/{eid}/questions/mail/sent" style="display:inline"><button class="ghost">보냄으로 표시</button></form></p>
<p class="mute">보낸 뒤에는 질문 상태가 '발송'으로 바뀝니다. 아웃룩으로 보냈다면 [보냄으로 표시]를 눌러 주세요.</p></div>
<p><a href="/e/{eid}#questions">← 감사 건으로</a></p>"""
        self._send(page("질문서 메일", body, *self._flash()))

    def question_eml(self, eid):
        e, cfg = self._eng(eid), self.app.cfg()
        d = questions.draft(self.store, e, cfg, self.app.today())
        if not d.to:
            raise ValueError("회사 담당자 이메일이 없습니다. 회사 화면에서 넣어 주세요.")
        msg = questions.message(self.store, e, cfg, d, mailbox.sender_for(cfg, e) or cfg.get("sender_email", ""))
        self._download(mailer.eml_bytes(msg), f"{e['client_name']}_질문서.eml", "message/rfc822")

    def question_send(self, eid):
        e, cfg = self._eng(eid), self.app.cfg()
        d = questions.draft(self.store, e, cfg, self.app.today())
        if not d.to:
            raise ValueError("회사 담당자 이메일이 없습니다.")
        try:
            mailbox.send(cfg, questions.message(self.store, e, cfg, d), mailbox.sender_for(cfg, e))
        except (OSError, mailbox.smtplib.SMTPException) as exc:
            raise ValueError(f"보내지 못했습니다: {exc}") from exc
        questions.mark_sent(self.store, d.ids, self.app.today())
        self.app.note([f"질문서 보냄: {e['client_name']} {len(d.ids)}건 → {d.to}"])
        self._redirect(f"/e/{eid}#questions", f"질문 {len(d.ids)}건을 보냈습니다.")

    def question_mark_sent(self, eid):
        e, cfg = self._eng(eid), self.app.cfg()
        d = questions.draft(self.store, e, cfg, self.app.today())
        questions.mark_sent(self.store, d.ids, self.app.today())
        self._redirect(f"/e/{eid}#questions", f"질문 {len(d.ids)}건을 보냄으로 표시했습니다.")

    def confirm_fetch(self, eid):
        e, cfg = self._eng(eid), self.app.cfg()
        try:
            rec = confirmations.fetch(self.store, e, cfg, self.app.today())
        except nicebiz.LoginError as exc:
            raise ValueError(f"전자조회: {exc}") from exc
        except OSError as exc:
            raise ValueError(f"전자조회 사이트에 연결하지 못했습니다: {exc}") from exc
        c = rec["counts"]
        self._redirect(f"/e/{eid}#confirm", f"전자조회에서 가져왔습니다: 일치 {c['일치']}, 불일치 {c['불일치']}, 미회신 {c['미회신']}")

    def confirm_materiality(self, eid):
        f, _ = self._form()
        e = self._eng(eid)
        raw = (f.get("materiality") or "").replace(",", "").replace("원", "").strip()
        if raw:
            try:
                v = float(raw)
            except ValueError as exc:
                raise ValueError("중요성 금액은 숫자로 넣어 주세요.") from exc
            if v <= 0:
                raise ValueError("중요성 금액은 0보다 커야 합니다.")
        else:
            v = None
        confirmations.set_materiality(self.store, e["id"], v)
        self._redirect(f"/e/{eid}#confirm", "중요성 금액을 저장했습니다." if v else "중요성 금액을 지웠습니다.")

    def confirm_project(self, eid):
        f, _ = self._form()
        e = self._eng(eid)
        confirmations.set_project(self.store, e["id"], f.get("project_id") or None)
        self._redirect(f"/e/{eid}#confirm", "전자조회 프로젝트를 연결했습니다.")

    def confirm_xlsx(self, eid):
        from .. import confirm
        from ..reports import write_excel
        import tempfile

        e = self._eng(eid)
        paths = confirmations.stored_paths(self.store, e["id"])
        if not paths:
            raise ValueError("먼저 전자조회에서 가져오거나 회신현황 엑셀을 올려 주세요.")
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            out = Path(tmp) / "out.xlsx"
            res = confirm.build(paths)
            a = confirmations.assess(confirmations.latest(self.store, e["id"]),
                                     confirmations.materiality(self.store, e["id"]),
                                     float(self.app.cfg().get("confirm_min_reply_rate", 80)))
            if a:
                res.notes = ([("기준 충족: " if a["ok"] else "확인 필요: ")
                              + (" / ".join(a["reasons"]) or f"회신률 {a['rate']}%, 미회신 금액 {a['noreply_amount']:,.0f}원")]
                             + a["notes"] + res.notes)
            write_excel(out, res.sheets())
            data = out.read_bytes()
        self._download(data, f"{e['client_name']}_{e['fiscal_year']}_{e['kind']}_채권채무조회.xlsx",
                       mailer.XLSX[0] + "/" + mailer.XLSX[1])

    def confirm_reconcile(self, eid):
        import tempfile

        from .. import confirm
        from ..reports import write_excel

        e = self._eng(eid)
        _, files = self._form()
        if not files:
            raise ValueError("회신현황 엑셀 파일을 고르세요.")
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            paths = []
            for i, (name, data) in enumerate(files):
                if not name.lower().endswith(".xlsx"):
                    raise ValueError(f"{name}: .xlsx 파일만 받습니다.")
                p = Path(tmp) / f"{i:02d}_{Path(name).name}"
                p.write_bytes(data)
                paths.append(p)
            uploaded_status = [(n, d) for (n, d), p in zip(files, paths) if confirm.is_status_file(_xlsx_rows(p))]
            if uploaded_status:
                confirmations.attach(self.store, e, self.app.cfg(), uploaded_status, self.app.today())
            else:
                paths += confirmations.stored_paths(self.store, e["id"])
            res = confirm.build(paths)
            out = Path(tmp) / "out.xlsx"
            write_excel(out, res.sheets())
            data = out.read_bytes()
        self._download(data, f"{e['client_name']}_{e['fiscal_year']}_{e['kind']}_채권채무조회.xlsx",
                       mailer.XLSX[0] + "/" + mailer.XLSX[1])

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
        self._redirect(f"/e/{eid}#tab-mail", f"{d.label} 메일을 보냄으로 기록했습니다.")

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
        self._redirect(f"/e/{eid}#tab-mail", f"{d.label} 메일을 보냈습니다.")

    def link_done(self, eid, lid):
        self.store.set_link_done(int(lid))
        self._redirect(f"/e/{eid}#tab-files", "링크를 처리함으로 표시했습니다.")

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

    def nice_save(self):
        f, _ = self._form()
        u = f.get("nice_user", "").strip()
        self.store.set_setting("nice_user", u)
        self.store.set_setting("nice_auto", f.get("nice_auto") == "1")
        if f.get("confirm_min_reply_rate"):
            rate = float(f["confirm_min_reply_rate"])
            if not 0 < rate <= 100:
                raise ValueError("회신률 기준은 0~100% 사이로 넣어 주세요.")
            self.store.set_setting("confirm_min_reply_rate", rate)
        note = "전자조회 계정을 저장했습니다."
        if u and f.get("password"):
            note += " " + confirmations.set_password(u, f["password"], remember=f.get("remember") == "1")
        self._redirect("/settings#nice", note)

    def nice_test(self):
        cfg = self.app.cfg()
        try:
            projects = confirmations.refresh_projects(self.store, cfg)
        except nicebiz.LoginError as exc:
            raise ValueError(f"전자조회: {exc}") from exc
        except OSError as exc:
            raise ValueError(f"전자조회 사이트에 연결하지 못했습니다: {exc}") from exc
        live = sum(1 for p in projects if p.get("in_progress", True))
        self._redirect("/settings#nice", f"전자조회 로그인 성공: 진행 중 프로젝트 {live}개, 종료 {len(projects) - live}개를 불러왔습니다.")

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
{inp('question_subject', '질문서 제목', style='width:100%')}{area('question_body', '질문서 본문 ({{질문 표}} {{질문 건수}} {{회신기한}} 사용 가능)', 10)}
{inp('question_due_days', '질문서 회신기한 (보낸 날 + n일)', type='number', style='width:80px')}
</div><p><button>설정 저장</button></p></form>{self._mail_settings(cfg)}{self._nice_settings(cfg)}"""
        self._send(page("설정", body, *self._flash()))

    def _nice_settings(self, cfg) -> str:
        u = confirmations.user(cfg)
        pw = ("들어 있음" if confirmations.ready(cfg) else "<span class='warn'>없음</span>") if u else ""
        n = len(confirmations.cached_projects(self.store))
        test = (f"<form method='post' action='/settings/nice-test' style='margin-top:8px'><button class='ghost'>로그인 시험 (프로젝트 목록 불러오기)</button>"
                f"{f' <span class=mute>불러온 프로젝트 {n}개</span>' if n else ''}</form>") if u else ""
        return f"""<div class="card" id="nice"><h2 style="margin-top:0">전자조회 계정 (한공회 채권채무 전자조회)</h2>
<form method="post" action="/settings/nice" class="row">
<label>아이디<input name="nice_user" value="{esc(u)}" autocomplete="username"></label>
<label>비밀번호 {pw}<input name="password" type="password" autocomplete="current-password" placeholder="바꿀 때만 입력"></label>
<label style="flex-direction:row;gap:4px"><input type="checkbox" name="remember" value="1">윈도우 자격 증명 관리자에 저장</label>
<label>회신률 기준(%)<input name="confirm_min_reply_rate" type="number" min="1" max="100" step="1" value="{float(cfg.get('confirm_min_reply_rate', 80)):g}" style="width:80px"></label>
<label style="flex-direction:row;gap:4px"><input type="checkbox" name="nice_auto" value="1" {'checked' if cfg.get('nice_auto') else ''}>하루 한 번 자동으로 가져오기</label>
<button>저장</button></form>{test}
<p class="mute">회계사 본인 아이디로 로그인합니다. 이용약관상 ID는 다른 사람과 같이 쓰면 안 되므로, 팀원은 각자 자기 PC에서 자기 아이디를 넣어 주세요.
비밀번호를 여러 번 틀려 자동입력방지 문자가 나오면 사이트에서 한 번 직접 로그인한 뒤 다시 시도하세요.
<a href="https://kicpa.nicebizline.com/cc/CC0100M002GE.nice" target="_blank" rel="noopener">이용약관</a></p></div>"""

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
                  "request_body", "reminder_subject", "reminder_body", "dday_body", "question_subject", "question_body"):
            if k in f:
                new[k] = f[k].replace("\r\n", "\n").strip()
        for k in ("request_offset", "reminder_from", "due_offset", "question_due_days"):
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


def _xlsx_rows(path):
    from ..importers.douzone import _rows_of
    return _rows_of(path)


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
            if cfg.get("nice_auto") and confirmations.ready(cfg):
                now = datetime.now()
                if now.hour >= int(cfg.get("send_hour", 9)) and state.get("nice_on") != now.date():
                    state["nice_on"] = now.date()
                    with app.lock:
                        lines = list(lines) + confirmations.fetch_all(app.store, cfg, now.date())
            if lines:
                app.note(lines)
                for line in lines:
                    print(line)
            time.sleep(max(1, int(cfg.get("check_minutes") or 10)) * 60)

    t = threading.Thread(target=loop, daemon=True, name="mail")
    t.start()
    return t
