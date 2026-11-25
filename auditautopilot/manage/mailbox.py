"""메일 계정 연결: SMTP로 보내고 IMAP으로 받은 메일의 첨부를 저장한다 (파이썬 표준 라이브러리).

Gmail, 다음(한메일), 네이버 등 IMAP·SMTP를 지원하는 메일이면 쓸 수 있다.
비밀번호는 데이터베이스에 저장하지 않는다. 프로그램이 켜져 있는 동안 메모리에만 두고,
윈도우에서는 원하면 '자격 증명 관리자'에 저장한다(credstore.py).
"""

from __future__ import annotations

import imaplib
import re
import smtplib
import ssl
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from email import policy
from email.message import EmailMessage
from email.parser import BytesParser
from email.utils import getaddresses, parseaddr

from . import filing, schedule

# 서버 주소. 다음·네이버는 https://www.getmailbird.com/setup/ 의 설정값을 따랐다 (2026-11 확인).
PRESETS = {
    "Gmail": {"imap_host": "imap.gmail.com", "imap_port": 993, "smtp_host": "smtp.gmail.com", "smtp_port": 465},
    "다음(한메일)": {"imap_host": "imap.daum.net", "imap_port": 993, "smtp_host": "smtp.daum.net", "smtp_port": 465},
    "네이버": {"imap_host": "imap.naver.com", "imap_port": 993, "smtp_host": "smtp.naver.com", "smtp_port": 587},
}
# 주소의 도메인만으로 회사를 정하면 안 되는 개인 메일 도메인
PUBLIC_DOMAINS = {"gmail.com", "naver.com", "daum.net", "hanmail.net", "kakao.com", "nate.com", "hotmail.com",
                  "outlook.com", "yahoo.com", "icloud.com"}
LINK_HOSTS = ("bigfile", "largefile", "drive.google", "docs.google", "onedrive", "1drv.ms", "sharepoint",
              "dropbox", "wetransfer", "mybox", "naver.me", "kakao")
URL_RE = re.compile(r"https?://[^\s<>\"')\]]+")
TAG_RE = re.compile(r"\[([A-Za-z0-9가-힣_]+)-(\d{4})-(중간|기말)\]")

_PASSWORD: dict[str, str] = {}   # 메모리에만 둔다


def set_password(user: str, password: str, remember: bool = False) -> str:
    _PASSWORD[user] = password
    if remember:
        from . import credstore
        return credstore.save(user, password)
    return "이 프로그램이 켜져 있는 동안만 기억합니다."


def get_password(user: str) -> str | None:
    if user in _PASSWORD:
        return _PASSWORD[user]
    from . import credstore
    pw = credstore.load(user)
    if pw:
        _PASSWORD[user] = pw
    return pw


def accounts(cfg: dict) -> list[dict]:
    return list(cfg.get("mail_accounts") or [])


def sender_for(cfg: dict, engagement: dict | None = None) -> str:
    """보낼 때 쓸 계정: 회사에 정한 계정 → 기본 보내는 계정 → 첫 계정"""
    users = [a["user"] for a in accounts(cfg)]
    for u in ((engagement or {}).get("send_from"), cfg.get("send_from")):
        if u and u in users:
            return u
    return users[0] if users else ""


def ready(cfg: dict) -> list[str]:
    """비밀번호까지 들어 있어 바로 쓸 수 있는 계정들"""
    return [a["user"] for a in accounts(cfg) if get_password(a["user"])]


def account(cfg: dict, user: str | None = None) -> dict:
    user = user or sender_for(cfg)
    acc = next((a for a in accounts(cfg) if a["user"] == user), None)
    if acc is None:
        raise ValueError("설정 > 메일 계정에서 메일 계정을 먼저 추가해 주세요.")
    pw = get_password(user)
    if not pw:
        raise ValueError(f"{user} 의 비밀번호(앱 비밀번호)를 설정 화면에서 넣어 주세요.")
    return {"user": user, "password": pw, "imap_host": acc["imap_host"], "imap_port": int(acc["imap_port"]),
            "smtp_host": acc["smtp_host"], "smtp_port": int(acc["smtp_port"])}


# ── 보내기 ────────────────────────────────────────────
def _smtp(acc: dict):
    ctx = ssl.create_default_context()
    if acc["smtp_port"] == 465:
        s = smtplib.SMTP_SSL(acc["smtp_host"], acc["smtp_port"], context=ctx, timeout=30)
    else:
        s = smtplib.SMTP(acc["smtp_host"], acc["smtp_port"], timeout=30)
        s.starttls(context=ctx)
    s.login(acc["user"], acc["password"])
    return s


def send(cfg: dict, msg: EmailMessage, user: str | None = None) -> None:
    acc = account(cfg, user)
    del msg["X-Unsent"]
    del msg["From"]
    msg["From"] = acc["user"]   # 로그인한 계정과 보내는 주소를 맞춘다 (다르면 거절하는 메일이 많다)
    s = _smtp(acc)
    try:
        s.send_message(msg)
    finally:
        try:
            s.quit()
        except Exception:
            pass


def test_connection(cfg: dict, user: str | None = None) -> str:
    acc = account(cfg, user)
    s = _smtp(acc)
    s.quit()
    m = imaplib.IMAP4_SSL(acc["imap_host"], acc["imap_port"], ssl_context=ssl.create_default_context())
    m.login(acc["user"], acc["password"])
    m.logout()
    return f"{acc['user']}: 보내기(SMTP)와 받기(IMAP) 모두 연결됐습니다."


# ── 받기 ──────────────────────────────────────────────
@dataclass
class FetchResult:
    checked: int = 0
    saved: list[dict] = field(default_factory=list)
    links: int = 0
    skipped: int = 0
    errors: list[str] = field(default_factory=list)


def _find_engagement(store, subject: str, sender: str, others: list[str], today: date) -> dict | None:
    """제목의 [C001-2026-기말] 꼬리표, 없으면 보낸 사람 주소로 진행 중인 감사를 찾는다."""
    engs = store.engagements()
    m = TAG_RE.search(subject or "")
    if m:
        code, year, kind = m.groups()
        hit = [e for e in engs if e["client_code"] == code and str(e["fiscal_year"]) == year
               and e["kind"].startswith(kind)]
        if hit:
            return hit[0]
    addrs = {a.lower() for a in [sender, *others] if a}
    domains = {a.split("@")[-1] for a in addrs if "@" in a} - PUBLIC_DOMAINS

    def client_match(e) -> bool:
        known = {a.strip().lower() for a in re.split(r"[,;\s]+", f"{e['contact_email']},{e['cc']}") if a.strip()}
        if addrs & known:
            return True
        kd = {a.split("@")[-1] for a in known if "@" in a} - PUBLIC_DOMAINS
        return bool(domains & kd)

    active = [e for e in engs if client_match(e)
              and date.fromisoformat(e["fieldwork_end"]) + timedelta(days=60) >= today]
    if not active:
        return None
    # 오늘과 현장감사 기간이 가장 가까운 감사
    return min(active, key=lambda e: abs((date.fromisoformat(e["fieldwork_start"]) - today).days))


def big_file_links(text: str) -> list[str]:
    out = []
    for url in URL_RE.findall(text or ""):
        url = url.rstrip(".,")
        if any(h in url.lower() for h in LINK_HOSTS) and url not in out:
            out.append(url)
    return out


def _body_text(msg) -> str:
    parts = []
    for p in msg.walk():
        if p.get_content_maintype() == "text" and not p.get_filename():
            try:
                t = p.get_content()
            except Exception:
                continue
            if p.get_content_subtype() == "html":
                t = re.sub(r"<a [^>]*href=[\"']([^\"']+)[\"'][^>]*>", r" \1 ", t)
                t = re.sub(r"<[^>]+>", " ", t)
            parts.append(t)
    return "\n".join(parts)


def process_message(store, cfg: dict, raw: bytes, today: date) -> dict:
    """메일 한 통을 처리한다. 반환: {'engagement', 'files', 'links'}"""
    msg = BytesParser(policy=policy.default).parsebytes(raw)
    subject = str(msg.get("Subject") or "")
    sender = parseaddr(str(msg.get("From") or ""))[1]
    others = [a for _, a in getaddresses([str(msg.get("Cc") or ""), str(msg.get("Reply-To") or "")])]
    e = _find_engagement(store, subject, sender, others, today)
    if e is None:
        return {"engagement": None, "files": [], "links": []}
    files = []
    for part in msg.iter_attachments():
        name = part.get_filename()
        data = part.get_payload(decode=True)
        if not name or not data:
            continue
        files.append(filing.save_file(store, e, cfg, name, data=data, subject=subject, received_on=today,
                                      source=f"메일 {sender}"))
    links = big_file_links(_body_text(msg))
    for url in links:
        store.add_link(e["id"], url, sender, subject, today)
    return {"engagement": e, "files": files, "links": links}


def fetch(store, cfg: dict, today: date | None = None, days: int = 14) -> FetchResult:
    """비밀번호가 들어 있는 모든 계정의 받은편지함을 확인한다. 한 계정이 실패해도 나머지는 계속한다."""
    today = today or date.today()
    res = FetchResult()
    users = ready(cfg)
    if not users:
        raise ValueError("비밀번호까지 넣은 메일 계정이 없습니다. 설정 > 메일 계정을 확인해 주세요.")
    for user in users:
        try:
            _fetch_one(store, cfg, account(cfg, user), today, days, res)
        except (OSError, imaplib.IMAP4.error) as exc:
            res.errors.append(f"{user}: {exc}")
    return res


_MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


def _fetch_one(store, cfg: dict, acc: dict, today: date, days: int, res: FetchResult) -> None:
    m = imaplib.IMAP4_SSL(acc["imap_host"], acc["imap_port"], ssl_context=ssl.create_default_context())
    try:
        m.login(acc["user"], acc["password"])
        m.select("INBOX", readonly=True)
        d0 = today - timedelta(days=days)
        since = f"{d0.day:02d}-{_MONTHS[d0.month - 1]}-{d0.year}"   # 지역 설정과 무관하게 영문 월
        typ, data = m.uid("SEARCH", None, "SINCE", since)
        uids = (data[0] or b"").split() if typ == "OK" else []
        for uid in uids:
            key = f"{acc['user']}:{uid.decode()}"
            if store.mail_seen(key):
                continue
            typ, parts = m.uid("FETCH", uid, "(BODY.PEEK[])")
            raw = next((p[1] for p in parts if isinstance(p, tuple)), None)
            if raw is None:
                continue
            store.mark_mail_seen(key)
            mid = str(BytesParser(policy=policy.default).parsebytes(raw, headersonly=True).get("Message-ID") or "").strip()
            if mid and store.mail_seen("msgid:" + mid):
                continue   # 두 계정에 함께 온 같은 메일
            if mid:
                store.mark_mail_seen("msgid:" + mid)
            res.checked += 1
            out = process_message(store, cfg, raw, today)
            if out["engagement"] is None:
                res.skipped += 1
            res.saved.extend(out["files"])
            res.links += len(out["links"])
    finally:
        try:
            m.logout()
        except Exception:
            pass


# ── 자동 실행 ─────────────────────────────────────────
def send_due(store, cfg: dict, today: date, build_message) -> list[str]:
    """오늘 보낼 메일을 실제로 보내고 기록한다. 반환: 보낸 메일 설명 목록."""
    sent = []
    for e, d in schedule.all_due(store, today, cfg):
        if not d.to:
            continue
        msg = build_message(e, d)
        send(cfg, msg, sender_for(cfg, e))
        schedule.record_sent(store, d, today)
        sent.append(f"{e['client_name']} {d.label}")
    return sent


def run_cycle(store, cfg: dict, now: datetime, build_message, state: dict) -> list[str]:
    """백그라운드에서 주기적으로 부른다. 메일 확인 + (자동 발송이 켜져 있으면) 정해진 시각 이후 발송."""
    log = []
    if not ready(cfg):
        return log
    try:
        r = fetch(store, cfg, now.date())
        if r.saved or r.links:
            log.append(f"메일 확인: 첨부 {len(r.saved)}개 저장, 대용량 링크 {r.links}개")
        log.extend(f"메일 확인 실패 {e}" for e in r.errors)
    except Exception as exc:  # 네트워크 오류 등은 다음 주기에 다시
        log.append(f"메일 확인 실패: {exc}")
    if cfg.get("auto_send") and now.hour >= int(cfg.get("send_hour", 9)) and state.get("sent_on") != now.date():
        try:
            sent = send_due(store, cfg, now.date(), build_message)
            state["sent_on"] = now.date()
            log.extend(f"자동 발송: {s}" for s in sent)
        except Exception as exc:
            log.append(f"자동 발송 실패: {exc}")
    return log
