"""요청·독촉 메일 일정과 메일 초안.

- 최초 요청: 현장감사 시작 D-7 (주말·공휴일이면 그 전 영업일)
- 독촉: D-3부터 D-1까지 매일 + 당일(D-0), 미제출 항목이 있을 때만
- 모든 항목을 받으면 독촉은 멈춘다.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta

KIND_LABEL = {"request": "최초 요청", "reminder": "독촉", "dday": "당일 안내"}


def _d(v) -> date:
    return v if isinstance(v, date) else date.fromisoformat(str(v))


def is_business_day(d: date, cfg: dict) -> bool:
    if cfg.get("skip_weekends", True) and d.weekday() >= 5:
        return False
    return d.isoformat() not in set(cfg.get("holidays") or [])


def request_date(start: date, cfg: dict) -> date:
    d = start - timedelta(days=int(cfg.get("request_offset", 7)))
    while not is_business_day(d, cfg):
        d -= timedelta(days=1)
    return d


def due_date(start: date, cfg: dict) -> date:
    d = start - timedelta(days=int(cfg.get("due_offset", 4)))
    while not is_business_day(d, cfg):
        d -= timedelta(days=1)
    return d


@dataclass
class Planned:
    kind: str
    on: date
    d_n: int

    @property
    def label(self) -> str:
        return f"{KIND_LABEL[self.kind]} (D-{self.d_n})" if self.d_n else KIND_LABEL[self.kind] + " (D-day)"


def plan(engagement: dict, cfg: dict) -> list[Planned]:
    start = _d(engagement["fieldwork_start"])
    out = [Planned("request", request_date(start, cfg), (start - request_date(start, cfg)).days)]
    for n in range(int(cfg.get("reminder_from", 3)), 0, -1):
        d = start - timedelta(days=n)
        if is_business_day(d, cfg):
            out.append(Planned("reminder", d, n))
    out.append(Planned("dday", start, 0))
    return out


def engagement_tag(engagement: dict) -> str:
    """메일 제목에 넣어 답장을 감사 건과 맞추는 꼬리표. 예: C001-2026-중간"""
    return f"{engagement['client_code']}-{engagement['fiscal_year']}-{engagement['kind'][:2]}"


@dataclass
class Draft:
    engagement_id: int
    kind: str
    planned_on: date
    d_n: int
    to: str
    cc: str
    subject: str
    body: str
    item_codes: list[str] = field(default_factory=list)
    overdue: bool = False

    @property
    def label(self) -> str:
        return Planned(self.kind, self.planned_on, self.d_n).label


def _fmt(d) -> str:
    d = _d(d)
    return f"{d.year}년 {d.month}월 {d.day}일({'월화수목금토일'[d.weekday()]})"


def items_table(items: list[dict], with_status: bool = False) -> str:
    lines = []
    for i in items:
        extra = f" - {i['status']}" if with_status else ""
        req = f" (요청일 {i['requested_on']})" if i.get("requested_on") else ""
        note = f" ※ {i['note']}" if i.get("note") else ""
        lines.append(f"[{i['code']}] {i['title']}{req}{extra}{note}")
    return "\n".join(lines)


def render(text: str, values: dict) -> str:
    for k, v in values.items():
        text = text.replace("{" + k + "}", str(v))
    return text


def build_draft(engagement: dict, items: list[dict], planned: Planned, cfg: dict,
                overdue: bool = False) -> Draft:
    start = _d(engagement["fieldwork_start"])
    values = {
        "회계법인": cfg.get("firm_name") or "회계법인",
        "작성자": cfg.get("author") or "",
        "연락처": cfg.get("contact") or cfg.get("sender_email") or "",
        "담당자": engagement.get("contact_name") or "담당자",
        "회사명": engagement["client_name"],
        "사업연도": f"{engagement['fiscal_year']}년",
        "감사구분": engagement["kind"],
        "현장감사 시작일": _fmt(start),
        "현장감사 종료일": _fmt(engagement["fieldwork_end"]),
        "제출기한": _fmt(due_date(start, cfg)),
        "미제출 건수": len(items),
        "D-n": f"D-{planned.d_n}" if planned.d_n else "당일",
        "감사코드": engagement_tag(engagement),
        "요청 자료 표": items_table(items),
        "미제출 자료 표": items_table(items, with_status=True),
    }
    if planned.kind == "request":
        subject, body = cfg["request_subject"], cfg["request_body"]
    else:
        subject = cfg["reminder_subject"]
        body = cfg["dday_body"] if planned.kind == "dday" else cfg["reminder_body"]
    return Draft(engagement["id"], planned.kind, planned.on, planned.d_n,
                 engagement.get("contact_email") or "", engagement.get("cc") or "",
                 render(subject, values), render(body, values), [i["code"] for i in items], overdue)


def due_drafts(store, engagement: dict, today: date, cfg: dict) -> list[Draft]:
    """오늘 보낼 메일. 이미 보낸 것(메일 기록)은 빼고, 놓친 최초 요청은 늦게라도 보낸다."""
    start = _d(engagement["fieldwork_start"])
    end = _d(engagement["fieldwork_end"])
    if today > end:
        return []
    eid = engagement["id"]
    out: list[Draft] = []
    steps = plan(engagement, cfg)
    req = steps[0]
    if not store.request_sent(eid):
        if today >= req.on:
            items = [i for i in store.items(eid) if i["status"] == "미요청"]
            if items:
                out.append(build_draft(engagement, items, req, cfg, overdue=today > req.on))
        return out  # 최초 요청 전에는 독촉하지 않는다
    for p in steps[1:]:
        if p.on == today and not store.mail_sent(eid, p.kind, p.on):
            items = store.outstanding(eid)
            if items:
                out.append(build_draft(engagement, items, p, cfg))
    return out


def all_due(store, today: date, cfg: dict) -> list[tuple[dict, Draft]]:
    out = []
    for e in store.engagements():
        for d in due_drafts(store, e, today, cfg):
            out.append((e, d))
    return out


def record_sent(store, draft: Draft, today: date) -> None:
    """메일을 보낸 뒤 기록한다. 최초 요청이면 '미요청' 항목을 '요청'으로 바꾼다."""
    store.log_mail(draft.engagement_id, draft.kind, draft.planned_on, draft.to, draft.subject,
                   draft.body, draft.item_codes)
    if draft.kind == "request":
        store.mark_requested(draft.engagement_id, today)
