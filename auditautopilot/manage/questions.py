"""질문서: 작성 → 승인대기 → 승인 → 발송 → 답변받음 → 종결.

승인한 질문만 회사에 보낼 수 있다. 질문서 엑셀(답변 칸 포함)을 붙여 보내고,
회사가 답변 칸을 채워 보낸 엑셀을 올리면 질문마다 답변이 들어간다.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from datetime import date, timedelta
from email.message import EmailMessage
from email.utils import formatdate

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from .. import confirm
from . import schedule

STATES = ("작성", "승인대기", "승인", "발송", "답변받음", "종결")
EDITABLE = ("작성", "승인대기")
# 버튼 → (바꿀 수 있는 이전 상태, 새 상태)
ACTIONS = {
    "submit": (("작성",), "승인대기"),
    "approve": (("작성", "승인대기"), "승인"),
    "revert": (("승인대기", "승인"), "작성"),
    "close": (("발송", "답변받음"), "종결"),
    "reopen": (("종결",), "답변받음"),
}
XLSX = ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def _iso(d) -> str | None:
    return d.isoformat() if isinstance(d, date) else d


def all_(store, eid: int) -> list[dict]:
    return store._all("SELECT * FROM questions WHERE engagement_id = ? ORDER BY code", eid)


def get(store, qid: int) -> dict | None:
    return store._one("SELECT * FROM questions WHERE id = ?", qid)


def _next_code(store, eid: int) -> str:
    codes = [r["code"] for r in store._all("SELECT code FROM questions WHERE engagement_id = ?", eid)]
    n = max((int(m.group(1)) for c in codes if (m := re.fullmatch(r"Q-(\d+)", c))), default=0)
    return f"Q-{n + 1:03d}"


def add(store, eid: int, question: str, account: str = "", amount: float | None = None, source: str = "",
        today: date | None = None) -> int:
    question = (question or "").strip()
    if not question:
        raise ValueError("질문 내용을 넣어 주세요.")
    with store.tx() as c:
        cur = c.execute("INSERT INTO questions(engagement_id, code, account, question, amount, source, created_on) "
                        "VALUES(?, ?, ?, ?, ?, ?, ?)",
                        (eid, _next_code(store, eid), account.strip(), question, amount, source,
                         _iso(today or date.today())))
    return cur.lastrowid


def edit(store, qid: int, question: str | None = None, account: str | None = None, amount=..., note: str | None = None):
    q = get(store, qid)
    if q is None:
        raise ValueError("질문을 찾지 못했습니다.")
    if (question is not None or account is not None or amount is not ...) and q["status"] not in EDITABLE:
        raise ValueError(f"{q['code']}: 승인되었거나 보낸 질문은 고칠 수 없습니다. [되돌리기] 후 고쳐 주세요.")
    with store.tx() as c:
        if question is not None:
            if not question.strip():
                raise ValueError("질문 내용을 비울 수 없습니다.")
            c.execute("UPDATE questions SET question = ? WHERE id = ?", (question.strip(), qid))
        if account is not None:
            c.execute("UPDATE questions SET account = ? WHERE id = ?", (account.strip(), qid))
        if amount is not ...:
            c.execute("UPDATE questions SET amount = ? WHERE id = ?", (amount, qid))
        if note is not None:
            c.execute("UPDATE questions SET note = ? WHERE id = ?", (note, qid))


def act(store, ids: list[int], action: str, by: str = "", today: date | None = None) -> tuple[int, list[str]]:
    """여러 질문의 상태를 한 번에 바꾼다. 반환: (바꾼 건수, 건너뛴 질문 번호)"""
    if action == "delete":
        done, skipped = 0, []
        for qid in ids:
            q = get(store, qid)
            if q and q["status"] in EDITABLE:
                with store.tx() as c:
                    c.execute("DELETE FROM questions WHERE id = ?", (qid,))
                done += 1
            elif q:
                skipped.append(q["code"])
        return done, skipped
    if action not in ACTIONS:
        raise ValueError("알 수 없는 작업입니다.")
    allowed, new = ACTIONS[action]
    today = today or date.today()
    done, skipped = 0, []
    for qid in ids:
        q = get(store, qid)
        if not q:
            continue
        if q["status"] not in allowed:
            skipped.append(q["code"])
            continue
        with store.tx() as c:
            if new == "승인":
                c.execute("UPDATE questions SET status = ?, approved_by = ?, approved_on = ? WHERE id = ?",
                          (new, by, _iso(today), qid))
            elif new == "작성":
                c.execute("UPDATE questions SET status = ?, approved_by = NULL, approved_on = NULL WHERE id = ?", (new, qid))
            else:
                c.execute("UPDATE questions SET status = ? WHERE id = ?", (new, qid))
        done += 1
    return done, skipped


def approved(store, eid: int) -> list[dict]:
    return [q for q in all_(store, eid) if q["status"] == "승인"]


def mark_sent(store, ids: list[int], today: date | None = None) -> None:
    with store.tx() as c:
        for qid in ids:
            c.execute("UPDATE questions SET status = '발송', sent_on = ? WHERE id = ? AND status = '승인'",
                      (_iso(today or date.today()), qid))


# ── 채권채무 조회에서 초안 만들기 ────────────────────────
def from_confirmations(store, engagement: dict, today: date | None = None) -> int:
    """마지막으로 가져온 회신현황의 불일치·미회신 건으로 질문 초안을 만든다. 이미 만든 건은 건너뛴다."""
    from . import confirmations

    paths = confirmations.stored_paths(store, engagement["id"])
    if not paths:
        raise ValueError("먼저 채권채무 조회 회신현황을 가져와 주세요.")
    res = confirm.build(paths)
    have = {q["source"] for q in all_(store, engagement["id"])}
    n = 0
    for r in res.rows:
        if r["status"] == confirm.STATUS_OK:
            continue
        key = f"채권채무조회:{r['biz_no'] or confirm.norm_name(r['partner'])}:{r['label']}:{r['currency']}"
        if key in have:
            continue
        amt = f"{r['sent']:,.0f}원" if r["currency"] == "KRW" else f"{r['currency']} {r['sent']:,.2f}"
        if r["status"] == confirm.STATUS_DIFF:
            diff = r["diff"] or 0
            dtxt = f"{diff:+,.0f}원" if r["currency"] == "KRW" else f"{r['currency']} {diff:+,.2f}"
            text = (f"{r['partner']} {r['account']} 채권채무조회 결과, 회사 장부 금액 {amt}과 거래처 회신 금액이 "
                    f"{dtxt} 차이 납니다. 차이 원인(시점 차이, 미착, 반품·할인 등)과 조정 내역을 설명해 주시고 근거 자료를 보내 주세요.")
        else:
            text = (f"{r['partner']} {r['account']} {amt} 조회서에 거래처 회신이 없습니다. "
                    f"기말 이후 입금(지급) 내역, 세금계산서·거래명세서 등 잔액을 확인할 수 있는 자료를 보내 주세요.")
        add(store, engagement["id"], text, r["account"], r["sent"] if r["currency"] == "KRW" else None, key, today)
        n += 1
    return n


# ── 질문서 엑셀 ─────────────────────────────────────────
HEADERS = ["번호", "계정", "질문", "금액", "답변", "답변일", "비고"]


def xlsx_bytes(engagement: dict, qs: list[dict]) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "질문서"
    ws.append([f"{engagement['client_name']} {engagement['fiscal_year']} {engagement['kind']} 질문서"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append(["'답변' 칸에 적어 회신해 주세요. 근거 자료는 파일 이름 앞에 질문 번호(예: Q-001)를 붙여 보내 주시면 됩니다."])
    ws.append([])
    ws.append(HEADERS)
    head = ws.max_row
    for c in ws[head]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDE6F4")
    for q in qs:
        ws.append([q["code"], q["account"], q["question"], q["amount"], q["answer"] or None, q["answered_on"], None])
    for col, w in zip("ABCDEFG", (9, 14, 70, 16, 60, 12, 20)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=head + 1):
        row[2].alignment = row[4].alignment = Alignment(wrap_text=True, vertical="top")
        row[3].number_format = "#,##0"
        row[4].fill = PatternFill("solid", fgColor="FFF8DC")
    ws.freeze_panes = ws.cell(row=head + 1, column=1)
    _request_sheet(wb, qs)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


REQUEST_HEADERS = ["번호", "구분", "전표일자", "전표번호", "거래처·적요", "금액", "요청 자료", "제출 여부"]


def _request_sheet(wb, qs: list[dict]) -> None:
    """표본 증빙 요청(source '샘플:매출:2026-03-31:50266')을 따로 모은 '요청자료' 탭."""
    rows = [q for q in qs if (q.get("source") or "").startswith("샘플:")]
    if not rows:
        return
    ws = wb.create_sheet("요청자료")
    ws.append(["요청자료 목록 (표본 증빙)"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append(["자료 파일 이름 앞에 번호(예: Q-001)를 붙여 보내 주세요. 답변은 '질문서' 탭의 같은 번호 칸에 적어 주시면 됩니다."])
    ws.append([])
    ws.append(REQUEST_HEADERS)
    head = ws.max_row
    for c in ws[head]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDE6F4")
    for q in rows:
        _, kind, day, no = (q["source"].split(":", 3) + ["", "", ""])[:4]
        head_txt, _, ask = q["question"].rpartition(": ")
        desc = head_txt.split(f"{day} {no}", 1)[-1]
        if q["amount"] is not None:
            desc = desc.replace(f"{q['amount']:,.0f}원", "")
        desc = re.sub(r"\s+", " ", desc).strip()
        ws.append([q["code"], f"{kind} 표본", day, no, desc, q["amount"], ask, None])
    for col, w in zip("ABCDEFGH", (9, 10, 12, 10, 36, 16, 60, 12)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=head + 1):
        row[5].number_format = "#,##0"
        row[4].alignment = row[6].alignment = Alignment(wrap_text=True, vertical="top")
        row[7].fill = PatternFill("solid", fgColor="FFF8DC")


def import_answers(store, eid: int, raw: bytes, today: date | None = None) -> tuple[int, list[str]]:
    """답변 칸을 채운 질문서 엑셀을 읽는다. 반환: (반영 건수, 못 찾은 번호)"""
    wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    try:
        rows = [tuple(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()
    head = next((i for i, r in enumerate(rows[:30]) if r and "번호" in [str(x or "").strip() for x in r]
                 and "답변" in [str(x or "").strip() for x in r]), None)
    if head is None:
        raise ValueError("질문서 엑셀이 아닙니다 ('번호'와 '답변' 머리글을 찾지 못함).")
    pos = {str(v or "").strip(): i for i, v in enumerate(rows[head])}
    by_code = {q["code"]: q for q in all_(store, eid)}
    n, missing = 0, []
    for r in rows[head + 1:]:
        code = str(r[pos["번호"]] or "").strip() if pos["번호"] < len(r) else ""
        ans = r[pos["답변"]] if pos["답변"] < len(r) else None
        if not code or ans is None or not str(ans).strip():
            continue
        q = by_code.get(code)
        if q is None:
            missing.append(code)
            continue
        if str(ans).strip() == (q["answer"] or "").strip():
            continue
        with store.tx() as c:
            c.execute("UPDATE questions SET answer = ?, answered_on = ?, "
                      "status = CASE WHEN status IN ('발송', '승인') THEN '답변받음' ELSE status END WHERE id = ?",
                      (str(ans).strip(), _iso(today or date.today()), q["id"]))
        n += 1
    return n, missing


# ── 메일 ───────────────────────────────────────────────
@dataclass
class QDraft:
    to: str
    cc: str
    subject: str
    body: str
    ids: list[int]


def table_text(qs: list[dict]) -> str:
    lines = []
    for q in qs:
        amt = f" ({q['amount']:,.0f}원)" if q.get("amount") else ""
        acc = f"[{q['account']}] " if q.get("account") else ""
        lines.append(f"{q['code']}. {acc}{q['question']}{amt}")
    return "\n".join(lines)


def draft(store, engagement: dict, cfg: dict, today: date | None = None) -> QDraft:
    qs = approved(store, engagement["id"])
    if not qs:
        raise ValueError("승인한 질문이 없습니다. 질문을 승인한 뒤 보내 주세요.")
    today = today or date.today()
    due = today + timedelta(days=int(cfg.get("question_due_days", 5)))
    values = {
        "회계법인": cfg.get("firm_name") or "회계법인",
        "작성자": cfg.get("author") or "",
        "연락처": cfg.get("contact") or cfg.get("sender_email") or "",
        "담당자": engagement.get("contact_name") or "담당자",
        "회사명": engagement["client_name"],
        "사업연도": f"{engagement['fiscal_year']}년",
        "감사구분": engagement["kind"],
        "감사코드": schedule.engagement_tag(engagement),
        "질문 건수": len(qs),
        "회신기한": f"{due.month}월 {due.day}일",
        "질문 표": table_text(qs),
    }
    return QDraft(engagement.get("contact_email") or "", engagement.get("cc") or "",
                  schedule.render(cfg.get("question_subject") or "", values),
                  schedule.render(cfg.get("question_body") or "", values), [q["id"] for q in qs])


def message(store, engagement: dict, cfg: dict, d: QDraft, sender: str = "") -> EmailMessage:
    msg = EmailMessage()
    if sender:
        msg["From"] = sender
    msg["To"] = d.to
    if d.cc:
        msg["Cc"] = d.cc
    msg["Subject"] = d.subject
    msg["Date"] = formatdate(localtime=True)
    msg["X-Unsent"] = "1"
    msg.set_content(d.body)
    qs = [q for q in all_(store, engagement["id"]) if q["id"] in d.ids]
    msg.add_attachment(xlsx_bytes(engagement, qs), maintype=XLSX[0], subtype=XLSX[1],
                       filename=f"{engagement['client_name']}_{engagement['kind']}_질문서.xlsx")
    return msg
