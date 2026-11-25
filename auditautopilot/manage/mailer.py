"""메일 파일(.eml) 만들기.

메일 계정 연결(SMTP·Microsoft 365 등)은 사용하는 메일이 정해지면 붙인다.
그 전까지는 초안을 .eml 로 내려받아 아웃룩에서 열고 [보내기]를 누르면 된다.
'X-Unsent: 1' 머리글이 있으면 아웃룩은 받은 메일이 아니라 보낼 수 있는 새 메일로 연다.
"""

from __future__ import annotations

import io
from email.message import EmailMessage
from email.utils import formatdate

from ..reports import write_excel

XLSX = ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def pbc_rows(items: list[dict]) -> list[dict]:
    return [{"요청번호": i["code"], "구분": i["section"], "자료명": i["title"], "상태": i["status"],
             "요청일": i.get("requested_on") or "", "비고": i.get("note") or ""} for i in items]


def pbc_xlsx_bytes(items: list[dict], sheet: str = "자료요청목록") -> bytes:
    import tempfile
    from pathlib import Path

    with tempfile.TemporaryDirectory() as d:
        p = write_excel(Path(d) / "pbc.xlsx", {sheet: pbc_rows(items)})
        return p.read_bytes()


def build_message(draft, sender: str, items: list[dict], attachment_name: str | None = None) -> EmailMessage:
    msg = EmailMessage()
    if sender:
        msg["From"] = sender
    msg["To"] = draft.to
    if draft.cc:
        msg["Cc"] = draft.cc
    msg["Subject"] = draft.subject
    msg["Date"] = formatdate(localtime=True)
    msg["X-Unsent"] = "1"
    msg.set_content(draft.body)
    if attachment_name and items:
        msg.add_attachment(pbc_xlsx_bytes(items), maintype=XLSX[0], subtype=XLSX[1], filename=attachment_name)
    return msg


def eml_bytes(msg: EmailMessage) -> bytes:
    buf = io.BytesIO()
    buf.write(msg.as_bytes())
    return buf.getvalue()
