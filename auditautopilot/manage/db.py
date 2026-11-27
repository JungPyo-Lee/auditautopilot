"""감사 건·PBC·메일 기록 저장소 (SQLite, 파이썬 표준 라이브러리).

혼자 쓸 때는 data/audit.db 파일 하나에 모두 저장된다.
"""

from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path

from . import pbc

# 설계서의 상태 흐름: 미요청 → 요청 → 일부수령 → 수령 → 검토완료, 또는 보완요청
STATUSES = ("미요청", "요청", "일부수령", "수령", "검토완료", "보완요청", "해당없음")
OUTSTANDING = ("미요청", "요청", "일부수령", "보완요청")  # 독촉 대상
KINDS = ("중간감사", "기말감사")

SCHEMA = """
CREATE TABLE IF NOT EXISTS clients (
  id INTEGER PRIMARY KEY,
  code TEXT NOT NULL UNIQUE,
  name TEXT NOT NULL,
  seq INTEGER NOT NULL DEFAULT 1,
  contact_name TEXT NOT NULL DEFAULT '',
  contact_email TEXT NOT NULL DEFAULT '',
  cc TEXT NOT NULL DEFAULT '',
  industry TEXT NOT NULL DEFAULT '일반',
  send_from TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS engagements (
  id INTEGER PRIMARY KEY,
  client_id INTEGER NOT NULL REFERENCES clients(id),
  fiscal_year INTEGER NOT NULL,
  kind TEXT NOT NULL,
  fieldwork_start TEXT NOT NULL,
  fieldwork_end TEXT NOT NULL,
  year_seq INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL,
  UNIQUE (client_id, fiscal_year, kind)
);
CREATE TABLE IF NOT EXISTS pbc_items (
  id INTEGER PRIMARY KEY,
  engagement_id INTEGER NOT NULL REFERENCES engagements(id) ON DELETE CASCADE,
  code TEXT NOT NULL,
  section TEXT NOT NULL,
  title TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT '미요청',
  requested_on TEXT,
  received_on TEXT,
  note TEXT NOT NULL DEFAULT '',
  sort INTEGER NOT NULL DEFAULT 0,
  UNIQUE (engagement_id, code)
);
CREATE TABLE IF NOT EXISTS received_files (
  id INTEGER PRIMARY KEY,
  engagement_id INTEGER NOT NULL REFERENCES engagements(id) ON DELETE CASCADE,
  item_id INTEGER REFERENCES pbc_items(id) ON DELETE SET NULL,
  path TEXT NOT NULL,
  original_name TEXT NOT NULL,
  received_on TEXT NOT NULL,
  source TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS mail_log (
  id INTEGER PRIMARY KEY,
  engagement_id INTEGER NOT NULL REFERENCES engagements(id) ON DELETE CASCADE,
  kind TEXT NOT NULL,
  planned_on TEXT NOT NULL,
  sent_at TEXT NOT NULL,
  to_addr TEXT NOT NULL,
  subject TEXT NOT NULL,
  body TEXT NOT NULL,
  item_codes TEXT NOT NULL DEFAULT '[]',
  UNIQUE (engagement_id, kind, planned_on)
);
CREATE TABLE IF NOT EXISTS mail_seen (
  key TEXT PRIMARY KEY,
  seen_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS big_links (
  id INTEGER PRIMARY KEY,
  engagement_id INTEGER NOT NULL REFERENCES engagements(id) ON DELETE CASCADE,
  url TEXT NOT NULL,
  sender TEXT NOT NULL DEFAULT '',
  subject TEXT NOT NULL DEFAULT '',
  received_on TEXT NOT NULL,
  done INTEGER NOT NULL DEFAULT 0,
  UNIQUE (engagement_id, url)
);
CREATE TABLE IF NOT EXISTS questions (
  id INTEGER PRIMARY KEY,
  engagement_id INTEGER NOT NULL REFERENCES engagements(id) ON DELETE CASCADE,
  code TEXT NOT NULL,
  account TEXT NOT NULL DEFAULT '',
  question TEXT NOT NULL,
  amount REAL,
  source TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL DEFAULT '작성',
  created_on TEXT,
  approved_by TEXT,
  approved_on TEXT,
  sent_on TEXT,
  answer TEXT NOT NULL DEFAULT '',
  answered_on TEXT,
  note TEXT NOT NULL DEFAULT '',
  UNIQUE(engagement_id, code)
);
CREATE TABLE IF NOT EXISTS settings (
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL
);
"""


def _iso(d) -> str | None:
    if d is None or d == "":
        return None
    return d.isoformat() if isinstance(d, (date, datetime)) else str(d)


class Store:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        if str(path) != ":memory:":
            self.path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(path), check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.executescript(SCHEMA)
        cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(clients)")}
        if "send_from" not in cols:   # 이전 데이터베이스에 열 추가
            self.conn.execute("ALTER TABLE clients ADD COLUMN send_from TEXT NOT NULL DEFAULT ''")
            self.conn.commit()

    @contextmanager
    def tx(self):
        try:
            yield self.conn
            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise

    def _all(self, sql, *args) -> list[dict]:
        return [dict(r) for r in self.conn.execute(sql, args).fetchall()]

    def _one(self, sql, *args) -> dict | None:
        r = self.conn.execute(sql, args).fetchone()
        return dict(r) if r else None

    # ── 설정 ──────────────────────────────────────────
    def get_setting(self, key: str, default=None):
        r = self._one("SELECT value FROM settings WHERE key = ?", key)
        return json.loads(r["value"]) if r else default

    def set_setting(self, key: str, value) -> None:
        with self.tx() as c:
            c.execute("INSERT INTO settings(key, value) VALUES(?, ?) "
                      "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                      (key, json.dumps(value, ensure_ascii=False)))

    # ── 회사 ──────────────────────────────────────────
    def add_client(self, code: str, name: str, contact_name: str = "", contact_email: str = "",
                   cc: str = "", industry: str = "일반", seq: int | None = None) -> int:
        code, name = code.strip(), name.strip()
        if not code or not name:
            raise ValueError("회사코드와 회사명은 꼭 필요합니다.")
        if industry not in pbc.TEMPLATES:
            raise ValueError(f"업종 템플릿은 {', '.join(pbc.TEMPLATES)} 중 하나입니다.")
        if seq is None:
            seq = (self._one("SELECT COALESCE(MAX(seq), 0) + 1 AS n FROM clients") or {"n": 1})["n"]
        with self.tx() as c:
            cur = c.execute(
                "INSERT INTO clients(code, name, seq, contact_name, contact_email, cc, industry) "
                "VALUES(?, ?, ?, ?, ?, ?, ?)",
                (code, name, seq, contact_name.strip(), contact_email.strip(), cc.strip(), industry))
        return cur.lastrowid

    def update_client(self, client_id: int, **fields) -> None:
        allowed = {"name", "seq", "contact_name", "contact_email", "cc", "industry", "send_from"}
        sets = {k: v for k, v in fields.items() if k in allowed}
        if not sets:
            return
        with self.tx() as c:
            c.execute(f"UPDATE clients SET {', '.join(f'{k} = ?' for k in sets)} WHERE id = ?",
                      (*sets.values(), client_id))

    def clients(self) -> list[dict]:
        return self._all("SELECT * FROM clients ORDER BY seq, code")

    def client(self, client_id: int) -> dict | None:
        return self._one("SELECT * FROM clients WHERE id = ?", client_id)

    # ── 감사 건 ───────────────────────────────────────
    def add_engagement(self, client_id: int, fiscal_year: int, kind: str, fieldwork_start, fieldwork_end,
                       sections=None, template: dict | None = None, year_seq: int = 1) -> int:
        if kind not in KINDS:
            raise ValueError(f"감사구분은 {', '.join(KINDS)} 중 하나입니다.")
        start, end = _iso(fieldwork_start), _iso(fieldwork_end)
        if not start or not end:
            raise ValueError("현장감사 시작일과 종료일이 필요합니다.")
        if end < start:
            raise ValueError("현장감사 종료일이 시작일보다 빠릅니다.")
        client = self.client(client_id)
        if client is None:
            raise ValueError("회사를 찾지 못했습니다.")
        template = template or pbc.TEMPLATES[client["industry"]]
        items = pbc.build_items(template, sections or pbc.DEFAULT_SECTIONS[kind])
        with self.tx() as c:
            cur = c.execute(
                "INSERT INTO engagements(client_id, fiscal_year, kind, fieldwork_start, fieldwork_end, "
                "year_seq, created_at) VALUES(?, ?, ?, ?, ?, ?, ?)",
                (client_id, int(fiscal_year), kind, start, end, year_seq,
                 datetime.now().isoformat(timespec="seconds")))
            eid = cur.lastrowid
            c.executemany(
                "INSERT INTO pbc_items(engagement_id, code, section, title, sort) VALUES(?, ?, ?, ?, ?)",
                [(eid, it["code"], it["section"], it["title"], i) for i, it in enumerate(items)])
        return eid

    def update_engagement(self, engagement_id: int, **fields) -> None:
        allowed = {"fieldwork_start", "fieldwork_end", "year_seq"}
        sets = {k: _iso(v) if k.startswith("fieldwork") else v for k, v in fields.items() if k in allowed}
        if sets:
            with self.tx() as c:
                c.execute(f"UPDATE engagements SET {', '.join(f'{k} = ?' for k in sets)} WHERE id = ?",
                          (*sets.values(), engagement_id))

    _ENG_SQL = ("SELECT e.*, c.code AS client_code, c.name AS client_name, c.seq AS client_seq, "
                "c.contact_name, c.contact_email, c.cc, c.industry, c.send_from "
                "FROM engagements e JOIN clients c ON c.id = e.client_id")

    def engagements(self) -> list[dict]:
        rows = self._all(self._ENG_SQL + " ORDER BY e.fieldwork_start DESC, c.seq")
        counts = self._all("SELECT engagement_id, status, COUNT(*) AS n FROM pbc_items "
                           "GROUP BY engagement_id, status")
        by: dict[int, dict[str, int]] = {}
        for r in counts:
            by.setdefault(r["engagement_id"], {})[r["status"]] = r["n"]
        for e in rows:
            e["counts"] = by.get(e["id"], {})
            e["outstanding"] = sum(e["counts"].get(s, 0) for s in OUTSTANDING)
            e["total"] = sum(n for s, n in e["counts"].items() if s != "해당없음")
        return rows

    def engagement(self, engagement_id: int) -> dict | None:
        return self._one(self._ENG_SQL + " WHERE e.id = ?", engagement_id)

    def delete_engagement(self, engagement_id: int) -> None:
        with self.tx() as c:
            c.execute("DELETE FROM engagements WHERE id = ?", (engagement_id,))

    # ── PBC 항목 ──────────────────────────────────────
    def items(self, engagement_id: int) -> list[dict]:
        return self._all("SELECT * FROM pbc_items WHERE engagement_id = ? ORDER BY sort, id", engagement_id)

    def outstanding(self, engagement_id: int) -> list[dict]:
        return [i for i in self.items(engagement_id) if i["status"] in OUTSTANDING]

    def add_item(self, engagement_id: int, section: str, title: str, code: str | None = None) -> int:
        if section not in pbc.SECTIONS:
            raise ValueError(f"구분은 {', '.join(pbc.SECTIONS)} 중 하나입니다.")
        existing = self.items(engagement_id)
        if not code:
            prefix = pbc.SECTION_PREFIX[section] + "-"
            nums = [int(i["code"][len(prefix):]) for i in existing
                    if i["code"].startswith(prefix) and i["code"][len(prefix):].isdigit()]
            code = f"{prefix}{max(nums, default=0) + 1:02d}"
        sort = max((i["sort"] for i in existing), default=-1) + 1
        with self.tx() as c:
            cur = c.execute("INSERT INTO pbc_items(engagement_id, code, section, title, sort) "
                            "VALUES(?, ?, ?, ?, ?)", (engagement_id, code, section, title.strip(), sort))
        return cur.lastrowid

    def remove_unrequested_items(self, engagement_id: int) -> int:
        """아직 요청하지 않았고(미요청) 받은 파일도 없는 항목을 지운다. 반환: 지운 수"""
        with self.tx() as c:
            cur = c.execute("DELETE FROM pbc_items WHERE engagement_id = ? AND status = '미요청' "
                            "AND id NOT IN (SELECT item_id FROM received_files WHERE item_id IS NOT NULL)", (engagement_id,))
        return cur.rowcount

    def update_item(self, item_id: int, status: str | None = None, note: str | None = None,
                    on: date | None = None) -> None:
        item = self._one("SELECT * FROM pbc_items WHERE id = ?", item_id)
        if item is None:
            raise ValueError("요청 항목을 찾지 못했습니다.")
        today = _iso(on or date.today())
        sets: dict = {}
        if status is not None and status != item["status"]:
            if status not in STATUSES:
                raise ValueError(f"상태는 {', '.join(STATUSES)} 중 하나입니다.")
            sets["status"] = status
            if status == "요청" and not item["requested_on"]:
                sets["requested_on"] = today
            if status in ("일부수령", "수령") and not item["received_on"]:
                sets["received_on"] = today
        if note is not None and note != item["note"]:
            sets["note"] = note
        if sets:
            with self.tx() as c:
                c.execute(f"UPDATE pbc_items SET {', '.join(f'{k} = ?' for k in sets)} WHERE id = ?",
                          (*sets.values(), item_id))

    def mark_requested(self, engagement_id: int, on: date) -> int:
        """최초 요청 메일을 보낸 뒤 '미요청' 항목을 '요청'으로 바꾼다."""
        with self.tx() as c:
            cur = c.execute("UPDATE pbc_items SET status = '요청', requested_on = ? "
                            "WHERE engagement_id = ? AND status = '미요청'", (_iso(on), engagement_id))
        return cur.rowcount

    # ── 받은 파일 ─────────────────────────────────────
    def add_file(self, engagement_id: int, path: str, original_name: str, item_id: int | None,
                 received_on: date, source: str = "") -> int:
        with self.tx() as c:
            cur = c.execute("INSERT INTO received_files(engagement_id, item_id, path, original_name, "
                            "received_on, source) VALUES(?, ?, ?, ?, ?, ?)",
                            (engagement_id, item_id, path, original_name, _iso(received_on), source))
        return cur.lastrowid

    def files(self, engagement_id: int) -> list[dict]:
        return self._all("SELECT f.*, i.code AS item_code, i.title AS item_title FROM received_files f "
                         "LEFT JOIN pbc_items i ON i.id = f.item_id WHERE f.engagement_id = ? "
                         "ORDER BY f.received_on DESC, f.id DESC", engagement_id)

    def file(self, file_id: int) -> dict | None:
        return self._one("SELECT * FROM received_files WHERE id = ?", file_id)

    def set_file(self, file_id: int, path: str, item_id: int | None) -> None:
        with self.tx() as c:
            c.execute("UPDATE received_files SET path = ?, item_id = ? WHERE id = ?", (path, item_id, file_id))

    # ── 메일 기록 ─────────────────────────────────────
    def log_mail(self, engagement_id: int, kind: str, planned_on: date, to_addr: str, subject: str,
                 body: str, item_codes: list[str]) -> int:
        with self.tx() as c:
            cur = c.execute(
                "INSERT INTO mail_log(engagement_id, kind, planned_on, sent_at, to_addr, subject, body, "
                "item_codes) VALUES(?, ?, ?, ?, ?, ?, ?, ?) "
                # 재발송이면 같은 줄을 마지막으로 보낸 내용으로 바꾼다
                "ON CONFLICT(engagement_id, kind, planned_on) DO UPDATE SET sent_at = excluded.sent_at, "
                "to_addr = excluded.to_addr, subject = excluded.subject, body = excluded.body, "
                "item_codes = excluded.item_codes",
                (engagement_id, kind, _iso(planned_on), datetime.now().isoformat(timespec="seconds"),
                 to_addr, subject, body, json.dumps(item_codes, ensure_ascii=False)))
        return cur.lastrowid

    def mails(self, engagement_id: int | None = None) -> list[dict]:
        if engagement_id is None:
            return self._all("SELECT * FROM mail_log ORDER BY sent_at DESC")
        return self._all("SELECT * FROM mail_log WHERE engagement_id = ? ORDER BY sent_at DESC", engagement_id)

    def mail_sent(self, engagement_id: int, kind: str, planned_on: date) -> bool:
        return self._one("SELECT 1 FROM mail_log WHERE engagement_id = ? AND kind = ? AND planned_on = ?",
                         engagement_id, kind, _iso(planned_on)) is not None

    def request_sent(self, engagement_id: int) -> bool:
        return self._one("SELECT 1 FROM mail_log WHERE engagement_id = ? AND kind = 'request'",
                         engagement_id) is not None

    # ── 받은 메일 ─────────────────────────────────────
    def mail_seen(self, key: str) -> bool:
        return self._one("SELECT 1 FROM mail_seen WHERE key = ?", key) is not None

    def mark_mail_seen(self, key: str) -> None:
        with self.tx() as c:
            c.execute("INSERT OR IGNORE INTO mail_seen(key, seen_at) VALUES(?, ?)",
                      (key, datetime.now().isoformat(timespec="seconds")))

    def add_link(self, engagement_id: int, url: str, sender: str, subject: str, received_on: date) -> None:
        with self.tx() as c:
            c.execute("INSERT OR IGNORE INTO big_links(engagement_id, url, sender, subject, received_on) "
                      "VALUES(?, ?, ?, ?, ?)", (engagement_id, url, sender, subject, _iso(received_on)))

    def links(self, engagement_id: int, include_done: bool = False) -> list[dict]:
        sql = "SELECT * FROM big_links WHERE engagement_id = ?" + ("" if include_done else " AND done = 0")
        return self._all(sql + " ORDER BY received_on DESC, id DESC", engagement_id)

    def set_link_done(self, link_id: int) -> None:
        with self.tx() as c:
            c.execute("UPDATE big_links SET done = 1 WHERE id = ?", (link_id,))
