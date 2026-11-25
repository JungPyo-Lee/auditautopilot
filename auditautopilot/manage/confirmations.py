"""감사 건마다 전자조회 회신현황을 가져와 보관하고 대사 결과를 만든다.

- 전자조회 사이트에 로그인해서 회신현황·미회신현황 엑셀을 받는다 (nicebiz.py)
- 받은 파일은 폴더 규칙의 '채권채무조회' 분류에 저장한다 (내용이 바뀌었을 때만 새로 저장)
- 마지막 결과 요약은 settings 표에 'confirm:<감사건 id>'로 둔다
"""

from __future__ import annotations

import hashlib
import io
from datetime import date, datetime
from pathlib import Path

import openpyxl

from .. import confirm
from . import filing, mailbox, nicebiz

CATEGORY = "채권채무조회"
FILES = {"reply": "전자조회_회신현황.xlsx", "noreply": "전자조회_미회신현황.xlsx"}


def _key(eid: int) -> str:
    return f"confirm:{eid}"


def latest(store, eid: int) -> dict | None:
    return store.get_setting(_key(eid))


def project_of(store, eid: int) -> str | None:
    return store.get_setting(f"nice_project:{eid}")


def set_project(store, eid: int, project_id: str | None) -> None:
    store.set_setting(f"nice_project:{eid}", project_id or None)


def cached_projects(store) -> list[dict]:
    return store.get_setting("nice_projects", []) or []


def user(cfg: dict) -> str:
    return (cfg.get("nice_user") or "").strip()


def set_password(user_id: str, password: str, remember: bool = False) -> str:
    return mailbox.set_password(nicebiz.CRED_KEY + user_id, password, remember)


def ready(cfg: dict) -> bool:
    u = user(cfg)
    return bool(u and mailbox.get_password(nicebiz.CRED_KEY + u))


def _rows(raw: bytes) -> list[tuple]:
    wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    try:
        return [tuple(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()


def login(cfg: dict, client_factory=nicebiz.Client):
    u = user(cfg)
    if not u:
        raise ValueError("설정 > 전자조회 계정에 아이디를 넣어 주세요.")
    pw = mailbox.get_password(nicebiz.CRED_KEY + u)
    if not pw:
        raise ValueError("설정 > 전자조회 계정에 비밀번호를 넣어 주세요.")
    c = client_factory()
    c.login(u, pw)
    return c


def refresh_projects(store, cfg: dict, client=None, client_factory=nicebiz.Client) -> list[dict]:
    c = client or login(cfg, client_factory)
    projects = [{**p.__dict__, "label": p.label} for p in c.projects(include_ended=True)]
    store.set_setting("nice_projects", projects)
    return projects


def _record(store, engagement: dict, sources: list[tuple[str, bytes]], extra: dict) -> dict:
    rows = []
    for _, raw in sources:
        r = _rows(raw)
        if confirm.is_status_file(r):
            rows += confirm.read_status_rows(r)
    res = confirm.reconcile(rows)
    rec = {**extra,
           "summary": res.summary,
           "counts": {s: len(res.by_status(s)) for s in (confirm.STATUS_OK, confirm.STATUS_DIFF, confirm.STATUS_NONE)}}
    store.set_setting(_key(engagement["id"]), rec)
    return rec


def fetch(store, engagement: dict, cfg: dict, today: date | None = None, client_factory=nicebiz.Client) -> dict:
    """사이트에서 이 감사 건의 회신·미회신 현황을 받아 저장하고 요약을 돌려준다."""
    today = today or date.today()
    c = login(cfg, client_factory)
    projects = refresh_projects(store, cfg, client=c)
    pid = project_of(store, engagement["id"])
    proj = next((p for p in projects if p["id"] == pid), None)
    if proj is None:
        found = nicebiz.match_project([nicebiz.Project(**{k: p[k] for k in ("id", "year", "name", "kind", "created", "in_progress")})
                                       for p in projects], engagement["client_name"], engagement["fiscal_year"])
        if found is None:
            raise ValueError(f"전자조회 프로젝트 중 '{engagement['client_name']}'을(를) 찾지 못했습니다. "
                             "감사 건 화면에서 프로젝트를 직접 골라 주세요.")
        proj = next(p for p in projects if p["id"] == found.id)
        set_project(store, engagement["id"], proj["id"])
    prev = latest(store, engagement["id"]) or {}
    hashes, paths, sources = dict(prev.get("hashes") or {}), dict(prev.get("paths") or {}), []
    for kind, name in FILES.items():
        raw = c.excel(proj["id"], kind, proj.get("in_progress", True))
        sources.append((name, raw))
        h = hashlib.sha256(raw).hexdigest()
        if hashes.get(kind) != h or not paths.get(kind) or not Path(paths[kind]).exists():
            saved = filing.save_file(store, engagement, cfg, f"{today:%Y%m%d}_{name}", data=raw, received_on=today,
                                     source="전자조회", category=CATEGORY)
            paths[kind], hashes[kind] = str(saved["path"]), h
    return _record(store, engagement, sources, {
        "project": proj["label"], "project_id": proj["id"], "fetched_at": datetime.now().replace(microsecond=0).isoformat(" "),
        "paths": paths, "hashes": hashes, "source": "전자조회 사이트"})


def attach(store, engagement: dict, cfg: dict, files: list[tuple[str, bytes]], today: date | None = None) -> dict:
    """사이트에서 직접 내려받은 엑셀을 올렸을 때."""
    today = today or date.today()
    status = [(n, d) for n, d in files if confirm.is_status_file(_rows(d))]
    if not status:
        raise ValueError("전자조회 회신현황 형식의 엑셀이 없습니다.")
    paths = {}
    for i, (name, raw) in enumerate(status):
        saved = filing.save_file(store, engagement, cfg, name, data=raw, received_on=today, source="업로드",
                                 category=CATEGORY)
        paths[f"upload{i}"] = str(saved["path"])
    return _record(store, engagement, status, {"fetched_at": datetime.now().replace(microsecond=0).isoformat(" "),
                                                "paths": paths, "source": "직접 올림"})


def stored_paths(store, eid: int) -> list[Path]:
    rec = latest(store, eid) or {}
    return [Path(p) for p in (rec.get("paths") or {}).values() if Path(p).exists()]


def fetch_all(store, cfg: dict, today: date | None = None, client_factory=nicebiz.Client) -> list[str]:
    """프로젝트가 연결된 감사 건을 모두 가져온다 (백그라운드에서 하루 한 번)."""
    log = []
    for e in store.engagements():
        if not project_of(store, e["id"]):
            continue
        try:
            c = fetch(store, e, cfg, today, client_factory)["counts"]
            log.append(f"전자조회 {e['client_name']}: 일치 {c['일치']}, 불일치 {c['불일치']}, 미회신 {c['미회신']}")
        except Exception as exc:  # noqa: BLE001 - 한 건이 실패해도 나머지는 계속
            log.append(f"전자조회 {e['client_name']} 실패: {exc}")
    return log
