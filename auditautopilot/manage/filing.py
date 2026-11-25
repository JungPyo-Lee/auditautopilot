"""받은 자료를 폴더 규칙대로 저장한다.

규칙 예 (JP님 현재 규칙):
  {루트}/{회사순번:02}.{회사명}/{연도순번:02}.{연도}/{감사구분번호:02}.{감사구분}/{분류번호:02}.{분류}/{원본파일명}
  → data/감사/01.한국상사/01.2026/01.중간감사/01.PBC/매출채권명세.xlsx

쓸 수 있는 칸:
  {루트} {회사코드} {회사명} {회사순번} {연도} {연도순번} {감사구분} {감사구분약칭} {감사구분번호}
  {분류} {분류번호} {수신일} {요청번호} {항목} {원본파일명}
  숫자 칸은 {회사순번:02} 처럼 자릿수를 정할 수 있다.
마지막 칸이 파일 이름이 된다. 같은 이름이 있으면 덮어쓰지 않고 _v2, _v3을 붙인다.
"""

from __future__ import annotations

import re
import shutil
import unicodedata
from datetime import date
from pathlib import Path

INVALID = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
TOKEN = re.compile(r"\{([^{}:]+)(?::([^{}]+))?\}")
MAX_SEGMENT = 60      # 윈도우 경로 길이 제한(260자)에 걸리지 않도록 칸마다 자른다
KIND_NO = {"중간감사": 1, "기말감사": 2}
KIND_SHORT = {"중간감사": "중간", "기말감사": "기말"}
TOKENS = ("루트", "회사코드", "회사명", "회사순번", "연도", "연도순번", "감사구분", "감사구분약칭", "감사구분번호",
          "분류", "분류번호", "수신일", "요청번호", "항목", "원본파일명")


def clean(segment: str, limit: int = MAX_SEGMENT) -> str:
    s = INVALID.sub("_", unicodedata.normalize("NFC", str(segment))).strip().rstrip(". ")
    return s[:limit].rstrip(". ") or "_"


def _stem_ext(name: str) -> tuple[str, str]:
    p = Path(name)
    return (p.stem, p.suffix) if p.suffix and len(p.suffix) <= 6 else (name, "")


def context(engagement: dict, category: str, cfg: dict, original_name: str, item: dict | None = None,
            received_on: date | None = None) -> dict:
    cats = list(cfg.get("categories") or [])
    unmatched = cfg.get("unmatched_category", "미분류")
    cat_no = cats.index(category) + 1 if category in cats else 99
    if category == unmatched:
        cat_no = 99
    return {
        "루트": cfg.get("root", "data/감사"),
        "회사코드": engagement["client_code"],
        "회사명": engagement["client_name"],
        "회사순번": int(engagement.get("client_seq") or 1),
        "연도": int(engagement["fiscal_year"]),
        "연도순번": int(engagement.get("year_seq") or 1),
        "감사구분": engagement["kind"],
        "감사구분약칭": KIND_SHORT.get(engagement["kind"], engagement["kind"]),
        "감사구분번호": KIND_NO.get(engagement["kind"], 9),
        "분류": category,
        "분류번호": cat_no,
        "수신일": (received_on or date.today()).strftime("%Y%m%d"),
        "요청번호": item["code"] if item else "",
        "항목": item["title"] if item else "",
        "원본파일명": original_name,
    }


def _fill(template: str, ctx: dict) -> str:
    def sub(m):
        key, spec = m.group(1).strip(), m.group(2)
        if key not in ctx:
            raise ValueError(f"폴더 규칙에 모르는 칸이 있습니다: {{{key}}}. 쓸 수 있는 칸: {', '.join(TOKENS)}")
        v = ctx[key]
        if spec:
            try:
                return format(int(v), spec)
            except (TypeError, ValueError):
                return str(v)
        return str(v)
    return TOKEN.sub(sub, template)


def build_path(rule: str, ctx: dict) -> Path:
    """규칙으로 저장 경로를 만든다. 원본파일명의 확장자는 항상 유지한다."""
    parts = [p for p in re.split(r"[\\/]", rule) if p.strip()]
    if not parts:
        raise ValueError("폴더 규칙이 비어 있습니다.")
    *dirs, name_tpl = parts
    root_tpl = None
    if dirs and dirs[0].strip() == "{루트}":
        root_tpl = dirs.pop(0)
    folder = [clean(_fill(d, ctx)) for d in dirs]
    stem, ext = _stem_ext(ctx["원본파일명"])
    name_ctx = {**ctx, "원본파일명": stem}
    filled = _fill(name_tpl, name_ctx)
    filled = re.sub(r"_{2,}", "_", filled).strip("_ ") or stem   # 빈 칸(요청번호 없음 등)으로 생긴 '__' 정리
    filename = clean(filled, 120) + ext
    base = Path(ctx["루트"]) if root_tpl else Path()
    return base.joinpath(*folder, filename)


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem, ext = path.stem, path.suffix
    n = 2
    while (candidate := path.with_name(f"{stem}_v{n}{ext}")).exists():
        n += 1
    return candidate


def preview(rule: str, cfg: dict) -> str:
    sample = {"client_code": "C001", "client_name": "한국상사", "client_seq": 1, "fiscal_year": 2026,
              "year_seq": 1, "kind": "기말감사"}
    item = {"code": "기말-12", "title": "전직원퇴직금추계액"}
    ctx = context(sample, (cfg.get("categories") or ["PBC"])[0], cfg, "퇴직금추계.xlsx", item, date(2027, 1, 8))
    return build_path(rule, ctx).as_posix()


# ── 요청 항목 맞추기 ─────────────────────────────────
def _squash(s: str) -> str:
    return re.sub(r"[\s()\[\]{}_\-.,/·:]+", "", unicodedata.normalize("NFC", s or "")).lower()


def match_item(items: list[dict], *texts: str) -> dict | None:
    """파일명·메일 제목에서 요청 항목을 찾는다.

    1) 요청번호(예: 중간-03, 중간_03, [중간 03])가 있으면 그 항목
    2) 자료명이 파일명 안에 통째로 들어 있으면 가장 긴 자료명
    """
    joined = " ".join(t for t in texts if t)
    norm = unicodedata.normalize("NFC", joined)
    for it in items:
        prefix, _, num = it["code"].partition("-")
        if num and re.search(rf"{re.escape(prefix)}\s*[-_ ]?\s*0*{int(num) if num.isdigit() else num}(?!\d)", norm):
            return it
    hay = _squash(joined)
    best = None
    for it in items:
        key = _squash(it["title"])
        if len(key) >= 3 and key in hay and (best is None or len(key) > len(_squash(best["title"]))):
            best = it
    return best


def save_file(store, engagement: dict, cfg: dict, original_name: str, *, data: bytes | None = None,
                       src: Path | None = None, subject: str = "", received_on: date | None = None,
                       item_id: int | None = None, source: str = "", move: bool = False,
                       category: str | None = None) -> dict:
    """한 파일을 저장하고 PBC 상태를 갱신한다. 반환: {'path', 'item', 'matched'}
    category를 주면 PBC 항목과 맞추지 않고 그 분류 폴더에 저장한다 (예: 채권채무조회)."""
    items = store.items(engagement["id"])
    received_on = received_on or date.today()
    if item_id:
        item = next((i for i in items if i["id"] == item_id), None)
    elif category:
        item = None
    else:
        item = match_item(items, original_name, subject)
    if not category:
        category = (cfg.get("categories") or ["PBC"])[0] if item else cfg.get("unmatched_category", "미분류")
    ctx = context(engagement, category, cfg, original_name, item, received_on)
    dest = unique_path(build_path(cfg.get("folder_rule") or "", ctx))
    dest.parent.mkdir(parents=True, exist_ok=True)
    if data is not None:
        dest.write_bytes(data)
    elif src is not None:
        (shutil.move if move else shutil.copy2)(str(src), str(dest))
    else:
        raise ValueError("저장할 내용이 없습니다.")
    store.add_file(engagement["id"], str(dest), original_name, item["id"] if item else None, received_on, source)
    if item and item["status"] in ("미요청", "요청", "보완요청"):
        store.update_item(item["id"], status="수령", on=received_on)
    return {"path": dest, "item": item, "matched": item is not None}


def inbox_dir(engagement: dict, cfg: dict) -> Path:
    """수동으로 받은 파일(대용량 링크 등)을 넣어 두는 폴더."""
    return Path(cfg.get("root", "data/감사")) / "_수신함" / clean(f"{engagement['client_code']}_{engagement['client_name']}"
                                                                 f"_{engagement['fiscal_year']}_{engagement['kind']}")


def sweep_inbox(store, engagement: dict, cfg: dict, today: date | None = None) -> list[dict]:
    """수신함 폴더의 파일을 규칙대로 옮긴다."""
    box = inbox_dir(engagement, cfg)
    box.mkdir(parents=True, exist_ok=True)
    results = []
    for f in sorted(p for p in box.iterdir() if p.is_file() and not p.name.startswith(("~$", "."))):
        results.append(save_file(store, engagement, cfg, f.name, src=f, received_on=today,
                                          source="수신함", move=True))
    return results


def refile(store, engagement: dict, cfg: dict, file_id: int, item_id: int | None) -> Path:
    """미분류 파일을 요청 항목에 연결하고 규칙 위치로 옮긴다."""
    rec = store.file(file_id)
    if rec is None:
        raise ValueError("파일 기록을 찾지 못했습니다.")
    items = store.items(engagement["id"])
    item = next((i for i in items if i["id"] == item_id), None) if item_id else None
    category = (cfg.get("categories") or ["PBC"])[0] if item else cfg.get("unmatched_category", "미분류")
    ctx = context(engagement, category, cfg, rec["original_name"], item, date.fromisoformat(rec["received_on"]))
    dest = unique_path(build_path(cfg.get("folder_rule") or "", ctx))
    src = Path(rec["path"])
    if src.resolve() != dest.resolve():
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(src), str(dest))
    store.set_file(file_id, str(dest), item["id"] if item else None)
    if item and item["status"] in ("미요청", "요청", "보완요청"):
        store.update_item(item["id"], status="수령")
    return dest
