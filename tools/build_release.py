"""윈도우용 배포 패키지(zip)를 만든다.

    python tools/build_release.py            # dist/auditautopilot-<버전>-windows.zip

설치 없이 압축만 풀어 쓰도록 만든다.
- app/      : auditautopilot 소스 + openpyxl·et_xmlfile (순수 파이썬, 압축을 푼 상태)
- wheels/   : 같은 것을 pip로 설치하고 싶을 때 쓰는 wheel 파일 (선택)
- *.bat     : 실행.bat(관리 화면), 명령창.bat(명령 실행), pip설치.bat(선택)
- 사용설명서.html : docs/manual.md를 HTML로 바꾼 것

회사 PC 앱 제어가 서명 없는 DLL·exe와 가상환경 pip.exe를 막으므로 exe로 묶지 않고
python.org 파이썬(서명됨)이 .py 파일을 그대로 읽게 한다. PYTHONPATH만 app 폴더로 잡는다.
"""

from __future__ import annotations

import html
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from auditautopilot import __version__  # noqa: E402

DEPS = ["openpyxl>=3.1"]

# 배치 파일은 ASCII 명령 + chcp 65001(UTF-8). 메시지는 한글.
PY_DETECT = r"""set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (where python >nul 2>nul && set "PY=python")
if not defined PY (
  echo [오류] 파이썬을 찾지 못했습니다. https://www.python.org/downloads/ 에서 3.10 이상을 설치하고
  echo        설치 첫 화면의 "Add python.exe to PATH"를 체크한 뒤 다시 실행하세요.
  pause
  exit /b 1
)
%PY% -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)"
if errorlevel 1 (
  echo [오류] 파이썬 3.10 이상이 필요합니다. 지금 버전:
  %PY% --version
  pause
  exit /b 1
)
"""

HEAD = """@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set "PYTHONPATH=%~dp0app"
set "PYTHONUTF8=1"
"""

BAT_RUN = HEAD + PY_DETECT + r"""echo 관리 화면을 엽니다. 이 창을 닫거나 Ctrl+C를 누르면 끝납니다.
echo 데이터: %~dp0data
%PY% -m auditautopilot serve --db "%~dp0data\audit.db" %*
echo.
echo 관리 화면이 끝났습니다.
pause
"""

BAT_CMD = HEAD + PY_DETECT + r"""doskey aap=%PY% -m auditautopilot $*
echo auditautopilot 명령창입니다. 예:
echo   aap --help
echo   aap materiality 재무제표등.xlsx --benchmark 매출액 --overall-pct 1 -o 중요성.xlsx
echo   aap workpaper prepare 조서.xlsx -o 조서_작업용.xlsx --assignee 홍길동
echo 파일은 이 창으로 끌어다 놓으면 경로가 들어갑니다.
echo.
cmd /k
"""

BAT_PIP = HEAD + PY_DETECT + r"""echo (선택) 이 PC 사용자 계정에 auditautopilot을 설치합니다. 인터넷 없이 wheels 폴더만 씁니다.
echo 설치하지 않아도 실행.bat·명령창.bat은 그대로 됩니다.
%PY% -m pip install --user --no-index --find-links "%~dp0wheels" --upgrade auditautopilot
if errorlevel 1 (
  echo [오류] 설치하지 못했습니다. 실행.bat을 그대로 쓰세요.
) else (
  echo 설치했습니다. 어느 폴더에서나 "python -m auditautopilot serve" 로 실행할 수 있습니다.
)
pause
"""


def run(cmd: list[str], **kw) -> None:
    print("$", " ".join(cmd))
    subprocess.run(cmd, check=True, **kw)


def md_to_html(md: str, title: str) -> str:
    """docs/manual.md 정도만 다루는 작은 변환기: 제목, 문단, 목록, 표, 코드 블록, **굵게**, `코드`, [링크](url)."""

    def inline(s: str) -> str:
        s = html.escape(s, quote=False)
        s = re.sub(r"`([^`]+)`", r"<code>\1</code>", s)
        s = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", s)
        s = re.sub(r"\[([^\]]+)\]\(([^)\s]+)\)", r'<a href="\2">\1</a>', s)
        return s

    out: list[str] = []
    lines = md.splitlines()
    i = 0
    toc: list[tuple[int, str, str]] = []
    while i < len(lines):
        ln = lines[i]
        if ln.startswith("```"):
            j = i + 1
            buf = []
            while j < len(lines) and not lines[j].startswith("```"):
                buf.append(lines[j])
                j += 1
            out.append("<pre>" + html.escape("\n".join(buf)) + "</pre>")
            i = j + 1
            continue
        m = re.match(r"(#{1,4}) (.*)", ln)
        if m:
            lvl = len(m.group(1))
            anchor = f"s{len(toc)}"
            if lvl == 2:
                toc.append((lvl, anchor, m.group(2)))
            out.append(f'<h{lvl} id="{anchor}">{inline(m.group(2))}</h{lvl}>')
            i += 1
            continue
        if ln.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{3,}:?", c) for c in cells):
                    rows.append(cells)
                i += 1
            t = ["<table>"]
            for k, r in enumerate(rows):
                tag = "th" if k == 0 else "td"
                t.append("<tr>" + "".join(f"<{tag}>{inline(c)}</{tag}>" for c in r) + "</tr>")
            t.append("</table>")
            out.append("".join(t))
            continue
        if re.match(r"\s*([-*]|\d+\.) ", ln):
            ordered = bool(re.match(r"\s*\d+\. ", ln))
            tag = "ol" if ordered else "ul"
            items = []
            while i < len(lines) and re.match(r"\s*([-*]|\d+\.) ", lines[i]):
                text = re.sub(r"^\s*([-*]|\d+\.) ", "", lines[i])
                i += 1
                while i < len(lines) and lines[i].startswith("   ") and lines[i].strip():
                    text += " " + lines[i].strip()
                    i += 1
                items.append(f"<li>{inline(text)}</li>")
            out.append(f"<{tag}>" + "".join(items) + f"</{tag}>")
            continue
        if ln.startswith("> "):
            out.append(f'<p class="note">{inline(ln[2:])}</p>')
            i += 1
            continue
        if not ln.strip():
            i += 1
            continue
        buf = [ln]
        i += 1
        while i < len(lines) and lines[i].strip() and not re.match(r"(#|\||```|\s*([-*]|\d+\.) |> )", lines[i]):
            buf.append(lines[i])
            i += 1
        out.append("<p>" + inline(" ".join(buf)) + "</p>")
    nav = "".join(f'<li><a href="#{a}">{html.escape(t)}</a></li>' for _, a, t in toc)
    return f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>{html.escape(title)}</title>
<style>
:root{{--bg:#F7F4EB;--fg:#1d1d1b;--mute:#6b675e;--line:#ddd6c4;--accent:#0D1F08;--surface:#fff}}
body{{margin:0;background:var(--bg);color:var(--fg);font:15px/1.65 "Malgun Gothic","Apple SD Gothic Neo",sans-serif}}
main{{max-width:960px;margin:0 auto;padding:24px 16px 80px}}
h1{{font-size:26px;margin:8px 0 4px}}h2{{font-size:20px;margin:36px 0 8px;padding-top:12px;border-top:2px solid var(--accent)}}
h3{{font-size:16px;margin:22px 0 6px}}
table{{border-collapse:collapse;width:100%;background:var(--surface);margin:8px 0 14px;font-size:14px}}
th,td{{border:1px solid var(--line);padding:6px 9px;text-align:left;vertical-align:top}}th{{background:#ece6d6}}
code{{background:#ece6d6;padding:1px 5px;border-radius:4px;font-size:13px}}
pre{{background:#1f2a1c;color:#f3f0e6;padding:12px 14px;border-radius:8px;overflow-x:auto;font-size:13px}}
.note{{background:#fff8e0;border-left:4px solid #d4a72c;padding:8px 12px}}
nav ul{{columns:2;font-size:14px}}td:first-child{{white-space:nowrap}}a{{color:#31110F}}
</style></head><body><main><nav><ul>{nav}</ul></nav>{''.join(out)}</main></body></html>"""


def build() -> Path:
    name = f"auditautopilot-{__version__}-windows"
    dist = ROOT / "dist"
    dist.mkdir(exist_ok=True)
    shutil.rmtree(ROOT / "build", ignore_errors=True)  # 지난 빌드 찌꺼기가 wheel에 섞이지 않게
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        stage = tmp / name
        wheels = stage / "wheels"
        app = stage / "app"
        wheels.mkdir(parents=True)
        app.mkdir()

        # 1) wheel: 이 저장소 + 의존성(순수 파이썬 wheel만)
        run([sys.executable, "-m", "pip", "wheel", str(ROOT), "--no-deps", "-w", str(wheels), "-q"])
        run([sys.executable, "-m", "pip", "download", *DEPS, "--only-binary=:all:", "--platform", "any",
             "--python-version", "3.10", "--implementation", "py", "-d", str(wheels), "-q"])
        for whl in wheels.glob("*.whl"):
            if not whl.name.endswith("-none-any.whl"):
                raise SystemExit(f"순수 파이썬이 아닌 wheel이 섞였습니다: {whl.name}")

        # 2) app/: wheel을 풀어 둔다 (pip 없이 PYTHONPATH로 실행)
        for whl in wheels.glob("*.whl"):
            with zipfile.ZipFile(whl) as z:
                for member in z.namelist():
                    if ".dist-info/" in member:
                        continue
                    z.extract(member, app)

        # 3) 배치 파일 (CRLF)
        for fname, body in (("실행.bat", BAT_RUN), ("명령창.bat", BAT_CMD), ("pip설치.bat", BAT_PIP)):
            (stage / fname).write_bytes(body.replace("\r\n", "\n").replace("\n", "\r\n").encode("utf-8"))

        # 4) 설명서
        md = (ROOT / "docs" / "manual.md").read_text(encoding="utf-8")
        (stage / "사용설명서.html").write_text(md_to_html(md, f"auditautopilot {__version__} 사용설명서"),
                                            encoding="utf-8")
        (stage / "사용설명서.md").write_text(md, encoding="utf-8")
        (stage / "data").mkdir()
        (stage / "data" / "여기에_데이터가_저장됩니다.txt").write_text(
            "audit.db(감사 데이터)와 감사 폴더(받은 자료·조서)가 이 폴더에 생깁니다.\r\n"
            "새 버전으로 바꿀 때는 이 data 폴더를 통째로 새 폴더에 옮기세요.\r\n", encoding="utf-8")

        out = dist / f"{name}.zip"
        if out.exists():
            out.unlink()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for p in sorted(stage.rglob("*")):
                if "__pycache__" in p.parts:
                    continue
                z.write(p, p.relative_to(tmp))
    print(f"저장: {out}")
    return out


if __name__ == "__main__":
    build()
