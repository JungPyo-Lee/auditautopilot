"""데이터를 어디에 둘지 정한다.

업무 데이터(audit.db, 받은 자료·조서가 쌓이는 '감사' 폴더)는 프로그램 폴더 **바깥**의
고정된 사용자 폴더에 둔다. 그래야 새 버전으로 프로그램 폴더를 통째로 덮어써도 데이터가 지워지지 않는다.

- 윈도우: %LOCALAPPDATA%\\auditautopilot   (보통 C:\\Users\\<사용자>\\AppData\\Local\\auditautopilot)
- macOS: ~/Library/Application Support/auditautopilot
- 그 밖: ~/.local/share/auditautopilot  (XDG_DATA_HOME 따름)
- 환경변수 AUDITAUTOPILOT_DATA 를 주면 그 경로를 그대로 쓴다.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

APP = "auditautopilot"


def data_home() -> Path:
    env = os.environ.get("AUDITAUTOPILOT_DATA")
    if env:
        return Path(env).expanduser()
    if sys.platform.startswith("win"):
        base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
        return Path(base) / APP
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / APP
    base = os.environ.get("XDG_DATA_HOME") or str(Path.home() / ".local" / "share")
    return Path(base) / APP


def default_db() -> Path:
    return data_home() / "audit.db"


def default_root() -> Path:
    """받은 자료·조서가 쌓이는 '감사' 폴더의 기본 위치."""
    return data_home() / "감사"


# 예전 버전(0.2.x 이하)이 쓰던 프로그램 폴더 안 상대 경로
LEGACY_ROOT = "data/감사"


def migrate_old_install(program_data: Path) -> str | None:
    """예전 버전은 프로그램 폴더 안 data\\ 에 저장했다.
    새 위치가 비어 있고 옛 data 폴더에 audit.db 가 있으면 통째로 옮긴다. 옮긴 위치를 돌려준다(없으면 None)."""
    home = data_home()
    if (home / "audit.db").exists():
        return None
    old_db = program_data / "audit.db"
    if not old_db.exists():
        return None
    home.mkdir(parents=True, exist_ok=True)
    for item in sorted(program_data.iterdir()):
        dest = home / item.name
        if dest.exists():
            continue
        try:
            shutil.move(str(item), str(dest))
        except OSError:
            # 옮기지 못하면 복사라도 (원본은 그대로 둔다)
            if item.is_dir():
                shutil.copytree(item, dest, dirs_exist_ok=True)
            else:
                shutil.copy2(item, dest)
    return str(home)
