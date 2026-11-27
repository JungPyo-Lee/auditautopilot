"""업데이트.bat 가 부른다. 같은 폴더에 둔 새 배포 zip을 찾아 프로그램 파일만 교체한다.

- 업무 데이터(audit.db, 감사 폴더)는 이 폴더 바깥(%LOCALAPPDATA%\\auditautopilot)에 있으므로 건드리지 않는다.
- app 폴더와 .bat·설명서·wheels·버전.txt·update.py 만 새 것으로 바꾼다.
- 새 zip은 같은 폴더에 두면 된다. 여러 개면 가장 최근 것을 쓴다.
- 안전하게: 바꾸기 전에 지금 app 폴더를 app_백업 으로 옮겨 두고, 다 되면 지운다. 중간에 실패하면 되돌린다.
"""

from __future__ import annotations

import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPLACE = {"app", "wheels", "실행.bat", "명령창.bat", "업데이트.bat", "pip설치.bat",
           "사용설명서.html", "사용설명서.md", "버전.txt", "update.py"}


def find_zip() -> Path | None:
    zips = [p for p in HERE.glob("*.zip") if p.is_file()]
    if not zips:
        return None
    return max(zips, key=lambda p: p.stat().st_mtime)


def main() -> int:
    z = find_zip()
    if z is None:
        print("새 zip 파일을 찾지 못했습니다. 받은 'auditautopilot-...-windows.zip'을 이 폴더에 둔 뒤 다시 실행하세요.")
        return 1
    print(f"새 버전 파일: {z.name}")
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        with zipfile.ZipFile(z) as zf:
            zf.extractall(tmp)
        tops = [p for p in tmp.iterdir() if p.is_dir()]
        src = tops[0] if len(tops) == 1 else tmp        # zip 안 최상위 폴더(auditautopilot-x.y.z-windows)
        if not (src / "app").exists():
            print("이 zip은 auditautopilot 배포 파일이 아닌 것 같습니다 (app 폴더가 없습니다).")
            return 1

        backup = HERE / "app_백업"
        if (HERE / "app").exists():
            if backup.exists():
                shutil.rmtree(backup, ignore_errors=True)
            shutil.move(str(HERE / "app"), str(backup))
        try:
            for name in REPLACE:
                s = src / name
                if not s.exists():
                    continue
                d = HERE / name
                if d.exists():
                    if d.is_dir():
                        shutil.rmtree(d, ignore_errors=True)
                    else:
                        d.unlink()
                shutil.move(str(s), str(d))
        except Exception as exc:                        # 실패하면 되돌린다
            print(f"업데이트 중 문제가 생겨 되돌립니다: {exc}")
            if backup.exists():
                if (HERE / "app").exists():
                    shutil.rmtree(HERE / "app", ignore_errors=True)
                shutil.move(str(backup), str(HERE / "app"))
            return 1
        shutil.rmtree(backup, ignore_errors=True)

    ver = (HERE / "버전.txt").read_text(encoding="utf-8").strip() if (HERE / "버전.txt").exists() else ""
    print(f"업데이트했습니다. {ver}")
    print("이제 받은 zip 파일은 지워도 됩니다. 실행.bat 으로 열어 확인하세요. 데이터는 그대로입니다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
