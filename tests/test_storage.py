"""데이터를 프로그램 바깥 고정 폴더에 두고, 예전 설치본을 옮기는지 확인한다."""
from pathlib import Path

from auditautopilot.manage import config, storage
from auditautopilot.manage.db import Store


def test_data_home_follows_env(monkeypatch, tmp_path):
    monkeypatch.setenv("AUDITAUTOPILOT_DATA", str(tmp_path / "보관"))
    assert storage.data_home() == tmp_path / "보관"
    assert storage.default_db() == tmp_path / "보관" / "audit.db"
    assert storage.default_root() == tmp_path / "보관" / "감사"


def test_fresh_and_legacy_root_become_absolute(monkeypatch, tmp_path):
    monkeypatch.setenv("AUDITAUTOPILOT_DATA", str(tmp_path / "보관"))
    s = Store(tmp_path / "audit.db")
    assert config.load(s)["root"] == str(storage.default_root())     # 새 DB
    s.set_setting("root", "data/감사")                                # 예전 상대 경로
    assert config.load(s)["root"] == str(storage.default_root())
    assert s.get_setting("root") == str(storage.default_root())      # 저장도 됨
    s.set_setting("root", str(tmp_path / "내폴더"))                   # 사용자가 정한 경로는 그대로
    assert config.load(s)["root"] == str(tmp_path / "내폴더")


def test_migrates_old_program_folder_data(monkeypatch, tmp_path):
    monkeypatch.setenv("AUDITAUTOPILOT_DATA", str(tmp_path / "home"))
    prog = tmp_path / "prog" / "data"
    (prog / "감사" / "01.가나").mkdir(parents=True)
    (prog / "audit.db").write_bytes(b"OLD")
    (prog / "감사" / "01.가나" / "a.xlsx").write_text("x")
    moved = storage.migrate_old_install(prog)
    assert moved == str(tmp_path / "home")
    assert (tmp_path / "home" / "audit.db").read_bytes() == b"OLD"
    assert (tmp_path / "home" / "감사" / "01.가나" / "a.xlsx").exists()
    assert storage.migrate_old_install(prog) is None                 # 이미 옮겼으면 다시 안 함
