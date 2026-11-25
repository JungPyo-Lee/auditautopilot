"""관리 화면: 내려받는 파일 이름, 파일 오류 처리, 질문 모두 선택. 가짜 회사만 쓴다."""

import re
import urllib.parse
from datetime import datetime

from auditautopilot.manage import web
from auditautopilot.workpaper import adjust

from test_adjust import AJE_ROWS, _aje, bs_prior, filled  # noqa: F401
from test_douzone import dz  # noqa: F401
from test_leadsheet import tb_file, template  # noqa: F401
from test_manage import _multipart, _open, server, store  # noqa: F401


def test_stamped():
    now = datetime(2026, 11, 25, 21, 30, 5)
    assert web.stamped("가나_수정분개반영.xlsx", now) == "가나_수정분개반영_20261125_213005.xlsx"
    assert web.stamped("감사보고서.docx", now) == "감사보고서_20261125_213005.docx"
    assert web.stamped("질문서.eml", now) == "질문서.eml"


def test_download_names_and_file_errors(filled, tmp_path, server, store, monkeypatch):  # noqa: F811
    base, _ = server
    cid = store.add_client("C014", "가나상사", "김담당", "kim@example.com")
    eid = store.add_engagement(cid, 2026, "기말감사", "2027-01-11", "2027-01-15")
    store.set_setting(f"leadsheet:{eid}", {"made_at": "x", "path": str(filled), "lines": []})
    aje = _aje(tmp_path, filled, AJE_ROWS)
    data, hdr = _multipart([], [("aje", "수정분개.xlsx", aje.read_bytes())])
    assert _open(base + f"/e/{eid}/adjust", data, hdr)[0] == 303
    names = []
    for _ in range(2):                                   # 두 번 받아도 이름이 겹치지 않게 시각이 붙는다
        st, h, _ = _open(base + f"/e/{eid}/adjusted.xlsx")
        assert st == 200
        names.append(urllib.parse.unquote(h["Content-Disposition"].split("''", 1)[1]))
    assert all(re.search(r"_수정분개반영_\d{8}_\d{6}\.xlsx$", n) for n in names)
    # 다시 반영해도(같은 이름 파일이 있어도) 새 이름으로 저장된다
    data, hdr = _multipart([], [("aje", "수정분개.xlsx", aje.read_bytes())])
    assert _open(base + f"/e/{eid}/adjust", data, hdr)[0] == 303
    assert store.get_setting(f"adjust:{eid}")["path"].endswith("_v2.xlsx")

    def locked(*a, **k):
        raise PermissionError(13, "Permission denied")
    monkeypatch.setattr(adjust, "apply", locked)
    data, hdr = _multipart([], [("aje", "수정분개.xlsx", aje.read_bytes())])
    st, h, _ = _open(base + f"/e/{eid}/adjust", data, hdr)
    loc = urllib.parse.unquote(h["Location"])
    assert st == 303 and "엑셀에서 열려 있으면 닫고" in loc


def test_question_select_all_buttons(server, store):  # noqa: F811
    base, _ = server
    cid = store.add_client("C015", "가나상사", "김담당", "kim@example.com")
    eid = store.add_engagement(cid, 2026, "기말감사", "2027-01-11", "2027-01-15")
    page = _open(base + f"/e/{eid}")[2].decode()
    assert "모두 선택" in page and "모두 해제" in page and "input[name=ids][form=qact]" in page
