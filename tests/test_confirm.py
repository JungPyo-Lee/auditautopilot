"""채권채무 전자조회 회신현황 정리·대사. 가짜 거래처·숫자만 쓴다."""

from datetime import datetime

import openpyxl

from auditautopilot import confirm
from auditautopilot.cli import main

from test_douzone import _save, dz  # noqa: F401  (픽스처 재사용)

HEAD = ["기업명", "조회차수", "사업자등록번호", "담당자명", "담당자이메일", "일치여부", "통화", "구분",
        "발송금액", "회신금액", "차이금액", "최초발송일시", "회신일시"]
SENT = datetime(2026, 11, 20, 15, 30, 33, 53000)
REPLY = datetime(2026, 11, 24, 9, 0, 0)


def _status(tmp_path):
    return _save(tmp_path, "채권채무_전자조회_회신현황.xlsx", [
        HEAD,
        ["가나상사", 1, "1112233333", "김가나", "a@ganada.test", "Y", "KRW", "외상매출금(커머스)", 800000, 800000, "0.00", SENT, REPLY],
        ["（주）다라상사", 1, "2223344444", "이다라", "b@dara.test", "N", "KRW", "외상매출금", 250000, 240000, "-10000.00", SENT, REPLY],
        ["사아물산", 1, "3334455555", "박사아", "c@saa.test", "", "KRW", "외상매입금", 250000, None, None, SENT, None],
        ["FOREIGN CO., LTD", 1, "4445566666", " ", "d@foreign.test", "N", "USD", "accounts receivable", 1000.5, 1400000, "+1398999.5", SENT, REPLY],
        ["가나상사", 1, "1112233333", "김가나", "a@ganada.test", "Y", "KRW", "미수금", 5000, 5000, "0.00", SENT, REPLY],
    ])


def test_read_status(tmp_path):
    rows = {(r["partner"], r["account"]): r for r in confirm.load_status(_status(tmp_path))}
    assert rows[("가나상사", "외상매출금")]["status"] == "일치"
    dara = rows[("（주）다라상사", "외상매출금")]
    assert dara["status"] == "불일치" and dara["diff"] == -10000
    saa = rows[("사아물산", "외상매입금")]
    assert saa["status"] == "미회신" and saa["replied"] is None and saa["diff"] is None
    fx = rows[("FOREIGN CO., LTD", "외상매출금")]
    assert fx["currency"] == "USD" and "단위" in fx["note"]
    assert rows[("가나상사", "외상매출금")]["sent_at"].microsecond == 0


def test_account_and_name_normalize():
    assert confirm.account_of("외상매입금(커머스)") == "외상매입금"
    assert confirm.account_of("accounts payable") == "외상매입금"
    assert confirm.account_of("Other receivables") == "미수금"
    assert confirm.norm_name("（주）다라상사") == confirm.norm_name("다라상사") == confirm.norm_name("주식회사 다라상사")


def test_latest_round_wins(tmp_path):
    p = _save(tmp_path, "s.xlsx", [
        HEAD,
        ["가나상사", 1, "1112233333", "", "", "", "KRW", "외상매출금", 800000, None, None, SENT, None],
        ["가나상사", 2, "1112233333", "", "", "Y", "KRW", "외상매출금", 800000, 800000, "0", SENT, REPLY],
    ])
    [r] = confirm.load_status(p)
    assert r["round"] == 2 and r["status"] == "일치"


def test_reconcile_with_books(tmp_path, dz):  # noqa: F811
    res = confirm.build([_status(tmp_path), dz["bs"], dz["ar"], dz["ap"]])
    by = {(r["partner"], r["account"]): r for r in res.rows}
    assert by[("가나상사", "외상매출금")]["book"] == 800000 and by[("가나상사", "외상매출금")]["book_diff"] == 0
    assert by[("（주）다라상사", "외상매출금")]["book"] == 250000        # (주) 표기가 달라도 같은 거래처
    assert by[("사아물산", "외상매입금")]["book"] == 250000
    assert "book" not in by[("FOREIGN CO., LTD", "외상매출금")]           # 외화는 장부 대사 제외
    s = {x["계정"]: x for x in res.summary}
    ar = s["외상매출금"]
    assert (ar["조회 건수"], ar["일치"], ar["불일치"], ar["미회신"], ar["외화 건수"]) == (3, 1, 2, 0, 1)
    assert ar["발송금액(원화)"] == 1050000 and ar["재무상태표 잔액"] == 1000000 and ar["조회 비율(원화)"] == 105.0
    assert s["외상매입금"]["미회신 금액(원화)"] == 250000
    assert any("외화" in n for n in res.notes)


def test_cli_confirm(tmp_path, dz, capsys):  # noqa: F811
    out = tmp_path / "조회.xlsx"
    assert main(["confirm", str(_status(tmp_path)), str(dz["bs"]), str(dz["ar"]), "-o", str(out)]) == 0
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ["요약", "불일치_차이조정", "미회신_대체절차", "명세", "참고"]
    diff = list(wb["불일치_차이조정"].iter_rows(values_only=True))
    assert "차이 원인" in diff[0] and len(diff) == 3
    assert "외상매출금: 3건" in capsys.readouterr().out


def test_not_a_status_file(tmp_path, dz):  # noqa: F811
    assert main(["confirm", str(dz["bs"]), "-o", str(tmp_path / "x.xlsx")]) == 1
