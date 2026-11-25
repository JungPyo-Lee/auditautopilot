"""명령줄 실행 진입점.

예)
  python -m auditautopilot workpaper prepare 조서.xlsx -o 조서_작업용.xlsx --assignee 이정표
  python -m auditautopilot workpaper print-copy 조서.xlsx -o 조서_인쇄용.xlsx
  python -m auditautopilot import semusarang 재무제표등.xlsx -o 정리.xlsx
  python -m auditautopilot import douzone-journal 분개장.xlsx -o 정리.xlsx
  python -m auditautopilot materiality --benchmark 매출액 --amount 18669084279 --overall-pct 1
  python -m auditautopilot analytics 재무제표등.xlsx -o 분석적검토.xlsx --threshold-amount 300000000
  python -m auditautopilot sample 재무제표등.xlsx --account 108 --method mus --n 25 -o 표본.xlsx
"""

from __future__ import annotations

import argparse
import re
import sys

from . import analytics, materiality, sampling
from .importers import douzone, semusarang
from .reports import write_excel
from .workpaper import make_print_copy, prepare_workpaper

TB_HEADERS = {
    "code": "계정코드", "name": "계정과목", "level": "구분",
    "debit_balance": "차변잔액", "debit_total": "차변합계",
    "credit_total": "대변합계", "credit_balance": "대변잔액", "balance": "잔액(차-대)",
}
FS_HEADERS = {"code": "계정코드", "name": "계정과목", "is_account": "계정여부", "current": "당기", "prior": "전기"}
JOURNAL_HEADERS = {
    "date": "일자", "voucher_no": "전표번호", "side": "구분", "account_code": "계정코드",
    "account_name": "계정과목", "debit": "차변", "credit": "대변",
    "partner_code": "거래처코드", "partner": "거래처", "memo": "적요",
}
def _rename(rows: list[dict], mapping: dict[str, str], keep: list[str] | None = None) -> list[dict]:
    keys = keep
    out = []
    for r in rows:
        ks = keys or list(r)
        out.append({mapping.get(k, k): r.get(k) for k in ks})
    return out


PARTNER_HEADERS = {
    "account_code": "계정코드", "partner_code": "거래처코드", "partner": "거래처",
    "reg_no": "등록번호", "opening": "전기이월", "debit": "차변", "credit": "대변", "balance": "잔액",
}


def _months_from_period(label: str) -> int | None:
    """'제 14(당)기 2026년1월1일 ~ 2026년9월30일' -> 9"""
    m = re.findall(r"(\d{4})년\s*(\d{1,2})월", label or "")
    if len(m) >= 2:
        (y1, m1), (y2, m2) = m[0], m[1]
        return (int(y2) - int(y1)) * 12 + int(m2) - int(m1) + 1
    return None


def cmd_workpaper(args) -> int:
    if args.action == "prepare":
        report = prepare_workpaper(args.src, args.output, assignee=args.assignee,
                                   hide_procedures=not args.keep_procedures,
                                   add_index=not args.no_index,
                                   aggressive_names=args.aggressive_names)
    else:
        report = make_print_copy(args.src, args.output, assignee=args.assignee)
    print("\n".join(report.lines()))
    print(f"저장: {args.output}")
    return 0


def cmd_import(args) -> int:
    if args.kind == "semusarang":
        b = semusarang.load(args.src)
        sheets = {
            "시산표": _rename(b.tb, TB_HEADERS),
            "재무상태표": _rename(b.bs, FS_HEADERS),
            "손익계산서": _rename(b.pl, FS_HEADERS),
            "분개장": _rename(b.journal, JOURNAL_HEADERS),
            "거래처별잔액": _rename(b.partners, PARTNER_HEADERS),
        }
        print(f"시산표 {len(b.tb)}행, 분개장 {len(b.journal):,}행, 거래처 잔액 {len(b.partners):,}행")
        if b.periods:
            print(f"기간: {b.periods.get('current', '')}")
    else:
        j = douzone.load_journal(args.src)
        sheets = {"분개장": _rename(j, JOURNAL_HEADERS)}
        diff = sum(r["debit"] for r in j) - sum(r["credit"] for r in j)
        print(f"분개장 {len(j):,}행, 차대 차이 {diff:,.0f}")
    write_excel(args.output, sheets)
    print(f"저장: {args.output}")
    return 0


def cmd_materiality(args) -> int:
    amount = args.amount
    if amount is None:
        if not args.src:
            print("--amount 또는 세무사랑 파일 경로가 필요합니다.", file=sys.stderr)
            return 2
        b = semusarang.load(args.src)
        fs = b.pl if args.benchmark in ("세전이익", "매출액") else b.bs
        amount = materiality.find_benchmark_amount(fs, args.benchmark)
        if amount is None:
            print(f"재무제표에서 '{args.benchmark}' 금액을 찾지 못했습니다. --amount 로 직접 입력하세요.", file=sys.stderr)
            return 2
        if args.months is None and args.benchmark in ("세전이익", "매출액"):
            args.months = _months_from_period(b.periods.get("current", "")) or 12
    result = materiality.compute(
        args.benchmark, amount,
        overall_pct=args.overall_pct or materiality.DEFAULT_PERCENTAGES.get(args.benchmark, 1.0),
        performance_pct=args.performance_pct, trivial_pct=args.trivial_pct,
        months=args.months or 12, rounding_unit=args.rounding,
    )
    if amount < 0:
        print(f"주의: 기준 금액이 음수({amount:,.0f})입니다. 절대값으로 계산했으니 기준 선택을 다시 검토하세요.")
    table = result.as_rows()
    for r in table:
        v = r["값"]
        print(f"{r['항목']}: {v:,.0f}" if isinstance(v, float) else f"{r['항목']}: {v}")
    if args.output:
        write_excel(args.output, {"중요성": table})
        print(f"저장: {args.output}")
    return 0


def cmd_analytics(args) -> int:
    b = semusarang.load(args.src)
    months = args.months or _months_from_period(b.periods.get("current", "")) or 12
    rule = analytics.VarianceRule(args.threshold_amount, args.threshold_pct, annualize_factor=12 / months)
    bs = analytics.compare_periods(b.bs, rule)
    pl = analytics.compare_periods(b.pl, rule, annualize=months < 12)
    questions = analytics.draft_questions(bs + pl)
    cols = {"code": "계정코드", "name": "계정과목", "prior": "전기", "current": "당기",
            "current_compared": "비교금액(연환산)", "diff": "증감", "pct": "증감률(%)",
            "flag": "검토대상", "reason": "사유"}
    keep = list(cols)
    sheets = {
        "재무상태표 증감": _rename(bs, cols, keep),
        "손익계산서 증감": _rename(pl, cols, keep),
        "질문서 초안": questions,
    }
    if b.journal:
        sheets["월별 추이"] = _rename(analytics.monthly_trend(b.journal),
                                   {"account_code": "계정코드", "account_name": "계정과목"})
    write_excel(args.output, sheets)
    print(f"기준: {args.threshold_amount:,.0f}원 이상 그리고 {args.threshold_pct}% 이상, 손익 연환산 {months}개월→12개월")
    print(f"검토대상 재무상태표 {sum(r['flag'] for r in bs)}개, 손익계산서 {sum(r['flag'] for r in pl)}개, 질문 {len(questions)}건")
    print(f"저장: {args.output}")
    return 0


def cmd_sample(args) -> int:
    b = semusarang.load(args.src)
    code = str(int(args.account)).zfill(3) if args.account.isdigit() else args.account
    if args.source == "partners":
        pop = _rename([r for r in b.partners if r["account_code"] == code], PARTNER_HEADERS)
        amount_col = "잔액"
    else:
        pop = [{**r, "금액": r["debit"] - r["credit"]} for r in b.journal if r["account_code"] == code]
        pop = _rename(pop, JOURNAL_HEADERS)
        amount_col = "금액"
    if not pop:
        print(f"계정 {args.account}의 모집단이 비어 있습니다.", file=sys.stderr)
        return 2
    result = sampling.sample(pop, amount_col, args.method, n=args.n,
                             threshold=args.threshold, seed=args.seed)
    write_excel(args.output, {"표본": result.selected, "추출근거": result.summary(), "모집단": pop})
    print(f"모집단 {result.population_count:,}건 중 {len(result.selected)}건 추출 (방법 {result.method}, 시드 {result.seed})")
    print(f"저장: {args.output}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="auditautopilot", description="회계감사 업무 자동화 도구")
    sub = p.add_subparsers(dest="command", required=True)

    wp = sub.add_parser("workpaper", help="입증감사조서 정리")
    wp.add_argument("action", choices=["prepare", "print-copy"])
    wp.add_argument("src")
    wp.add_argument("-o", "--output", required=True)
    wp.add_argument("--assignee", help="이 담당자의 인덱스 탭만 남긴다 (총괄 시트 기준)")
    wp.add_argument("--keep-procedures", action="store_true", help="절차서 탭을 숨기지 않는다")
    wp.add_argument("--no-index", action="store_true", help="목차 시트를 만들지 않는다")
    wp.add_argument("--aggressive-names", action="store_true",
                    help="쓰이지 않는 숨김 이름 정의도 지운다")
    wp.set_defaults(func=cmd_workpaper)

    im = sub.add_parser("import", help="회계 프로그램 내보내기 파일을 표준 형식으로 정리")
    im.add_argument("kind", choices=["semusarang", "douzone-journal"])
    im.add_argument("src")
    im.add_argument("-o", "--output", required=True)
    im.set_defaults(func=cmd_import)

    ma = sub.add_parser("materiality", help="중요성 계산")
    ma.add_argument("src", nargs="?", help="세무사랑 재무제표 파일 (금액 자동 찾기)")
    ma.add_argument("--benchmark", required=True, choices=list(materiality.DEFAULT_PERCENTAGES))
    ma.add_argument("--amount", type=float)
    ma.add_argument("--overall-pct", type=float)
    ma.add_argument("--performance-pct", type=float, default=materiality.DEFAULT_PERFORMANCE_PCT)
    ma.add_argument("--trivial-pct", type=float, default=materiality.DEFAULT_TRIVIAL_PCT)
    ma.add_argument("--months", type=int, help="기준 금액의 기간(개월). 12 미만이면 연환산")
    ma.add_argument("--rounding", type=int, default=1000, help="내림 단위 (기본 1,000원)")
    ma.add_argument("-o", "--output")
    ma.set_defaults(func=cmd_materiality)

    an = sub.add_parser("analytics", help="분석적 검토 (세무사랑 파일)")
    an.add_argument("src")
    an.add_argument("-o", "--output", required=True)
    an.add_argument("--threshold-amount", type=float, required=True, help="보통 수행중요성")
    an.add_argument("--threshold-pct", type=float, default=10.0)
    an.add_argument("--months", type=int, help="당기 기간(개월). 생략하면 파일 머리글에서 읽는다")
    an.set_defaults(func=cmd_analytics)

    sa = sub.add_parser("sample", help="표본추출 (세무사랑 파일)")
    sa.add_argument("src")
    sa.add_argument("--account", required=True, help="계정코드 (예: 108)")
    sa.add_argument("--source", choices=["partners", "journal"], default="partners",
                    help="partners=거래처별 잔액, journal=분개장 전표")
    sa.add_argument("--method", choices=list(sampling.METHODS), default="mus")
    sa.add_argument("--n", type=int, default=25)
    sa.add_argument("--threshold", type=float, help="이 금액 이상은 전수 추출")
    sa.add_argument("--seed", type=int)
    sa.add_argument("-o", "--output", required=True)
    sa.set_defaults(func=cmd_sample)
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except (ValueError, KeyError, FileNotFoundError) as exc:
        print(f"오류: {exc}", file=sys.stderr)
        return 1
