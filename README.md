# auditautopilot

회계감사 중간·기말 업무 자동화 도구. 설계서: [감사업무 자동화 시스템 설계서](https://claude.ai/artifact/14tNiAD9BqWFPDhNwmcY1A)

지금 들어 있는 기능 (2단계: 핵심 계산)

| 기능 | 명령 | 설명 |
| --- | --- | --- |
| 조서 정리 | `workpaper prepare` | 절차서 탭 숨김, 담당자별 탭만 보기, 맨 앞 목차 시트, 깨진 이름 정의(#REF!) 정리 |
| 인쇄용 사본 | `workpaper print-copy` | 절차서 탭까지 모두 보이는 사본 |
| 데이터 불러오기 | `import semusarang`, `import douzone-journal` | 내보내기 엑셀을 표준 형식으로 정리 (주민등록번호 뒷자리 가림) |
| 중요성 | `materiality` | 전체·수행중요성, 명백히 사소한 금액. 중간감사는 연환산 |
| 분석적 검토 | `analytics` | 전기 대비 증감, 기준 초과 계정 표시, 질문서 초안, 월별 추이 |
| 표본추출 | `sample` | 기준금액 이상 전수, 무작위, 체계적, 금액단위(MUS). 시드값 기록 |

관리 페이지, 메일 자동화, 첨부 자동 저장은 3단계에서 붙인다.

## 설치

Python 3.10 이상([python.org](https://www.python.org/downloads/) 설치본)이 필요하다.
필요한 패키지는 openpyxl 하나뿐이다.

```
python -m pip install -e ".[dev]"
```

윈도우 스마트 앱 컨트롤·회사 앱 제어 정책은 서명 없는 DLL(pandas, numpy, lxml 등)과
가상환경의 `pip.exe`를 막는다. 그래서 이 도구는 순수 파이썬 패키지만 쓰고,
설치는 `pip` 대신 `python -m pip`으로 한다. 새 기능에도 컴파일된 패키지를 추가하지 않는다
(`tests/test_cli.py::test_no_compiled_packages_are_imported`가 확인한다).

## 사용 예

```
# 조서: 내 담당 인덱스만 보이게 + 목차 + 깨진 이름 정리 (원본은 그대로 두고 새 파일로 저장)
python -m auditautopilot workpaper prepare 입증감사조서.xlsx -o 조서_작업용.xlsx --assignee 홍길동

# 인쇄용 사본 (절차서 탭 포함 전체 표시)
python -m auditautopilot workpaper print-copy 입증감사조서.xlsx -o 조서_인쇄용.xlsx

# 세무사랑 '재무제표 등' 파일 정리
python -m auditautopilot import semusarang 재무제표등.xlsx -o 정리.xlsx

# 중요성: 파일에서 매출액을 찾아 9개월이면 자동 연환산
python -m auditautopilot materiality 재무제표등.xlsx --benchmark 매출액 --overall-pct 1 -o 중요성.xlsx

# 분석적 검토: 수행중요성 이상이면서 10% 이상 변동한 계정
python -m auditautopilot analytics 재무제표등.xlsx -o 분석적검토.xlsx --threshold-amount 858395000

# 더존: 재무상태표·손익계산서·거래처원장·분개장을 함께 주면 종류를 알아서 맞춘다
python -m auditautopilot import douzone 재무상태표과목별.xlsx 손익계산서과목별.xlsx "거래처원장 (1).xlsx" -o 정리.xlsx
python -m auditautopilot analytics 재무상태표과목별.xlsx 손익계산서과목별.xlsx -o 분석적검토.xlsx --threshold-amount 266196000
python -m auditautopilot sample 재무상태표과목별.xlsx "거래처원장 (1).xlsx" --account 외상매출금 --threshold 500000000 -o 표본.xlsx

# 표본: 외상매출금(108) 거래처 잔액에서 10억 이상 전수 + 나머지 MUS 25건
python -m auditautopilot sample 재무제표등.xlsx --account 108 --threshold 1000000000 --method mus --n 25 -o 표본.xlsx
```

## 지원하는 파일 형식

- **세무사랑 '재무제표 등' 묶음**: TB, BS, PL, 분개장, 계정코드 이름의 거래처원장 시트
- **더존 분개장**: 전표일자 | 전표번호 | 구분 | Code | 계정과목 | 차변 | 대변 | 거래처 | Code | 적요
- **더존 재무상태표·손익계산서(과목별)**: 계정코드가 없어 과목명으로 구분한다. 차감계정은
  `대손충당금(외상매출금)`, `감가상각누계액(건물)`처럼 앞 계정 이름을 붙인다. 분개장을 함께 주면 계정코드를 채운다.
- **더존 거래처원장(잔액)**: 거래처 | 전기(월)이월 | 차변 | 대변 | 잔액. 파일에 계정이 없으므로
  합계 잔액을 재무상태표 계정 금액과 맞춰 어느 계정인지 찾는다(같은 금액이 여럿이면 전기이월까지 비교).
- 더존 파일은 여러 개를 한 번에 주면 종류를 알아서 맞춘다. 분개장이 재무제표 기간보다 길면 기간 끝까지만 쓴다.
- 더존 합계잔액시산표는 샘플을 받으면 추가한다.

## 조서 정리 방식

조서 파일은 openpyxl로 다시 저장하면 도형(화살표 등)이 사라지기 때문에, 엑셀 내부 XML 중
시트 목록·이름 정의·목차 시트만 고치고 나머지는 그대로 복사한다.

- **절차서 탭**: 숫자가 없는 탭(A, B, DER, BB,DD)은 숨긴다. 일반사항·총괄·공시점검·TB·BS·IS는 항상 보인다.
- **담당자 보기**: 총괄 시트의 `인덱스 | 계정과목 | 담당자` 표를 읽어, 다른 사람 인덱스의 탭을 숨긴다.
- **이름 정의 정리**: `#REF!`로 깨진 이름 중 어느 수식에서도 쓰이지 않는 것만 지운다.
  `--aggressive-names`를 주면 쓰이지 않는 숨김 이름도 지운다.
- 숨긴 탭은 엑셀에서 탭 우클릭 > 숨기기 취소로 언제든 다시 볼 수 있다.

## 자료 보안

- 실제 회사 자료(엑셀, 조서)는 이 저장소에 올리지 않는다. `.gitignore`가 `*.xlsx`, `*.xls`, `data/`, `output/`를 막는다.
- 테스트는 `tests/conftest.py`가 만드는 가짜 파일만 쓴다.

## 테스트

```
python -m pytest
```

실제 조서·회사 자료는 저장소 안 `data` 폴더에 넣어 쓴다 (`.gitignore`로 막혀 있어 GitHub에 올라가지 않는다).
