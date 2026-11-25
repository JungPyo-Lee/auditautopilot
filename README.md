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

3단계: 관리 화면 (`serve`)

| 기능 | 설명 |
| --- | --- |
| 회사·감사 등록 | 회사코드, 담당자 이메일, 업종(일반·건설). 감사 등록 시 감사준비자료 양식으로 요청 목록 자동 생성 |
| 자료요청(PBC) 관리 | 항목별 상태(미요청·요청·일부수령·수령·검토완료·보완요청·해당없음), 요청일·수령일, 비고. 엑셀로 내려받기 |
| 메일 일정 | 현장감사 D-7 최초 요청, D-3부터 당일까지 매일 미제출 항목만 독촉. 주말·공휴일 건너뜀(설정) |
| 메일 보내기 | 오늘 보낼 메일을 현황 화면에 띄우고 [지금 보내기](메일 계정 연결 시) 또는 아웃룩용 .eml 내려받기. 자동 발송을 켜면 매일 정한 시각에 보냄 |
| 메일 받기 | 연결한 메일함을 주기적으로 확인해 회사 담당자(주소·회사 도메인) 또는 제목의 [C001-2026-기말] 꼬리표로 감사를 찾고 첨부를 폴더 규칙대로 저장. 대용량 링크는 감사 화면에 모아 보여줌. 읽음 표시는 바꾸지 않음 |
| 받은 자료 정리 | 화면에서 올리거나 수신함 폴더에 넣으면 폴더 규칙대로 저장, 파일명의 요청번호·자료명으로 항목을 찾아 '수령' 처리. 못 찾은 파일은 미분류 폴더 |
| 폴더 규칙 | `{루트}/{회사순번:02}.{회사명}/...` 처럼 칸을 조합. JP님 현재 규칙과 제안 규칙을 미리 넣어 둠 |

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

# 관리 화면 열기 (브라우저가 자동으로 열린다. 끝내려면 Ctrl+C)
python -m auditautopilot serve

# 표본: 외상매출금(108) 거래처 잔액에서 10억 이상 전수 + 나머지 MUS 25건
python -m auditautopilot sample 재무제표등.xlsx --account 108 --threshold 1000000000 --method mus --n 25 -o 표본.xlsx
```

## 메일 계정

IMAP(받기)·SMTP(보내기)를 쓰는 메일이면 된다. 설정 > 메일 계정에서 Gmail, 다음(한메일), 네이버를 고르면 서버 주소가 채워지고,
그 밖의 메일은 서버 주소를 직접 넣는다. 비밀번호는 보통 각 메일 보안 설정에서 만든 '앱 비밀번호'를 쓴다.
- [Gmail을 다른 메일 프로그램에 추가하기 (Google 고객센터)](https://support.google.com/mail/answer/7126229?hl=ko)
- 다음·네이버 서버 값: [Daum](https://www.getmailbird.com/setup/access-daum-net-via-imap-smtp), [Naver](https://www.getmailbird.com/setup/access-naver-com-via-imap-smtp)

비밀번호는 데이터베이스에 저장하지 않는다. 프로그램이 켜져 있는 동안 메모리에만 두고, 윈도우에서는 원하면
'자격 증명 관리자'(제어판)에 `auditautopilot:메일주소`로 저장한다.

## 관리 화면 쓰는 순서

1. `python -m auditautopilot serve` → 브라우저에서 http://127.0.0.1:8765
2. **설정**: 회계법인 이름, 작성자, 연락처, 저장 루트 폴더, 폴더 규칙, 공휴일
3. **회사**: 회사코드·회사명·담당자 이메일 등록
4. **감사 등록**: 감사구분과 현장감사 일정을 넣으면 요청 목록이 만들어진다
5. **현황**: 오늘 보낼 메일 → 미리보기 → 아웃룩으로 열기(.eml) → 보낸 뒤 [보냄으로 표시]
6. 받은 자료는 감사 화면에서 올리거나 수신함 폴더에 넣고 [수신함 정리]

데이터는 `data/audit.db` 한 파일에 저장된다. 기본으로 이 PC에서만 접속되며,
팀 공유(`--host 0.0.0.0`)는 로그인 기능을 붙이기 전까지 사무실 내부망에서만 쓴다.

## 지원하는 파일 형식

- **세무사랑 '재무제표 등' 묶음**: TB, BS, PL, 분개장, 계정코드 이름의 거래처원장 시트
- **더존 분개장**: 전표일자 | 전표번호 | 구분 | Code | 계정과목 | 차변 | 대변 | 거래처 | Code | 적요
- 더존 합계잔액시산표, 재무제표, 거래처원장은 샘플을 받으면 추가한다.

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
