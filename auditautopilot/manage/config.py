"""설정 기본값. 화면의 '설정'에서 바꾸면 데이터베이스에 저장된다."""

from __future__ import annotations

REQUEST_SUBJECT = "[{회계법인}] {회사명} {사업연도} {감사구분} 자료 요청 ({현장감사 시작일} 방문 예정) [{감사코드}]"
REQUEST_BODY = """{담당자}님, 안녕하세요. {회계법인} {작성자}입니다.

{현장감사 시작일}부터 {현장감사 종료일}까지 {사업연도} {감사구분} 현장감사를 진행할 예정입니다.
원활한 감사를 위해 아래(첨부) 목록의 자료를 {제출기한}까지 보내주시면 감사하겠습니다.

{요청 자료 표}

- 회신하실 때 메일 제목의 [{감사코드}]는 지우지 말아 주세요. 보내주신 자료가 자동으로 정리됩니다.
- 파일 이름 앞에 요청번호(예: 중간-03)를 붙여 주시면 더 정확하게 정리됩니다.
- 기본자료는 작년에 제출하셨다면 변동된 내용만 보내주셔도 됩니다.
- 용량이 큰 파일은 대용량 첨부로 보내주셔도 됩니다.

궁금하신 점은 {연락처}로 연락 주세요. 감사합니다."""

REMINDER_SUBJECT = "[{회계법인}] {회사명} 감사 자료 안내 - 미제출 {미제출 건수}건 (현장감사 {D-n}) [{감사코드}]"
REMINDER_BODY = """{담당자}님, 안녕하세요. {회계법인} {작성자}입니다.

보내주신 자료는 잘 받았습니다. 감사합니다.
{현장감사 시작일} 현장감사를 앞두고 아직 받지 못한 자료를 다시 안내드립니다.

{미제출 자료 표}

준비에 시간이 걸리는 자료는 가능한 일정을 회신해 주시면 감사 일정에 반영하겠습니다. 감사합니다."""

QUESTION_SUBJECT = "[{회계법인}] {회사명} {사업연도} {감사구분} 질문 드립니다 ({질문 건수}건) [{감사코드}]"
QUESTION_BODY = """{담당자}님, 안녕하세요. {회계법인} {작성자}입니다.

{사업연도} {감사구분}과 관련하여 아래 사항을 여쭙습니다.
첨부한 질문서 엑셀의 '답변' 칸에 적어 {회신기한}까지 이 메일에 회신해 주시고, 근거 자료가 있으면 함께 보내 주세요.

{질문 표}

문의: {작성자} ({연락처})
감사합니다.
"""
DDAY_BODY = """{담당자}님, 안녕하세요. {회계법인} {작성자}입니다.

오늘부터 현장감사가 시작됩니다. 아래 자료는 방문 중에 받을 수 있도록 준비 부탁드립니다.

{미제출 자료 표}

감사합니다."""

# JP님 현재 규칙: 감사/01.회사명/01.2026/01.중간감사/01.PBC
JP_RULE = "{루트}/{회사순번:02}.{회사명}/{연도순번:02}.{연도}/{감사구분번호:02}.{감사구분}/{분류번호:02}.{분류}/{원본파일명}"
# 제안 규칙: 감사/C001_한국상사/FY2026/2_기말/01_PBC/20270108_기말-12_매출채권명세_AR.xlsx
PROPOSED_RULE = ("{루트}/{회사코드}_{회사명}/FY{연도}/{감사구분번호}_{감사구분약칭}/{분류번호:02}_{분류}/"
                 "{수신일}_{요청번호}_{항목}_{원본파일명}")
RULE_PRESETS = {"JP님 현재 규칙": JP_RULE, "제안 규칙": PROPOSED_RULE}

DEFAULTS = {
    "firm_name": "",
    "author": "",
    "contact": "",
    "sender_email": "",
    "request_offset": 7,      # 현장감사 시작 D-7 최초 요청
    "reminder_from": 3,       # D-3부터 당일까지 매일 독촉
    "due_offset": 4,          # 제출기한 = D-4 (독촉 시작 전날)
    "skip_weekends": True,
    "holidays": [],           # 'YYYY-MM-DD' 목록. 공휴일은 직접 넣는다
    "root": "data/감사",
    "folder_rule": JP_RULE,
    "categories": ["PBC", "조서", "채권채무조회", "금융기관조회", "질문서"],
    "unmatched_category": "미분류",
    "nice_user": "",          # 한공회 전자조회(NICE 감사지원) 아이디. 비밀번호는 저장하지 않거나 자격 증명 관리자에
    "nice_auto": False,
    "confirm_min_reply_rate": 80,   # 채권채무 조회 회신률(건수) 기준 %. 이보다 낮으면 경고       # 켜면 하루 한 번(발송 시각 이후) 연결된 감사 건의 회신현황을 가져온다
    "request_subject": REQUEST_SUBJECT,
    "request_body": REQUEST_BODY,
    "reminder_subject": REMINDER_SUBJECT,
    "reminder_body": REMINDER_BODY,
    "dday_body": DDAY_BODY,
    "question_subject": QUESTION_SUBJECT,
    "question_body": QUESTION_BODY,
    "question_due_days": 5,   # 질문서 회신기한 = 보낸 날 + n일
    # 메일 계정 여러 개: [{user, provider, imap_host, imap_port, smtp_host, smtp_port}]
    # 비밀번호는 여기에 저장하지 않는다
    "mail_accounts": [],
    "send_from": "",          # 기본 보내는 계정 (회사별로 따로 정할 수 있다)
    "auto_send": False,       # 켜면 매일 send_hour 이후 오늘 메일을 자동 발송
    "send_hour": 9,
    "check_minutes": 10,      # 받은 메일 확인 주기
}


def load(store) -> dict:
    from . import storage
    cfg = {k: store.get_setting(k, v) for k, v in DEFAULTS.items()}
    # 저장 폴더: 예전 상대 경로('data/감사')나 빈 값이면 프로그램 바깥 고정 폴더로 바꾼다 (업데이트해도 안 지워지게)
    if not cfg["root"] or cfg["root"] == storage.LEGACY_ROOT:
        cfg["root"] = str(storage.default_root())
        if store.get_setting("root") != cfg["root"]:
            store.set_setting("root", cfg["root"])
    if not cfg["mail_accounts"] and store.get_setting("mail_user"):
        # 계정 하나만 있던 이전 설정을 옮긴다
        old = {k: store.get_setting(k) for k in ("mail_user", "mail_provider", "imap_host", "imap_port",
                                                   "smtp_host", "smtp_port")}
        cfg["mail_accounts"] = [{"user": old["mail_user"], "provider": old["mail_provider"] or "직접 입력",
                                 "imap_host": old["imap_host"], "imap_port": old["imap_port"],
                                 "smtp_host": old["smtp_host"], "smtp_port": old["smtp_port"]}]
        cfg["send_from"] = cfg["send_from"] or old["mail_user"]
        store.set_setting("mail_accounts", cfg["mail_accounts"])
        store.set_setting("send_from", cfg["send_from"])
    return cfg
