# -*- coding: utf-8 -*-
"""All Korean text used by the VBA module (key -> text). Stored in the hidden _text sheet.
'\\n' in a value becomes a line break inside VBA."""

SHEETS = {
    'home': '홈', 'list': '검수목록', 'cal': '달력', 'detail': '상세', 'import': '가져오기', 'mail': '메일',
    'nego': '협의기록', 'att': '첨부', 'maillog': '메일기록', 'hist': '변경이력', 'orig': '신청원본',
    'vendor': '업체', 'insp': '검수자', 'tpl': '메일양식', 'set': '설정', 'help': '사용안내',
}

STATUS = {'registered': '등록', 'nego': '일정협의중', 'confirmed': '확정', 'changed': '변경됨', 'done': '검수완료', 'canceled': '취소'}

# vendor request columns (header|aliases)
HEADERS = {
    'id': '신청번호|접수번호',
    'vendorName': '업체명|업체|회사명',
    'projectName': '프로젝트명|프로젝트|PJT|PJT명|project',
    'category': '검수 품목|검수품목|품목',
    'itemName': '품명/품번|품명|품명 및 품번|품명(품번)|품명품번|품번/품명',
    'spec': '품번/규격|품번|규격|품번규격',
    'quantity': '수량',
    'requestedDate': '검수 희망일|희망일|검수희망일',
    'requestedTime': '희망 시간|희망시간|시간',
    'location': '검수 장소|검수장소|장소',
    'contactName': '업체 담당자|담당자|업체담당자',
    'contactPhone': '연락처|전화번호|휴대폰',
    'contactEmail': '이메일|메일|e-mail|email',
    'note': '비고',
}

PLACEHOLDERS = {
    'vendor': '업체명', 'project': '프로젝트명', 'id': '신청번호', 'item': '품명', 'qty': '수량', 'category': '검수품목',
    'wishDate': '희망일', 'wishTime': '희망시간', 'date': '검수일', 'time': '검수시간', 'place': '검수장소', 'insp': '검수자',
    'contact': '담당자명', 'list': '목록', 'count': '건수', 'confirmedOn': '확정일', 'result': '결과',
    'failReason': '불합격사유', 'changeReason': '변경사유', 'cancelReason': '취소사유', 'sender': '보내는사람',
    'period': '기간', 'deadline': '기한',
}

TEXT = {
    'app.title': '검수일정 관리',
    'yes': '예', 'no': '아니오',
    'sep': ' · ',
    'weekdays': '일,월,화,수,목,금,토',
    'time.pm': '오후', 'time.am': '오전', 'time.half': '반',
    'res.pass': '합격', 'res.fail': '불합격',
    'dir.orig': '원본', 'dir.att': '첨부', 'dir.backup': '백업',
    'filter.all': '모든 파일 (*.*),*.*',
    'filter.excel': '엑셀 파일 (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls',
    'filter.xlsx': '엑셀 통합 문서 (*.xlsx),*.xlsx',

    # words
    'w.count': '{0}건', 'w.countItems': '{0}건(품목 {1}개)', 'w.days': '{0}일',
    'w.dday': 'D-Day', 'w.holiday': '휴무일', 'w.manual': '직접 입력', 'w.project': '프로젝트', 'w.qty': '수량',
    'w.when': '검수 일시', 'w.wish': '희망 일시', 'w.place': '장소', 'w.insp': '검수자', 'w.unconfirmed': '미확정',
    'w.all': '전체', 'w.year': '년', 'w.month': '월', 'w.none': '없음', 'w.negoStart': '협의 기록 추가로 협의 시작',
    'w.resultCleared': '결과 삭제', 'w.noResult': '결과 미입력', 'w.noNotice': '결과 통보 메일 미발송',
    'w.inspDate': '검수일',
    'w.when2': '일시', 'w.writer': '작성자', 'w.content': '내용', 'w.kind': '종류', 'w.file': '파일', 'w.subject': '제목',
    'w.to': '받는 사람', 'w.field': '항목', 'w.before': '이전 값', 'w.after': '새 값', 'w.reason': '사유',

    'f.status': '상태', 'f.created': '등록', 'f.force': '충돌 강제 진행', 'f.att': '첨부파일', 'f.mail': '메일 보냄',
    'hist.import': '업체 엑셀 가져오기 ({0} {1}행)',

    'left.days': '{0}일 남음', 'left.today': '오늘까지', 'left.over': '기한 {0}일 초과', 'left.overWord': '초과',

    # conflicts
    'cf.holiday': '검수 불가일 — {0} {1}',
    'cf.weekend': '주말 — {0}',
    'cf.day': '하루 최대 {0}건 초과 — {1}',
    'cf.insp': '검수자 1인당 하루 최대 {0}건 초과 — {1} {2}',
    'cf.vendor': '같은 업체 하루 최대 {0}건 초과 — {1} {2}',
    'cf.qty': '하루 최대 수량 {0}개 초과 — 합계 {1}개',
    'cf.slot': '같은 장소 {0}분 안에 다른 업체 검수 — {1} {2}',

    # to-do
    'todo.nego': '일정 협의 필요',
    'todo.negoLate': '협의 {0}일째',
    'todo.reconfirm': '재확정 필요',
    'todo.mail.confirmed': '확정 메일 미발송',
    'todo.mail.changed': '변경 메일 미발송',
    'todo.mail.canceled': '취소 메일 미발송',
    'todo.noResult': '결과 미입력 (통보 기한 {0}, {1})',
    'todo.notice': '결과 통보 메일 미발송 (기한 {0}, {1})',

    # questions
    'ask.cancel': '{0} 건을 취소합니다.\n취소 사유를 입력하세요. (업체 취소 메일에 들어갑니다)',
    'ask.change': '{0} 건은 확정된 일정입니다. 바꾸면 \'변경됨\'이 되고 다시 확정해야 합니다.\n변경 사유를 입력하세요.',
    'ask.result': '{0} 검수 결과를 입력하세요.\n1 = 합격, 2 = 불합격',
    'ask.fail': '{0} 불합격 사유를 입력하세요. (업체 결과 메일에 들어갑니다)',
    'ask.force': '충돌이 있지만 확정합니다. 강제 진행 사유를 입력하세요.',
    'ask.nego': '{0} 일정 협의 내용을 입력하세요.\n예: 업체 요청으로 12/3 오후로 조정 협의 중',
    'ask.attkind': '첨부 종류를 고르세요.\n1 = 사진, 2 = 성적서, 3 = 기타 (그 밖의 글자를 쓰면 그대로 종류로 저장)',
    'ask.vs.vendor': '일정 확인서를 만들 업체명을 입력하세요.',
    'ask.vs.from': '기간 시작일을 입력하세요. (예: 2026-12-01)',
    'ask.vs.to': '기간 종료일을 입력하세요. (예: 2026-12-31)',
    'att.pick': '{0}에 첨부할 파일을 고르세요 (여러 개 선택 가능)',
    'att.photo': '사진', 'att.report': '성적서', 'att.etc': '기타',

    # messages
    'msg.error': '처리 중 오류가 났습니다: {0}\n방금 바꾼 칸의 값을 확인하세요.',
    'msg.reason.required': '사유를 입력해야 진행할 수 있습니다. 변경하지 않았습니다.',
    'msg.readonly': '다른 사람이 이 파일을 열어 두어 \'읽기 전용\'으로 열렸습니다.\n\n보기는 할 수 있지만 바꾼 내용은 저장되지 않습니다. 수정이 필요하면 먼저 연 사람이 닫은 뒤 다시 여세요.',
    'msg.readonly.edit': '읽기 전용으로 열린 파일이라 바꾼 내용이 저장되지 않습니다. 먼저 연 사람이 닫은 뒤 다시 열어 수정하세요.',
    'msg.bigpaste': '한 번에 너무 많은 칸을 바꿔 자동 점검을 건너뛰었습니다. 이력은 남지 않았습니다.',
    'msg.status.invalid': '상태는 목록에서 고르세요: 등록, 일정협의중, 확정, 변경됨, 검수완료, 취소',
    'msg.status.notallowed': '\'{0}\' 상태에서 \'{1}\'(으)로 바꿀 수 없습니다.\n\n- 확정: 검수일·시간·장소·검수자를 넣은 뒤\n- 검수완료: [결과] 칸에 합격/불합격 입력\n- 검수완료·취소된 건은 다시 바꿀 수 없습니다.',
    'msg.confirm.missing': '{0} 건을 확정하려면 아래 칸을 먼저 채우세요.',
    'msg.insp.inactive': '\'{0}\'은(는) [검수자] 시트에 없거나 사용하지 않는 검수자입니다.',
    'msg.conflict.block': '{0} 건은 충돌 조건에 걸려 확정할 수 없습니다.\n\n- {1}\n\n(설정에서 \'충돌 시 강제 진행 허용\'을 \'예\'로 바꾸면 사유를 적고 진행할 수 있습니다)',
    'msg.conflict.ask': '{0} 건에 충돌 경고가 있습니다.\n\n- {1}\n\n그래도 확정할까요?',
    'msg.schedule.locked': '{0} 건은 \'{1}\' 상태라 일정을 바꿀 수 없습니다.',
    'msg.result.clear': '{0} 결과를 지우면 상태가 \'확정\'으로 돌아갑니다. 지울까요?',
    'msg.result.invalid': '결과는 \'합격\' 또는 \'불합격\'만 넣을 수 있습니다.',
    'msg.result.notconfirmed': '{0} 건은 아직 확정되지 않아 결과를 넣을 수 없습니다. 먼저 상태를 \'확정\'으로 바꾸세요.',
    'msg.time.invalid': '시간 형식이 올바르지 않습니다. 24시간 형식으로 적어 주세요. 예: 13:30, 1330, 오후 2시',
    'msg.date.invalid': '날짜 형식이 올바르지 않습니다. 예: 2026-12-01',
    'msg.qty.invalid': '수량은 1 이상의 정수로 입력하세요. 예: 10',
    'msg.email.invalid': '이메일 형식이 올바르지 않습니다. 예: name@company.co.kr (저장은 됐습니다)',
    'msg.pick.row': '[검수목록]에서 처리할 건의 행을 먼저 선택하세요.',
    'msg.filtered.date': '{0} 검수 건만 보이도록 필터했습니다. [필터 해제]로 전체를 다시 봅니다.',
    'msg.save.first': '먼저 이 파일을 공유 폴더에 .xlsm 으로 저장하세요.',
    'msg.dir.fail': '폴더를 만들 수 없습니다: {0}',
    'msg.copy.fail': '파일을 복사하지 못했습니다: {0}\n{1}',
    'msg.att.done': '{0}에 파일 {1}개를 첨부했습니다. [상세]에서 눌러 열 수 있습니다.',
    'msg.mail.drafted': '메일 {0}통을 작성했습니다.\n\n[메일] 시트에서 내용을 확인하고 [본문 복사] → 웹메일에 붙여넣어 보낸 뒤 [보냄으로 기록]을 누르세요.',
    'msg.mail.clear': '작성된 메일 초안을 모두 지울까요? (메일 기록은 지워지지 않습니다)',
    'msg.mail.pick': '[메일] 시트에서 메일 초안이 있는 행을 선택하세요.',
    'msg.copied': '{0}을(를) 복사했습니다. 웹메일에 Ctrl+V 로 붙여넣으세요.',
    'msg.copy.manual': '자동 복사가 되지 않았습니다. 선택된 칸을 더블클릭해 내용을 드래그한 뒤 Ctrl+C 로 복사하세요.',
    'msg.webmail.none': '[설정] 시트의 \'웹메일 주소\'를 먼저 입력하세요.',
    'msg.open.fail': '열 수 없습니다: {0}',
    'msg.mail.already': '{0}번 메일은 이미 보냄으로 기록했습니다. 한 번 더 기록할까요?',
    'msg.mail.marked': '메일 {0}통을 보냄으로 기록했습니다.',
    'mail.sentAt': '보냄 {0}',
    'msg.none': '검수 건이 없습니다.',
    'msg.reminder.none': '알림을 보낼 검수가 없습니다. (확정된 검수 중 기간 안에 있는 것, 검수자 이메일이 있는 것만 대상)',
    'msg.result.none': '결과 통보를 기다리는 검수가 없습니다.',
    'msg.setup.done': '버튼·폴더·보호 설정을 다시 만들었습니다.',
    'msg.tpl.notSaved': '양식을 저장하지 않았습니다. 열린 새 통합 문서를 직접 저장해도 됩니다.',
    'msg.tpl.saved': '업체용 검수 신청 양식을 저장했습니다.\n{0}\n\n이 파일을 업체에 메일로 보내세요.',
    'msg.vs.saved': '일정 확인서({0}건)를 저장했습니다.\n{1}',
    'msg.export.saved': '보이는 행만 내보냈습니다.\n{0}',
    'msg.imp.clear': '가져오기 미리보기를 모두 지울까요?',
    'msg.imp.empty': '먼저 [파일 선택]으로 업체 엑셀을 고르세요.',
    'msg.imp.nothing': '\'포함\'이 \'예\'인 행이 없습니다.',
    'msg.imp.withErrors': '오류가 있는 {0}행을 포함했습니다.\n형식이 잘못된 칸(수량·날짜·시간)은 빈 값으로 저장되고, 없는 신청번호는 새 번호로 등록됩니다.\n\n그대로 가져올까요?',
    'msg.imp.confirm': '{0}행을 가져올까요?',
    'msg.imp.done': '가져오기 완료\n\n신규 {0}건 · 수정 {1}건 · 제외 {2}건\n새 신청번호: {3}',
    'msg.imp.copyFail': '아래 원본 파일은 [원본] 폴더에 복사하지 못했습니다 (열려 있으면 닫고 직접 복사하세요):',
    'msg.imp.mailAsk': '새로 등록한 건의 접수 안내 메일을 작성할까요?',
    'msg.imp.rechecked': '다시 검사했습니다. 오류 {0}행',

    # import
    'imp.pick': '업체가 보낸 검수 신청 엑셀을 고르세요 (여러 개 선택 가능)',
    'imp.example': '예시',
    'imp.kind.new': '신규', 'imp.kind.update': '수정', 'imp.kind.error': '오류',
    'imp.noHeader': '검수 신청 양식의 머리글(업체명·검수 품목·품명/품번·수량 등)을 찾을 수 없습니다. 우리 양식으로 다시 작성해 달라고 요청하세요.',
    'imp.missingCols': '필수 열이 없습니다: {0}. 머리글 이름을 바꾸지 말고 다시 작성해 달라고 요청하세요.',
    'imp.openFail': '파일을 열 수 없습니다: {0}',
    'imp.e.req': '필수: {0}',
    'imp.e.qty': '수량은 1 이상 정수',
    'imp.e.date': '날짜 형식 오류 (예: 2026-12-01)',
    'imp.e.time': '시간 형식 오류 (예: 13:30)',
    'imp.e.email': '이메일 형식 오류',
    'imp.e.noId': '없는 신청번호 (포함하면 새 번호로 등록)',
    'imp.w.past': '희망일이 오늘({0})보다 이전',
    'imp.w.category': '설정의 검수 품목 목록에 없음',
    'imp.w.dupId': '같은 신청번호가 여러 번 나옴',
    'imp.w.conflict': '[충돌: {0}]',
    'imp.diff': '변경: {0}',
    'imp.noChange': '바뀐 내용 없음',
    'imp.rowLabel': '미리보기 {0}행',

    # home
    'home.todaySub': '완료 {0} · 예정 {1}',
    'home.cfFirst': '가장 가까운 날: {0}',
    'home.cfNone': '오늘 이후 충돌 없음',
    'home.todoSee': '아래 목록 확인',
    'home.todoNone': '모두 처리됨',
    'home.month': '이번 달({0}) 상태별: 등록 {1} · 일정협의중 {2} · 확정 {3} · 변경됨 {4} · 검수완료 {5} · 취소 {6}',
    'home.updated': '새로고침 {0} · 사용자 {1}',
    'home.readonly': '※ 읽기 전용으로 열림 (저장 안 됨)',

    # detail
    'detail.notfound': '신청번호 {0}을(를) 찾을 수 없습니다.',
    'detail.colItem': '항목', 'detail.colOrig': '업체 신청(원본)', 'detail.colCur': '현재',
    'detail.noOrig': '원본 없음 (직접 입력한 건)',
    'detail.nego': '일정 협의 기록', 'detail.att': '첨부파일', 'detail.mail': '메일 기록', 'detail.hist': '변경 이력 (최근 40개)',

    'popup.head': '{0} 님에게 배정된 검수 알림',
    'popup.up': '다가오는 검수 {0}건',
    'popup.result': '결과 통보가 필요한 검수 {0}건 (검수일 다음날부터 {1}일 안)',

    # vendor request form
    'tpl.sheet': '검수신청', 'tpl.guideSheet': '작성 안내', 'tpl.fileName': '검수신청양식.xlsx',
    'tpl.heads': '신청번호|업체명|프로젝트명|검수 품목|품명/품번|수량|검수 희망일|희망 시간|검수 장소|업체 담당자|연락처|이메일|비고',
    'tpl.required': '0|1|0|1|1|1|1|0|0|1|0|1|0',
    'tpl.example': '예시|한솔정밀|P-2611 CVD 챔버||BP-2205 베이스 플레이트|10||10:30|수입검사실 A|김철수|010-1234-5678|kim@hansol.example.com|예시 행입니다. 3행부터 작성하세요.',
    'tpl.v.category': '목록에 있는 검수 품목을 고르세요.',
    'tpl.v.date': '날짜로 입력하세요. 예: 2026-12-01',
    'tpl.v.dateHint': '예: 2026-12-01',
    'tpl.v.qty': '1 이상의 숫자만 입력하세요.',
    'tpl.v.timeHint': '24시간 형식. 예: 13:30',
    'tpl.guide': ('검수 신청서 작성 안내\n\n'
                  '1. [검수신청] 시트의 3행부터 검수 1건당 한 행씩 작성합니다. 2행은 작성 예시입니다.\n'
                  '2. 신규 신청은 [신청번호]를 비워 두세요. 등록 후 신청번호를 메일로 알려 드립니다.\n'
                  '3. 이미 받은 신청번호의 일정·수량 등을 고칠 때는 [신청번호]를 적고 바뀐 내용을 입력해 회신하세요.\n'
                  '4. [검수 품목]은 칸을 누르면 나오는 목록에서 고르세요. [프로젝트명]은 해당되는 프로젝트가 있으면 적어 주세요.\n'
                  '   [품명/품번]은 한 칸에 품번과 품명을 함께 적습니다. 예: BP-2205 베이스 플레이트\n'
                  '5. [검수 희망일]은 2026-12-01 처럼 날짜로, [희망 시간]은 13:30 처럼 24시간 형식으로 적습니다.\n'
                  '6. [수량]은 1 이상의 숫자만 적습니다 (단위 제외).\n'
                  '7. [업체 담당자]와 [이메일]은 일정 확정·결과 통보 메일을 받을 분으로 적어 주세요.\n'
                  '8. 머리글(1행)과 열 순서는 바꾸지 마세요.\n\n'
                  '* 표시가 있는 항목은 필수입니다.'),

    # vendor schedule
    'vs.sheet': '일정 확인서', 'vs.title': '검수 일정 확인서',
    'vs.vendor': '업체명', 'vs.period': '기간', 'vs.made': '작성일', 'vs.by': '작성자',
    'vs.heads': '신청번호|프로젝트명|품명/품번|수량|검수일|시간|장소|검수자|상태|비고',
    'vs.canceled': '취소됨', 'vs.changed': '일정 변경 — 재확정 예정', 'vs.unconfirmed': '일정 미확정',
    'vs.fileName': '검수일정확인서_{0}_{1}.xlsx',
    'export.fileName': '검수목록_{0}.xlsx',

    # buttons
    'btn.refresh': '새로고침', 'btn.import': '엑셀 가져오기', 'btn.template': '업체 신청 양식 만들기',
    'btn.list': '검수목록', 'btn.cal': '달력', 'btn.mail': '메일', 'btn.remind': '검수자 알림 메일 작성',
    'btn.resultRemind': '결과 통보 요청 메일 작성', 'btn.vendorSchedule': '업체 일정 확인서',
    'btn.backupFolder': '백업 폴더 열기', 'btn.origFolder': '업체 원본 폴더 열기',
    'btn.mailSel': '선택 건 메일 작성', 'btn.detail': '상세 보기', 'btn.nego': '협의 기록 추가',
    'btn.att': '첨부 추가', 'btn.clearFilter': '필터 해제', 'btn.export': '보이는 행 내보내기',
    'btn.recalc': '다시 계산', 'btn.home': '홈',
    'btn.pickFiles': '파일 선택', 'btn.recheck': '다시 검사', 'btn.commit': '가져오기 확정', 'btn.clear': '비우기',
    'btn.copyBody': '본문 복사', 'btn.copySubj': '제목 복사', 'btn.copyTo': '받는 사람 복사',
    'btn.webmail': '웹메일 열기', 'btn.markSent': '보냄으로 기록',
    'btn.prev': '◀ 이전 달', 'btn.today': '이번 달', 'btn.next': '다음 달 ▶',
    'btn.toList': '목록에서 보기', 'btn.mailOne': '메일 작성',
}

for k, v in SHEETS.items():
    TEXT['sh.' + k] = v
for k, v in STATUS.items():
    TEXT['st.' + k] = v
for k, v in HEADERS.items():
    TEXT['hdr.' + k] = v
for k, v in PLACEHOLDERS.items():
    TEXT['ph.' + k] = v

SIGN = '\n\n감사합니다.\n{보내는사람} 드림'
TEMPLATES = [
    ('registered', '접수 안내', '[검수 신청 접수] {신청번호} {품명} ({업체명})',
     '{담당자명} 님께,\n\n요청하신 검수 신청이 접수되었습니다.\n\n{목록}\n\n일정은 협의 후 확정되면 다시 안내드리겠습니다.' + SIGN),
    ('confirmed', '일정 확정', '[검수 일정 확정 안내] {업체명} {건수}건 ({확정일})',
     '{담당자명} 님께,\n\n요청하신 검수 신청 {건수}건의 일정이 아래와 같이 확정되어 안내드립니다.\n\n{목록}\n\n일정에 맞춰 준비 부탁드립니다.\n내용을 수정하실 때는 엑셀의 [신청번호] 칸에 위 신청번호를 적어 회신해 주세요.' + SIGN),
    ('changed', '일정 변경', '[검수 일정 변경 안내] {신청번호} {품명} ({업체명})',
     '{담당자명} 님께,\n\n신청번호 {신청번호} ({품명}) 검수 일정이 아래와 같이 변경되었습니다.\n\n- 프로젝트명: {프로젝트명}\n- 변경 일시: {검수일} {검수시간}\n- 검수 장소: {검수장소}\n- 변경 사유: {변경사유}\n\n확정되면 다시 안내드리겠습니다.' + SIGN),
    ('canceled', '취소 안내', '[검수 취소 안내] {신청번호} {품명} ({업체명})',
     '{담당자명} 님께,\n\n신청번호 {신청번호} ({품명}) 검수가 취소되었습니다.\n\n- 프로젝트명: {프로젝트명}\n- 취소 사유: {취소사유}\n\n문의 사항이 있으시면 회신 부탁드립니다.' + SIGN),
    ('pass', '결과 - 합격', '[검수 결과 안내 - 합격] {신청번호} {품명} ({업체명})',
     '{담당자명} 님께,\n\n신청번호 {신청번호} ({품명}) 검수 결과를 알려드립니다.\n\n- 프로젝트명: {프로젝트명}\n- 검수일: {검수일}\n- 결과: {결과}\n\n수고하셨습니다.' + SIGN),
    ('fail', '결과 - 불합격', '[검수 결과 안내 - 불합격] {신청번호} {품명} ({업체명})',
     '{담당자명} 님께,\n\n신청번호 {신청번호} ({품명}) 검수 결과를 알려드립니다.\n\n- 프로젝트명: {프로젝트명}\n- 검수일: {검수일}\n- 결과: {결과}\n- 불합격 사유: {불합격사유}\n\n시정 조치 후 재검수 일정을 협의해 주시기 바랍니다.' + SIGN),
    ('reminder', '검수자 알림', '[검수 알림] {검수자} 님 배정 검수 {건수}건 ({기간})',
     '{검수자} 님,\n\n배정된 검수 일정을 알려드립니다.\n\n{목록}\n\n상세 내용은 검수일정 관리 엑셀에서 확인하세요.\n\n{보내는사람} 드림'),
    ('resultReminder', '결과 통보 요청', '[결과 통보 요청] {검수자} 님 결과 통보 대기 {건수}건',
     '{검수자} 님,\n\n검수 후 업체에 결과를 아직 통보하지 않은 건이 있습니다.\n검수일 다음날부터 {기한}일 안에 결과를 입력하고 업체에 통보해 주세요.\n\n{목록}\n\n검수일정 관리 엑셀의 [검수목록]에서 [결과]를 입력하고 [선택 건 메일 작성]으로 결과 메일을 만들 수 있습니다.\n\n{보내는사람} 드림'),
]
