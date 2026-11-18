# -*- coding: utf-8 -*-
"""Builds 검수일정관리.xlsx (sheets, formats, validations, settings, texts).
The VBA code (vba/modInspection.bas, vba/clsAppEvents.cls) is imported into it once in Excel."""
import sys
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL

from texts import TEXT, SHEETS, STATUS, TEMPLATES, PLACEHOLDERS

OUT = sys.argv[1] if len(sys.argv) > 1 else 'dist/검수일정관리.xlsx'
DATA_ROWS = 3000  # editable rows prepared in the list sheet

NAVY = '1F2F63'
HEAD_FILL = PatternFill('solid', fgColor=NAVY)
HEAD_FONT = Font(bold=True, color='FFFFFF')
SUB_FILL = PatternFill('solid', fgColor='E8ECF4')
INPUT_FILL = PatternFill('solid', fgColor='FFF7D6')
AUTO_FILL = PatternFill('solid', fgColor='F2F2F2')
THIN = Side(style='thin', color='C8CFDC')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
UNLOCK = Protection(locked=False)
WRAP_TOP = Alignment(wrap_text=True, vertical='top')
DATE_FMT = 'yyyy-mm-dd (aaa)'
STAMP_FMT = 'yyyy-mm-dd hh:mm'

wb = Workbook()
wb.remove(wb.active)


def sheet(key):
    return wb.create_sheet(SHEETS[key])


def header(ws, row, heads, widths=None, fill=HEAD_FILL, font=HEAD_FONT):
    for i, h in enumerate(heads, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill = fill
        c.font = font
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = BOX
        if widths:
            ws.column_dimensions[CL(i)].width = widths[i - 1]
    ws.row_dimensions[row].height = 30


def name(nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def q(sheet_key):
    return "'" + SHEETS[sheet_key] + "'"


def hint(ws, cell, text):
    ws[cell] = text
    ws[cell].font = Font(color='55607A', size=9)


# ───────────────────────── 홈 ─────────────────────────
home = sheet('home')
home['A1'] = '검수일정 관리'
home['A1'].font = Font(bold=True, size=18, color=NAVY)
home['A2'].font = Font(color='55607A', size=9)
for r in (3, 4):
    home.row_dimensions[r].height = 30
for col, w in zip('ABCDEFGHIJK', [15, 14, 24, 12, 40, 2, 8, 17, 15, 30, 28]):
    home.column_dimensions[col].width = w
cards = [('A', '오늘 검수'), ('D', '이번 주 검수'), ('G', '충돌 경고 날짜'), ('J', '처리 필요')]
for col, label in cards:
    home[f'{col}5'] = label
    home[f'{col}5'].font = Font(bold=True, color='55607A')
    home[f'{col}6'].font = Font(bold=True, size=20, color=NAVY)
    home[f'{col}7'].font = Font(color='55607A', size=9)
home['A9'].font = Font(color='1F2937')
home['A10'] = '처리 필요  (신청번호를 더블클릭하면 상세)'
home['G10'] = '다가오는 검수 (7일 이내 · 확정)'
for c in ('A10', 'G10'):
    home[c].font = Font(bold=True, size=12, color=NAVY)
for i, h in enumerate(['신청번호', '업체명', '품명/품번', '상태', '처리할 일'], start=1):
    c = home.cell(row=11, column=i, value=h)
    c.fill = SUB_FILL
    c.font = Font(bold=True)
for i, h in enumerate(['D-day', '일시', '신청번호', '업체·품명', '장소·검수자'], start=7):
    c = home.cell(row=11, column=i, value=h)
    c.fill = SUB_FILL
    c.font = Font(bold=True)
home.row_dimensions[11].height = 22
home['A44'] = '충돌 경고 (오늘 이후)'
home['A44'].font = Font(bold=True, size=12, color=NAVY)
for r in range(12, 42):
    for c in range(1, 12):
        if c != 6:
            home.cell(row=r, column=c).border = Border(bottom=Side(style='hair', color='D9DEE7'))
home.freeze_panes = None

# ───────────────────────── 검수목록 ─────────────────────────
L = sheet('list')
LIST_COLS = [
    # header, width, editable, kind
    ('신청번호', 15, False, 'text'),
    ('상태', 11, True, 'status'),
    ('업체명', 14, True, 'text'),
    ('프로젝트명', 16, True, 'text'),
    ('검수 품목', 12, True, 'item'),
    ('품명/품번', 28, True, 'text'),
    ('수량', 7, True, 'qty'),
    ('검수일', 15, True, 'date'),
    ('검수시간', 9, True, 'time'),
    ('검수장소', 14, True, 'text'),
    ('검수자', 10, True, 'insp'),
    ('결과', 8, True, 'result'),
    ('실제 검수일', 15, True, 'date'),
    ('불합격 사유', 22, True, 'text'),
    ('처리 필요', 28, False, 'auto'),
    ('충돌', 30, False, 'auto'),
    ('결과 통보일', 15, False, 'adate'),
    ('희망일', 15, True, 'date'),
    ('희망시간', 9, True, 'time'),
    ('신청 장소', 14, True, 'text'),
    ('업체 담당자', 11, True, 'text'),
    ('연락처', 14, True, 'textfmt'),
    ('이메일', 24, True, 'text'),
    ('비고', 22, True, 'text'),
    ('결과 메모', 22, True, 'text'),
    ('확정일시', 16, False, 'stamp'),
    ('확정자', 10, False, 'auto'),
    ('상태 변경일시', 16, False, 'stamp'),
    ('등록일시', 16, False, 'stamp'),
    ('등록자', 10, False, 'auto'),
    ('원본 파일', 24, False, 'auto'),
    ('원본 행', 7, False, 'auto'),
    ('최종 수정', 24, False, 'autotext'),
]
assert len(LIST_COLS) == 33
L.row_dimensions[1].height = 30
hint(L, 'A2', '새 건: 맨 아래 빈 행에 업체명부터 입력하면 신청번호가 붙습니다 · 상태/결과는 칸의 목록에서 고르기 '
     '(확정하려면 검수일·시간·장소·검수자 필요) · 신청번호 더블클릭 = 상세 · 회색 칸은 자동 계산')
header(L, 3, [c[0] for c in LIST_COLS], [c[1] for c in LIST_COLS])
for i, (_h, _w, editable, kind) in enumerate(LIST_COLS, start=1):
    if not editable:
        L.cell(row=3, column=i).fill = PatternFill('solid', fgColor='55607A')
for r in range(4, 4 + DATA_ROWS):
    for i, (_h, _w, editable, kind) in enumerate(LIST_COLS, start=1):
        c = L.cell(row=r, column=i)
        if editable:
            c.protection = UNLOCK
        if kind in ('date', 'adate'):
            c.number_format = DATE_FMT
        elif kind == 'stamp':
            c.number_format = STAMP_FMT
        elif kind in ('time', 'textfmt', 'autotext'):
            c.number_format = '@'
L.freeze_panes = 'C4'
L.auto_filter.ref = f'A3:{CL(33)}{4 + DATA_ROWS - 1}'
last = 4 + DATA_ROWS - 1


def dv(ws, formula, rng, title, msg, kind='list', style='stop', **kw):
    d = DataValidation(type=kind, formula1=formula, allow_blank=True, showErrorMessage=True,
                       errorStyle=style, errorTitle=title, error=msg, **kw)
    for part in rng.split():
        d.add(part)
    ws.add_data_validation(d)


dv(L, '=list_status', f'B4:B{last}', '상태', '목록에서 고르세요: 등록, 일정협의중, 확정, 변경됨, 검수완료, 취소')
dv(L, '=list_items', f'E4:E{last}', '검수 품목', '설정의 검수 품목 목록에서 고르세요.', style='warning')
dv(L, '=list_insp', f'K4:K{last}', '검수자', '[검수자] 시트에 등록된 검수자를 고르세요.')
dv(L, '=list_result', f'L4:L{last}', '결과', '합격 또는 불합격을 고르세요.')
dv(L, '1', f'G4:G{last}', '수량', '1 이상의 정수로 입력하세요.', kind='whole', operator='greaterThanOrEqual')

st_colors = {'등록': 'E5E7EB', '일정협의중': 'FEF3C7', '확정': 'D1FAE5', '변경됨': 'FFE4C7', '검수완료': 'DBEAFE', '취소': 'F3F4F6'}
for label, color in st_colors.items():
    L.conditional_formatting.add(f'B4:B{last}', FormulaRule(formula=[f'$B4="{label}"'], fill=PatternFill('solid', fgColor=color)))
L.conditional_formatting.add(f'A4:A{last}', FormulaRule(formula=['$B4="취소"'], font=Font(strike=True, color='9CA3AF')))
L.conditional_formatting.add(f'O4:O{last}', FormulaRule(formula=['ISNUMBER(SEARCH("초과",$O4))'], fill=PatternFill('solid', fgColor='FFC7CE'), font=Font(color='9C0006', bold=True)))
L.conditional_formatting.add(f'O4:O{last}', FormulaRule(formula=['LEN($O4)>0'], fill=PatternFill('solid', fgColor='FFEB9C')))
L.conditional_formatting.add(f'P4:P{last}', FormulaRule(formula=['LEN($P4)>0'], fill=PatternFill('solid', fgColor='F8CBAD')))
L.conditional_formatting.add(f'H4:H{last}', FormulaRule(formula=['LEN($P4)>0'], fill=PatternFill('solid', fgColor='F8CBAD')))
L.conditional_formatting.add(f'L4:L{last}', FormulaRule(formula=['$L4="불합격"'], font=Font(color='C00000', bold=True)))

# ───────────────────────── 달력 ─────────────────────────
C = sheet('cal')
C['A1'] = '달력'
C['A1'].font = Font(bold=True, size=16, color=NAVY)
C['B2'] = '월'
C['C2'] = date.today().replace(day=1)
C['C2'].number_format = 'yyyy-mm'
C['C2'].protection = UNLOCK
C['C2'].fill = INPUT_FILL
C['E2'] = '검수자'
C['F2'].protection = UNLOCK
C['F2'].fill = INPUT_FILL
C['H2'].font = Font(bold=True, size=14, color=NAVY)
C.row_dimensions[3].height = 30
hint(C, 'A12', '날짜 칸을 더블클릭하면 그날 검수만 [검수목록]에 보여 줍니다. 주황색 = 충돌 경고, 회색 = 주말·휴무일, 파란색 = 오늘. 검수자 칸을 비우면 전체.')
dv(C, '=list_insp', 'F2', '검수자', '목록에서 고르거나 비워 두세요.', style='warning')
C.column_dimensions['A'].width = 3
for i, wd in enumerate(['일', '월', '화', '수', '목', '금', '토']):
    col = CL(2 + i)
    C.column_dimensions[col].width = 27
    c = C[f'{col}4']
    c.value = wd
    c.fill = HEAD_FILL
    c.font = HEAD_FONT
    c.alignment = Alignment(horizontal='center')
for r in range(5, 11):
    C.row_dimensions[r].height = 105
    for i in range(7):
        c = C.cell(row=r, column=2 + i)
        c.alignment = WRAP_TOP
        c.border = BOX
        c.font = Font(size=9)

# ───────────────────────── 상세 ─────────────────────────
D = sheet('detail')
D['A1'] = '검수 건 상세'
D['A1'].font = Font(bold=True, size=16, color=NAVY)
D['B2'] = '신청번호'
D['B2'].font = Font(bold=True)
D['C2'].protection = UNLOCK
D['C2'].fill = INPUT_FILL
D['E2'].font = Font(bold=True, color=NAVY)
D.row_dimensions[3].height = 30
for col, w in zip('ABCDEFGHIJK', [3, 18, 32, 32, 24, 22, 22, 22, 14, 14, 14]):
    D.column_dimensions[col].width = w
hint(D, 'F2', '신청번호를 입력하거나 [검수목록]에서 신청번호를 더블클릭하세요.')

# ───────────────────────── 가져오기 ─────────────────────────
I = sheet('import')
I.row_dimensions[1].height = 30
hint(I, 'A2', '① [파일 선택]으로 업체가 보낸 엑셀(여러 개 가능)을 고릅니다 → ② 판정·사유를 확인하고 칸을 직접 고친 뒤 [다시 검사] → '
     '③ \'포함\'을 예/아니오로 정하고 [가져오기 확정]')
hint(I, 'A3', '빨간 칸 = 오류, 노란 칸 = 경고, 주황 칸 = 충돌. 신청번호가 있는 행은 기존 건 수정(변경 내용이 사유에 표시)입니다.')
IMP = [('포함', 7), ('판정', 7), ('사유', 50), ('파일', 22), ('행', 5), ('신청번호', 14), ('업체명', 14), ('프로젝트명', 16),
       ('검수 품목', 11), ('품명/품번', 26), ('수량', 7), ('검수 희망일', 15), ('희망 시간', 9), ('검수 장소', 13),
       ('업체 담당자', 11), ('연락처', 14), ('이메일', 24), ('비고', 20), ('파일 경로', 10)]
header(I, 4, [h for h, _ in IMP], [w for _, w in IMP])
for r in range(5, 1005):
    I.cell(row=r, column=12).number_format = DATE_FMT
    I.cell(row=r, column=13).number_format = '@'
    I.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical='top')
dv(I, '=list_yesno', 'A5:A1004', '포함', '예 또는 아니오')
I.column_dimensions['S'].hidden = True
I.freeze_panes = 'D5'

# ───────────────────────── 메일 ─────────────────────────
M = sheet('mail')
M.row_dimensions[1].height = 30
hint(M, 'A2', '[검수목록]에서 행을 고르고 [선택 건 메일 작성] → 여기서 내용 확인·수정 → 행을 고르고 [본문 복사]/[제목 복사]/[받는 사람 복사] → '
     '웹메일에 붙여넣어 보내기 → [보냄으로 기록]')
hint(M, 'A3', '확정·접수 안내는 같은 업체를 한 통으로 묶습니다. 결과 메일을 보냄으로 기록하면 \'결과 통보일\'이 자동으로 들어갑니다.')
MAILH = [('번호', 5), ('종류', 12), ('받는 사람', 26), ('참조', 18), ('제목', 36), ('본문', 70), ('신청번호', 20), ('기록', 18), ('종류키', 8)]
header(M, 4, [h for h, _ in MAILH], [w for _, w in MAILH])
for r in range(5, 205):
    for c in range(1, 10):
        M.cell(row=r, column=c).alignment = WRAP_TOP
M.column_dimensions['I'].hidden = True
M.freeze_panes = 'A5'

# ───────────────────────── 기록 시트들 ─────────────────────────
def log_sheet(key, heads, fmts=None):
    ws = sheet(key)
    header(ws, 1, [h for h, _ in heads], [w for _, w in heads])
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{CL(len(heads))}1'
    for col, fmt in (fmts or {}).items():
        for r in range(2, 2001):
            ws.cell(row=r, column=col).number_format = fmt
    return ws


log_sheet('nego', [('일시', 17), ('신청번호', 15), ('내용', 70), ('작성자', 12)], {1: STAMP_FMT})
log_sheet('att', [('일시', 17), ('신청번호', 15), ('종류', 10), ('파일 (누르면 열림)', 36), ('전체 경로', 50), ('추가한 사람', 12)], {1: STAMP_FMT})
ml = log_sheet('maillog', [('일시', 17), ('종류키', 10), ('종류', 14), ('신청번호', 28), ('받는 사람', 28), ('제목', 50), ('기록자', 12)], {1: STAMP_FMT})
ml.column_dimensions['B'].hidden = True
log_sheet('hist', [('일시', 17), ('사용자', 12), ('신청번호', 15), ('항목', 14), ('이전 값', 24), ('새 값', 24), ('사유', 40)], {1: STAMP_FMT})
log_sheet('orig', [('신청번호', 15), ('등록일시', 17), ('업체명', 14), ('프로젝트명', 16), ('검수 품목', 11), ('품명/품번', 26), ('수량', 7),
                   ('희망일', 15), ('희망시간', 9), ('신청 장소', 13), ('업체 담당자', 11), ('연락처', 14), ('이메일', 24), ('비고', 20),
                   ('원본 파일·행', 30)], {2: STAMP_FMT, 8: DATE_FMT})

V = sheet('vendor')
header(V, 1, ['업체명', '담당자', '연락처', '이메일', '메모', '최근 신청일'], [18, 12, 15, 28, 30, 15])
V.freeze_panes = 'A2'
for r in range(2, 501):
    V.cell(row=r, column=6).number_format = DATE_FMT
    V.cell(row=r, column=3).number_format = '@'

P = sheet('insp')
header(P, 1, ['이름', '부서', '이메일', 'Windows 계정', '사용'], [12, 14, 28, 16, 8])
P.freeze_panes = 'A2'
hint(P, 'G1', '검수자를 한 줄에 한 명씩 등록하세요. 이름은 [검수목록]의 검수자 목록에 나옵니다.')
hint(P, 'G2', 'Windows 계정을 적으면, 그 계정으로 이 파일을 열 때 배정된 검수 알림 창이 뜹니다.')
hint(P, 'G3', '그만둔 검수자는 지우지 말고 \'사용\'을 \'아니오\'로 바꾸세요 (이전 기록 유지).')
dv(P, '=list_yesno', 'E2:E201', '사용', '예 또는 아니오')
for r in range(2, 202):
    P.cell(row=r, column=5).value = None

# ───────────────────────── 메일양식 ─────────────────────────
T_ = sheet('tpl')
header(T_, 1, ['종류키', '종류', '제목', '본문'], [12, 14, 46, 90])
for i, (key, label, subj, body) in enumerate(TEMPLATES, start=2):
    T_.cell(row=i, column=1, value=key).fill = AUTO_FILL
    T_.cell(row=i, column=2, value=label)
    T_.cell(row=i, column=3, value=subj).alignment = WRAP_TOP
    T_.cell(row=i, column=4, value=body).alignment = WRAP_TOP
    T_.row_dimensions[i].height = 200
T_.column_dimensions['A'].hidden = True
T_['F1'] = '쓸 수 있는 치환값'
T_['F1'].font = Font(bold=True)
for i, v in enumerate(PLACEHOLDERS.values(), start=2):
    T_.cell(row=i, column=6, value='{' + v + '}')
T_.column_dimensions['F'].width = 16
hint(T_, 'G1', '제목·본문은 자유롭게 고쳐도 됩니다. {목록}에는 건마다 한 줄씩 들어가고, 같은 업체 확정·접수는 한 통으로 묶입니다.')

# ───────────────────────── 설정 ─────────────────────────
S = sheet('set')
header(S, 1, ['항목', '값', '설명'], [26, 30, 70])
SETTINGS = [
    ('set_prefix', '신청번호 앞글자', 'IN', '신청번호 형식: 앞글자-연도-순번 (예: IN-2026-0001)'),
    ('set_maxPerDay', '하루 최대 검수 건수', 3, '같은 업체가 같은 장소로 오는 건은 1건으로 셉니다. 0 = 검사 안 함'),
    ('set_maxPerInsp', '검수자 1인당 하루 최대', 2, '0 = 검사 안 함'),
    ('set_slotMin', '같은 장소 겹침 기준(분)', 30, '같은 장소에서 이 시간 안에 다른 업체 검수가 있으면 경고. 0 = 검사 안 함'),
    ('set_maxPerVendor', '같은 업체 하루 최대', 2, '0 = 검사 안 함'),
    ('set_maxQty', '하루 최대 검수 수량', 0, '0 = 검사 안 함'),
    ('set_weekend', '주말 검수 경고', '예', '예/아니오'),
    ('set_allowForce', '충돌 시 강제 진행 허용', '예', '예 = 사유를 적고 확정 가능, 아니오 = 확정 막음'),
    ('set_resultDays', '결과 통보 기한(일)', 7, '검수일 다음날부터 이 일수 안에 업체에 결과를 통보해야 합니다'),
    ('set_remindDays', '검수자 알림: 며칠 앞까지', 1, '0 = 오늘 검수만, 1 = 오늘·내일'),
    ('set_negoLate', '협의 지연 표시(일)', 3, '일정협의중이 이 일수 이상이면 처리 필요에 표시'),
    ('set_popup', '파일 열 때 검수자 알림 창', '예', '예/아니오. [검수자] 시트의 Windows 계정 또는 이름으로 찾습니다'),
    ('set_sender', '보내는 사람 이름(메일 서명)', '', '예: 품질팀 홍길동'),
    ('set_cc', '기본 참조(CC)', '', '업체 메일에 넣을 참조 주소. 여러 개는 ; 로 구분'),
    ('set_webmail', '웹메일 주소', '', '[웹메일 열기] 버튼이 여는 주소. 예: https://mail.company.co.kr'),
    ('set_backupKeep', '백업 보관 개수', 50, '저장할 때마다 [백업] 폴더에 복사본을 남기고, 이 개수를 넘으면 오래된 것부터 지웁니다'),
]
yesno_cells = []
for i, (nm, label, val, desc) in enumerate(SETTINGS, start=2):
    S.cell(row=i, column=1, value=label).font = Font(bold=True)
    c = S.cell(row=i, column=2, value=val)
    c.fill = INPUT_FILL
    c.border = BOX
    S.cell(row=i, column=3, value=desc).font = Font(color='55607A', size=9)
    name(nm, f"{q('set')}!$B${i}")
    if val in ('예', '아니오'):
        yesno_cells.append(f'B{i}')
dv(S, '=list_yesno', ' '.join(yesno_cells), '예/아니오', '예 또는 아니오')

S['E1'] = '검수 품목'
S['G1'] = '휴무일(검수 불가일)'
S['H1'] = '이름'
S['J1'] = 'Windows 계정'
S['K1'] = '표시 이름'
for c in ('E1', 'G1', 'H1', 'J1', 'K1'):
    S[c].fill = HEAD_FILL
    S[c].font = HEAD_FONT
S.column_dimensions['E'].width = 16
S.column_dimensions['G'].width = 16
S.column_dimensions['H'].width = 14
S.column_dimensions['J'].width = 16
S.column_dimensions['K'].width = 14
for i, it in enumerate(['TANK', 'BODY', 'VMB_TURN-KEY', 'UNITTURN-KEY', "Part's"], start=2):
    S.cell(row=i, column=5, value=it)
for r in range(2, 102):
    S.cell(row=r, column=7).number_format = 'yyyy-mm-dd'
hint(S, 'E53', '검수 품목: 한 칸에 하나씩. 휴무일: 날짜와 이름(선택). 사용자: Windows 계정별로 이력에 남길 이름(비우면 엑셀 사용자 이름).')
name('set_items', f"{q('set')}!$E$2:$E$51")
name('set_holidays', f"{q('set')}!$G$2:$H$101")
name('set_users', f"{q('set')}!$J$2:$K$51")
name('list_items', f"OFFSET({q('set')}!$E$2,0,0,MAX(1,COUNTA({q('set')}!$E$2:$E$51)),1)")
name('list_insp', f"OFFSET({q('insp')}!$A$2,0,0,MAX(1,COUNTA({q('insp')}!$A$2:$A$201)),1)")

# ───────────────────────── 사용안내 ─────────────────────────
H = sheet('help')
H.column_dimensions['A'].width = 120
HELP = open(__file__.replace('build_workbook.py', 'help.txt'), encoding='utf-8').read().splitlines()
for i, line in enumerate(HELP, start=1):
    c = H.cell(row=i, column=1, value=line)
    c.alignment = Alignment(wrap_text=True, vertical='top')
    if line.startswith('■'):
        c.font = Font(bold=True, size=12, color=NAVY)
H['A1'].font = Font(bold=True, size=16, color=NAVY)

# ───────────────────────── _text (hidden) ─────────────────────────
X = wb.create_sheet('_text')
X['A1'] = 'key'
X['B1'] = 'text'
for i, (k, v) in enumerate(sorted(TEXT.items()), start=2):
    X.cell(row=i, column=1, value=k)
    X.cell(row=i, column=2, value=v.replace('\n', '\\n'))
X['D1'] = 'status'
for i, v in enumerate(STATUS.values(), start=2):
    X.cell(row=i, column=4, value=v)
X['E1'] = 'result'
X['E2'] = '합격'
X['E3'] = '불합격'
X['F1'] = 'yesno'
X['F2'] = '예'
X['F3'] = '아니오'
X.sheet_state = 'veryHidden'
name('list_status', "'_text'!$D$2:$D$7")
name('list_result', "'_text'!$E$2:$E$3")
name('list_yesno', "'_text'!$F$2:$F$3")

# tab colors
for key, color in [('home', NAVY), ('list', NAVY), ('cal', '2563EB'), ('detail', '2563EB'), ('import', '059669'), ('mail', '059669'),
                   ('set', '9CA3AF'), ('insp', '9CA3AF'), ('vendor', '9CA3AF'), ('tpl', '9CA3AF'), ('help', 'F59E0B')]:
    wb[SHEETS[key]].sheet_properties.tabColor = color

wb.active = 0
wb.save(OUT)
print('saved', OUT)
