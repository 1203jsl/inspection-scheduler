import ExcelJS from 'exceljs';
import type { ExcelTemplateDef, FieldKey } from '../../../shared/types';

/** 숨김 시트: 양식 버전과 드롭다운 목록 */
export const META_SHEET = '_meta';
export const META_VERSION_LABEL = 'TEMPLATE_VERSION';
/** 예시 행의 신청번호 칸 (가져올 때 건너뜀) */
export const EXAMPLE_MARK = '예시';
/** 데이터 입력 칸 수 */
export const DATA_ROWS = 300;

const EXAMPLE: Partial<Record<FieldKey, string | number | Date>> = {
  id: EXAMPLE_MARK,
  vendorName: '한솔정밀',
  itemName: '베이스 플레이트',
  spec: 'BP-2205 / AL6061',
  quantity: 10,
  requestedTime: '10:30',
  location: '수입검사실 A',
  contactName: '김철수',
  contactPhone: '010-1234-5678',
  contactEmail: 'kim@hansol.example.com',
  note: '예시 행입니다. 3행부터 작성하세요.',
};

const GUIDE = [
  ['검수 신청서 작성 안내'],
  [''],
  ['1. [검수신청] 시트의 3행부터 검수 1건당 한 행씩 작성합니다. 2행은 작성 예시입니다.'],
  ['2. 신규 신청은 [신청번호]를 비워 두세요. 등록 후 신청번호를 메일로 알려 드립니다.'],
  ['3. 이미 받은 신청번호의 일정·수량 등을 고칠 때는 [신청번호]를 적고 바뀐 내용을 입력해 회신하세요.'],
  ['4. [검수 품목]은 칸을 누르면 나오는 목록에서 고르세요.'],
  ['5. [검수 희망일]은 2026-11-14 처럼 날짜로, [희망 시간]은 13:30 처럼 24시간 형식으로 적습니다.'],
  ['6. [수량]은 1 이상의 숫자만 적습니다 (단위 제외).'],
  ['7. [업체 담당자]와 [이메일]은 일정 확정·결과 통보 메일을 받을 분으로 적어 주세요.'],
  ['8. 머리글(1행)과 열 순서는 바꾸지 마세요.'],
  [''],
  ['* 표시가 있는 항목은 필수입니다.'],
];

/** 빈 양식(.xlsx)을 만듭니다. */
export async function buildTemplate(def: ExcelTemplateDef, items: string[], today = new Date()): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = '검수일정 관리';
  wb.created = today;

  const ws = wb.addWorksheet(def.sheetName, { views: [{ state: 'frozen', ySplit: 1 }] });
  const guide = wb.addWorksheet('작성 안내');
  const meta = wb.addWorksheet(META_SHEET, { state: 'veryHidden' });

  // 숨김 시트: A1 버전, B열 검수 품목
  meta.getCell('A1').value = META_VERSION_LABEL;
  meta.getCell('A2').value = def.version;
  const list = items.length ? items : ['(설정에서 검수 품목을 등록하세요)'];
  list.forEach((it, i) => (meta.getCell(i + 1, 2).value = it));

  ws.columns = def.columns.map((c) => ({ header: c.header + (c.required ? ' *' : ''), key: c.key, width: c.width ?? 14 }));

  // 머리글
  const head = ws.getRow(1);
  head.height = 22;
  head.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2F63' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.protection = { locked: true };
  });

  // 예시 행
  const ex = new Date(Date.UTC(today.getFullYear(), today.getMonth(), today.getDate() + 7));
  const exampleRow = ws.getRow(2);
  def.columns.forEach((c, i) => {
    const cell = exampleRow.getCell(i + 1);
    cell.value = c.key === 'category' ? list[0] : c.key === 'requestedDate' ? ex : EXAMPLE[c.key] ?? '';
    cell.font = { italic: true, color: { argb: 'FF55607A' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F5F9' } };
    cell.protection = { locked: true };
  });

  // 입력 칸: 잠금 해제 + 유효성 검사
  const col = (k: FieldKey) => def.columns.findIndex((c) => c.key === k) + 1;
  const cCat = col('category'), cDate = col('requestedDate'), cQty = col('quantity'), cTime = col('requestedTime');
  const listRange = `${META_SHEET}!$B$1:$B$${list.length}`;
  const first = 3, last = 3 + DATA_ROWS - 1;
  for (let r = first; r <= last; r++) {
    const row = ws.getRow(r);
    for (let c = 1; c <= def.columns.length; c++) row.getCell(c).protection = { locked: false };
    if (cDate > 0) row.getCell(cDate).numFmt = 'yyyy-mm-dd';
    if (cTime > 0) row.getCell(cTime).numFmt = '@';
  }
  // 유효성 검사는 열 범위로 한 번만 지정 (칸마다 넣으면 범위가 겹쳐 엑셀이 '복구' 경고를 띄움)
  const dv = (colNo: number, rule: object) => {
    if (colNo <= 0) return;
    const L = ws.getColumn(colNo).letter;
    (ws as any).dataValidations.add(`${L}${first}:${L}${last}`, rule);
  };
  dv(cCat, {
    type: 'list', allowBlank: true, formulae: [listRange], showErrorMessage: true,
    errorStyle: 'warning', errorTitle: '검수 품목', error: '목록에 있는 검수 품목을 고르세요.',
  });
  dv(cDate, {
    type: 'date', operator: 'greaterThan', allowBlank: true, formulae: [new Date(Date.UTC(2020, 0, 1))],
    showErrorMessage: true, errorStyle: 'stop', errorTitle: '검수 희망일', error: '날짜로 입력하세요. 예: 2026-11-14',
    showInputMessage: true, promptTitle: '검수 희망일', prompt: '예: 2026-11-14',
  });
  dv(cQty, {
    type: 'whole', operator: 'greaterThanOrEqual', allowBlank: true, formulae: [1],
    showErrorMessage: true, errorStyle: 'stop', errorTitle: '수량', error: '1 이상의 숫자만 입력하세요.',
  });
  dv(cTime, {
    type: 'custom', allowBlank: true, formulae: ['TRUE'],
    showInputMessage: true, promptTitle: '희망 시간', prompt: '24시간 형식. 예: 13:30',
  });
  await ws.protect('', {
    selectLockedCells: true, selectUnlockedCells: true, formatColumns: true, formatRows: true,
    insertRows: false, deleteRows: false, sort: false, autoFilter: false,
  });

  guide.getColumn(1).width = 100;
  GUIDE.forEach((line, i) => {
    const cell = guide.getCell(i + 1, 1);
    cell.value = line[0];
    if (i === 0) cell.font = { bold: true, size: 14 };
  });

  return Buffer.from(await wb.xlsx.writeBuffer());
}
