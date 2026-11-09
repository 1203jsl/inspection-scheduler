import ExcelJS from 'exceljs';
import type { ExcelTemplateDef, FieldKey, Settings } from '../../../shared/types';
import { AppError } from '../errors';
import { templateDef } from '../schema';
import { cellText, normHeader } from './cells';
import { EXAMPLE_MARK, META_SHEET, META_VERSION_LABEL } from './template';

export interface RawRow {
  rowNumber: number;
  /** 열 이름별 원래 셀 값 */
  values: Partial<Record<FieldKey, unknown>>;
}

export interface ReadResult {
  templateVersion?: string;
  sheetName: string;
  rows: RawRow[];
  /** 찾지 못한 선택 열 등 */
  warnings: string[];
}

/** 머리글 행을 찾아 열 이름 → 열 번호 표를 만듭니다. 열은 위치가 아니라 이름으로 찾습니다. */
function findHeader(ws: ExcelJS.Worksheet, def: ExcelTemplateDef) {
  let best: { row: number; map: Partial<Record<FieldKey, number>>; hits: number } | null = null;
  const maxRow = Math.min(ws.rowCount, 10);
  for (let r = 1; r <= maxRow; r++) {
    const row = ws.getRow(r);
    const map: Partial<Record<FieldKey, number>> = {};
    let hits = 0;
    row.eachCell({ includeEmpty: false }, (cell, colNo) => {
      const h = normHeader(cellText(cell.value));
      if (!h) return;
      for (const c of def.columns) {
        if (map[c.key]) continue;
        const names = [c.header, ...(c.aliases ?? [])].map(normHeader);
        if (names.includes(h)) {
          map[c.key] = colNo;
          hits++;
          break;
        }
      }
    });
    if (!best || hits > best.hits) best = { row: r, map, hits };
  }
  // 머리글 4개 이상 맞으면 양식으로 봅니다 (빠진 필수 열은 따로 알려줌)
  return best && best.hits >= Math.min(4, Math.ceil(def.columns.length / 2)) ? best : null;
}

/** 업체가 보낸 엑셀 파일을 읽어 행 목록을 돌려줍니다. */
export async function readWorkbook(data: Buffer | ArrayBuffer, settings: Settings): Promise<ReadResult> {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(data as ArrayBuffer);
  } catch {
    throw new AppError('UNKNOWN', '엑셀 파일을 열 수 없습니다. .xlsx 형식인지 확인하세요. (.xls 는 엑셀에서 "다른 이름으로 저장 > Excel 통합 문서"로 바꾼 뒤 넣으세요)');
  }

  // 숨김 시트의 양식 버전
  let version: string | undefined;
  const meta = wb.getWorksheet(META_SHEET);
  if (meta && cellText(meta.getCell('A1').value) === META_VERSION_LABEL) version = cellText(meta.getCell('A2').value) || undefined;
  const def = templateDef(settings, version);

  // 머리글이 있는 첫 시트
  let ws: ExcelJS.Worksheet | undefined;
  let header: ReturnType<typeof findHeader> = null;
  for (const sheet of wb.worksheets) {
    if (sheet.name === META_SHEET) continue;
    const h = findHeader(sheet, def);
    if (h) {
      ws = sheet;
      header = h;
      break;
    }
  }
  if (!ws || !header) {
    throw new AppError('UNKNOWN', `검수 신청 양식의 머리글을 찾을 수 없습니다. 우리가 보낸 양식(${def.columns.map((c) => c.header).join(', ')})에 작성했는지 확인하세요.`);
  }

  const missingRequired = def.columns.filter((c) => c.required && !header!.map[c.key]).map((c) => `'${c.header}'`);
  if (missingRequired.length) {
    throw new AppError('UNKNOWN', `필수 열 ${missingRequired.join(', ')}을(를) 찾을 수 없습니다. 양식의 머리글 이름을 바꾸지 말고 다시 작성해 달라고 업체에 요청하세요.`);
  }
  const warnings = def.columns
    .filter((c) => !c.required && !header!.map[c.key] && c.key !== 'id')
    .map((c) => `'${c.header}' 열이 없어 빈 값으로 읽습니다.`);
  if (!header.map.id) warnings.push("'신청번호' 열이 없어 모든 행을 신규로 봅니다.");

  const rows: RawRow[] = [];
  for (let r = header.row + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const values: Partial<Record<FieldKey, unknown>> = {};
    let any = false;
    for (const c of def.columns) {
      const colNo = header.map[c.key];
      if (!colNo) continue;
      const v = row.getCell(colNo).value;
      values[c.key] = v;
      if (cellText(v) !== '') any = true;
    }
    if (!any) continue; // 빈 행
    if (cellText(values.id) === EXAMPLE_MARK) continue; // 예시 행
    rows.push({ rowNumber: r, values });
  }
  return { templateVersion: version, sheetName: ws.name, rows, warnings };
}
