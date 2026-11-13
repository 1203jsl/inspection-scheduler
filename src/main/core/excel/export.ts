import ExcelJS from 'exceljs';
import type { Inspection } from '../../../shared/types';
import { fmtYMD } from '../../../shared/dates';
import { statusBadge } from '../../../shared/workflow';

type NameOf = (id: string | null) => string;

const HEAD_FILL: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2F63' } };
const thin: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: 'FFD5DBE7' } }, bottom: { style: 'thin', color: { argb: 'FFD5DBE7' } },
  left: { style: 'thin', color: { argb: 'FFD5DBE7' } }, right: { style: 'thin', color: { argb: 'FFD5DBE7' } },
};

function styleHeader(row: ExcelJS.Row) {
  row.height = 22;
  row.eachCell((c) => {
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = HEAD_FILL;
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thin;
  });
}

const verdict = (i: Inspection) => (i.result ? (i.result.verdict === 'pass' ? '합격' : '불합격') : '');

/** 목록 내보내기 (화면의 현재 필터·정렬 그대로) */
export async function buildListExport(items: Inspection[], nameOf: NameOf, meta: { by: string; at: Date; filter: string }): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = '검수일정 관리';
  const ws = wb.addWorksheet('검수 목록', { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = [
    { header: '신청번호', key: 'id', width: 15 },
    { header: '업체명', key: 'vendor', width: 14 },
    { header: '검수 품목', key: 'cat', width: 13 },
    { header: '품명', key: 'item', width: 22 },
    { header: '품번/규격', key: 'spec', width: 20 },
    { header: '수량', key: 'qty', width: 8 },
    { header: '검수일', key: 'date', width: 14 },
    { header: '시간', key: 'time', width: 8 },
    { header: '장소', key: 'loc', width: 14 },
    { header: '검수자', key: 'insp', width: 10 },
    { header: '상태', key: 'status', width: 16 },
    { header: '결과', key: 'result', width: 8 },
    { header: '불합격 사유', key: 'fail', width: 30 },
    { header: '업체 담당자', key: 'contact', width: 11 },
    { header: '연락처', key: 'phone', width: 15 },
    { header: '이메일', key: 'email', width: 26 },
    { header: '희망일', key: 'req', width: 14 },
    { header: '비고', key: 'note', width: 24 },
  ];
  styleHeader(ws.getRow(1));
  for (const i of items) {
    const r = ws.addRow({
      id: i.id, vendor: i.vendorName, cat: i.category, item: i.itemName, spec: i.spec, qty: i.quantity ?? '',
      date: fmtYMD(i.schedule.date), time: i.schedule.time ?? '', loc: i.schedule.location, insp: nameOf(i.schedule.inspectorId),
      status: statusBadge(i).label, result: verdict(i), fail: i.result?.failReason ?? '',
      contact: i.contactName, phone: i.contactPhone, email: i.contactEmail, req: fmtYMD(i.requestedDate), note: i.note,
    });
    r.eachCell({ includeEmpty: true }, (c) => (c.border = thin));
    r.getCell('qty').alignment = { horizontal: 'right' };
  }
  ws.autoFilter = { from: 'A1', to: { row: 1, column: ws.columns.length } };
  const info = wb.addWorksheet('정보');
  info.getColumn(1).width = 14;
  info.getColumn(2).width = 60;
  [['내보낸 사람', meta.by], ['내보낸 시각', meta.at.toLocaleString('ko-KR')], ['필터', meta.filter || '전체'], ['건수', `${items.length}건`]]
    .forEach(([k, v], n) => { info.getCell(n + 1, 1).value = k; info.getCell(n + 1, 1).font = { bold: true }; info.getCell(n + 1, 2).value = v; });
  return Buffer.from(await wb.xlsx.writeBuffer());
}

/** 업체별 일정 확인서 (업체에 보내거나 출력용) */
export async function buildVendorSchedule(
  vendor: string, items: Inspection[], nameOf: NameOf, opt: { from: string; to: string; by: string; today: string },
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = '검수일정 관리';
  const ws = wb.addWorksheet('일정 확인서', { pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
  const cols = [
    { h: '신청번호', w: 15 }, { h: '품명', w: 22 }, { h: '품번/규격', w: 20 }, { h: '수량', w: 8 },
    { h: '검수일', w: 15 }, { h: '시간', w: 8 }, { h: '장소', w: 16 }, { h: '검수자', w: 10 }, { h: '상태', w: 16 }, { h: '비고', w: 26 },
  ];
  cols.forEach((c, k) => (ws.getColumn(k + 1).width = c.w));
  ws.mergeCells(1, 1, 1, cols.length);
  ws.getCell(1, 1).value = '검수 일정 확인서';
  ws.getCell(1, 1).font = { bold: true, size: 18 };
  ws.getCell(1, 1).alignment = { horizontal: 'center' };
  ws.getRow(1).height = 32;
  const metaRows: [string, string][] = [
    ['업체명', vendor],
    ['기간', `${fmtYMD(opt.from)} ~ ${fmtYMD(opt.to)}`],
    ['작성일', fmtYMD(opt.today)],
    ['작성자', opt.by],
  ];
  metaRows.forEach(([k, v], n) => {
    const r = 3 + n;
    ws.getCell(r, 1).value = k;
    ws.getCell(r, 1).font = { bold: true };
    ws.mergeCells(r, 2, r, 4);
    ws.getCell(r, 2).value = v;
  });
  const headRow = 3 + metaRows.length + 1;
  const head = ws.getRow(headRow);
  cols.forEach((c, k) => (head.getCell(k + 1).value = c.h));
  styleHeader(head);
  items.forEach((i, n) => {
    const r = ws.getRow(headRow + 1 + n);
    const remark = i.status === 'canceled' ? '취소됨' : i.status === 'changed' ? '일정 변경 — 재확정 예정' : i.result?.verdict === 'fail' ? `불합격: ${i.result.failReason}` : '';
    [i.id, i.itemName, i.spec, i.quantity ?? '', fmtYMD(i.schedule.date), i.schedule.time ?? '', i.schedule.location, nameOf(i.schedule.inspectorId), statusBadge(i).label, remark]
      .forEach((v, k) => {
        const c = r.getCell(k + 1);
        c.value = v as ExcelJS.CellValue;
        c.border = thin;
        if (i.status === 'canceled') c.font = { strike: true, color: { argb: 'FF8A2A36' } };
      });
  });
  const end = headRow + 1 + items.length + 1;
  ws.getCell(end, 1).value = items.length ? `총 ${items.length}건` : '이 기간에 잡힌 검수가 없습니다.';
  ws.getCell(end + 1, 1).value = '※ 일정이 맞지 않거나 바뀐 내용이 있으면 신청번호와 함께 회신해 주세요.';
  ws.getCell(end + 1, 1).font = { color: { argb: 'FF55607A' } };
  return Buffer.from(await wb.xlsx.writeBuffer());
}
