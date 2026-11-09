/**
 * 테스트·샘플용 엑셀 파일을 만듭니다.
 * scripts/make-samples.ts 가 같은 함수로 samples/ 폴더의 파일을 만듭니다.
 */
import ExcelJS from 'exceljs';
import { buildTemplate } from '../src/main/core/excel/template';
import { TEMPLATE_V1, DEFAULT_INSPECTION_ITEMS } from '../src/main/core/schema';

type Cell = string | number | Date | null;
const D = (y: number, m: number, d: number) => new Date(Date.UTC(y, m - 1, d));

/** 우리가 보낸 빈 양식에 업체가 3행부터 적어 회신한 파일 */
export async function vendorReply(rows: Cell[][], today = new Date(2026, 10, 9)): Promise<Buffer> {
  const buf = await buildTemplate(TEMPLATE_V1, DEFAULT_INSPECTION_ITEMS, today);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf as unknown as ArrayBuffer);
  const ws = wb.getWorksheet(TEMPLATE_V1.sheetName)!;
  rows.forEach((r, i) => r.forEach((v, j) => (ws.getCell(3 + i, j + 1).value = v)));
  return Buffer.from(await wb.xlsx.writeBuffer());
}

/** 양식을 쓰지 않고 업체가 직접 만든 파일 (열 순서·머리글이 다름) */
export async function freeformFile(headers: string[], rows: Cell[][]): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.addRow(['ABC 업체 검수 요청']); // 제목 행
  ws.addRow(headers);
  rows.forEach((r) => ws.addRow(r));
  return Buffer.from(await wb.xlsx.writeBuffer());
}

// 열 순서: 신청번호 업체명 검수품목 품명 품번/규격 수량 희망일 희망시간 장소 담당자 연락처 이메일 비고

/** 정상 신규 3건 */
export const NORMAL_ROWS: Cell[][] = [
  [null, '한솔정밀', '기계부품', '베이스 플레이트', 'BP-2205 / AL6061', 10, D(2026, 12, 1), '10:30', '수입검사실 A', '김철수', '010-1111-2222', 'kim@hansol.example.com', ''],
  [null, '미래기전', '전장품', '커넥터 하네스', 'CH-0330 / 12P', 120, D(2026, 12, 2), '10:00', '수입검사실 B', '이영희', '010-3333-4444', 'lee@mirae.example.com', '포장 개봉 검수'],
  [null, '청우소재', '소모품', 'PFA 튜브', 'PT-0610 / Φ6', 200, D(2026, 12, 3), '11:00', '수입검사실 A', '박민수', '031-555-6666', 'park@cheongwoo.example.com', null],
];

/** 오류·경고가 섞인 파일 (행마다 하나씩) */
export const MIXED_ROWS: Cell[][] = [
  /* 3 정상 */ [null, '대양산업', '기계부품', '가스 배관 어셈블리', 'GP-3350', 40, D(2026, 12, 4), '09:30', '조립동 검사장', '정대양', '010-7777-8888', 'jung@daeyang.example.com', ''],
  /* 4 날짜 형식 */ [null, '대양산업', '기계부품', '필터 엘리먼트', 'FE-100', 5, '2026.13.40', '09:30', '조립동 검사장', '정대양', '', 'jung@daeyang.example.com', ''],
  /* 5 품명 누락 */ [null, '우진전장', '전장품', null, 'CP-0872 / 24V', 5, D(2026, 12, 5), '14:00', '수입검사실 B', '최우진', '', 'choi@woojin.example.com', ''],
  /* 6 수량 형식 */ [null, '우진전장', '전장품', '제어 패널', 'CP-1000', '열 개', D(2026, 12, 5), '15:00', '수입검사실 B', '최우진', '', 'choi@woojin.example.com', ''],
  /* 7 이메일 형식 */ [null, '서원테크', '소모품', 'O-링 세트', 'OR-0415 / FKM', 500, D(2026, 12, 7), '13:30', '수입검사실 A', '한서원', '', 'han@seowon', ''],
  /* 8 과거 날짜(경고) */ [null, '서원테크', '소모품', '가스켓', 'GK-22', 50, D(2026, 11, 2), '13:30', '수입검사실 A', '한서원', '', 'han@seowon.example.com', ''],
  /* 9 모르는 품목(경고) */ [null, '한솔정밀', '가공품', '샤프트', 'SF-10', 8, D(2026, 12, 8), '10:00', '수입검사실 A', '김철수', '', 'kim@hansol.example.com', ''],
  /* 10 시간 형식 */ [null, '한솔정밀', '기계부품', '브래킷', 'BR-7', 30, D(2026, 12, 8), '25:00', '수입검사실 A', '김철수', '', 'kim@hansol.example.com', ''],
  /* 11 없는 신청번호 */ ['IN-2026-0099', '미래기전', '전장품', '센서 케이블', 'SC-5M', 20, D(2026, 12, 9), '10:00', '수입검사실 B', '이영희', '', 'lee@mirae.example.com', ''],
  /* 12 문자 날짜·오후 시간 → 정상 */ [null, '청우소재', '소모품', 'PFA 피팅', 'PF-1/4', 100, '2026/12/10', '오후 2시', '수입검사실 A', '박민수', '', 'park@cheongwoo.example.com', ''],
  /* 13 필수 여러 개 누락 */ [null, '청우소재', null, '테플론 시트', null, null, null, null, null, null, null, null, '담당자 확인 중'],
];

/** 정상 파일을 가져온 뒤(IN-2026-0001~0003) 업체가 보낸 수정본 */
export const REVISION_ROWS: Cell[][] = [
  ['IN-2026-0001', '한솔정밀', '기계부품', '베이스 플레이트', 'BP-2205 / AL6061', 12, D(2026, 12, 2), '10:30', '수입검사실 A', '김철수', '010-1111-2222', 'kim@hansol.example.com', '수량 변경'],
  ['in-2026-0003', '청우소재', '소모품', 'PFA 튜브', 'PT-0610 / Φ6', 200, D(2026, 12, 3), '11:00', '수입검사실 B', '박민수', '031-555-6666', 'park@cheongwoo.example.com', null],
  ['IN-2026-0002', '미래기전', '전장품', '커넥터 하네스', 'CH-0330 / 12P', 120, D(2026, 12, 2), '10:00', '수입검사실 B', '이영희', '010-3333-4444', 'lee@mirae.example.com', '포장 개봉 검수'],
  [null, '한솔정밀', '기계부품', '커버', 'CV-01', 4, D(2026, 12, 11), '16:00', '수입검사실 A', '김철수', '010-1111-2222', 'kim.new@hansol.example.com', '추가 신청'],
];
