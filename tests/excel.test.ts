import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { PreviewRow, Settings } from '../src/shared/types';
import { parseDate, parseQuantity, parseTime } from '../src/main/core/excel/cells';
import { readWorkbook } from '../src/main/core/excel/read';
import { buildTemplate } from '../src/main/core/excel/template';
import { buildPreviewRows, finalizePreview, nextId, normalizeId } from '../src/main/core/importer';
import { ImportService } from '../src/main/core/importService';
import { DataStore } from '../src/main/core/dataStore';
import { AppError } from '../src/main/core/errors';
import { defaultSettings, migrate, TEMPLATE_V1, DEFAULT_INSPECTION_ITEMS } from '../src/main/core/schema';
import { freeformFile, MIXED_ROWS, NORMAL_ROWS, REVISION_ROWS, vendorReply } from './fixtures';

const settings: Settings = defaultSettings('hong.gd', '홍길동', '2026-11-09T00:00:00Z');
const ctx = { settings, existing: [], today: '2026-11-09' };
const HONG = { account: 'hong.gd', displayName: '홍길동' };
const NOW = () => new Date(2026, 10, 9, 14, 5, 0);

async function previewOf(buf: Buffer, c = ctx) {
  const r = await readWorkbook(buf, settings);
  return finalizePreview(buildPreviewRows(r.rows, 0, 'a.xlsx', r.templateVersion, c), c);
}
const issue = (r: PreviewRow, k: string) => r.issues[k as keyof PreviewRow['issues']]?.map((i) => `${i.level}:${i.message}`) ?? [];

describe('셀 값 읽기', () => {
  it('날짜는 날짜셀·일련번호·여러 문자 형식을 받는다', () => {
    expect(parseDate(new Date(Date.UTC(2026, 10, 14)))).toBe('2026-11-14');
    expect(parseDate(46340)).toBe('2026-11-14');
    for (const s of ['2026-11-14', '2026.11.14', '2026/11/14', '2026. 11. 14.', '2026년 11월 14일', '20261114']) {
      expect(parseDate(s)).toBe('2026-11-14');
    }
    expect(parseDate('2026.13.40')).toBeNull();
    expect(parseDate('2026-02-30')).toBeNull();
    expect(parseDate('다음주')).toBeNull();
    expect(parseDate(null)).toBe('');
  });
  it('시간은 13:30, 엑셀 시간, 오후 2시 등을 받는다', () => {
    expect(parseTime('13:30')).toBe('13:30');
    expect(parseTime('9:05')).toBe('09:05');
    expect(parseTime(0.5625)).toBe('13:30');
    expect(parseTime(new Date(Date.UTC(1899, 11, 30, 13, 30)))).toBe('13:30');
    expect(parseTime('오후 2시')).toBe('14:00');
    expect(parseTime('오전 9시 30분')).toBe('09:30');
    expect(parseTime('14시')).toBe('14:00');
    expect(parseTime('25:00')).toBeNull();
    expect(parseTime('')).toBe('');
  });
  it('수량은 1 이상 정수 (쉼표·단위 허용)', () => {
    expect(parseQuantity(10)).toBe(10);
    expect(parseQuantity('1,200')).toBe(1200);
    expect(parseQuantity('20개')).toBe(20);
    expect(parseQuantity('20 EA')).toBe(20);
    expect(parseQuantity(0)).toBeNull();
    expect(parseQuantity(2.5)).toBeNull();
    expect(parseQuantity('열 개')).toBeNull();
  });
});

describe('빈 양식', () => {
  it('머리글·예시 행·작성 안내·숨김 버전·드롭다운·유효성 검사·보호가 들어간다', async () => {
    const buf = await buildTemplate(TEMPLATE_V1, DEFAULT_INSPECTION_ITEMS, new Date(2026, 10, 9));
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf as unknown as ArrayBuffer);
    const ws = wb.getWorksheet('검수신청')!;
    expect(ws.getRow(1).values).toContain('업체명 *');
    expect(ws.getCell('A2').value).toBe('예시');
    expect(wb.getWorksheet('작성 안내')).toBeTruthy();
    const meta = wb.getWorksheet('_meta')!;
    expect(meta.state).toBe('veryHidden');
    expect(meta.getCell('A2').value).toBe('v1');
    expect(meta.getCell('B2').value).toBe('BODY');
    // 파일에 실제 기록된 유효성 검사 범위: 열마다 하나, 겹침 없음 (겹치면 엑셀이 '복구' 경고)
    const zip = await JSZip.loadAsync(buf);
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const ranges = [...xml.matchAll(/<dataValidation [^>]*type="(\w+)"[^>]*sqref="([^"]+)"/g)].map((m) => `${m[2]}:${m[1]}`).sort();
    expect(ranges).toEqual(['C3:C302:list', 'F3:F302:whole', 'G3:G302:date', 'H3:H302:custom']);
    expect(ws.getCell('A1').protection?.locked).not.toBe(false); // 엑셀 기본값이 잠김
    expect(ws.getCell('B3').protection?.locked).toBe(false);
    expect((ws as any).sheetProtection?.sheet).toBe(true);
  });

  it('빈 양식을 그대로 넣으면 0행 (예시 행은 건너뜀)', async () => {
    const r = await readWorkbook(await buildTemplate(TEMPLATE_V1, DEFAULT_INSPECTION_ITEMS), settings);
    expect(r.templateVersion).toBe('v1');
    expect(r.rows).toHaveLength(0);
  });
});

describe('가져오기 판정', () => {
  it('정상 파일은 모두 신규, 경고 없음', async () => {
    const rows = await previewOf(await vendorReply(NORMAL_ROWS));
    expect(rows.map((r) => r.kind)).toEqual(['new', 'new', 'new']);
    expect(rows.every((r) => Object.keys(r.issues).length === 0 && r.defaultInclude)).toBe(true);
    expect(rows[0].parsed).toMatchObject({ quantity: 10, requestedDate: '2026-12-01', requestedTime: '10:30', contactEmail: 'kim@hansol.example.com' });
  });

  it('오류·경고 섞인 파일: 칸 단위 사유', async () => {
    const rows = await previewOf(await vendorReply(MIXED_ROWS));
    const at = (n: number) => rows.find((r) => r.rowNumber === n)!;
    expect(at(3).kind).toBe('new');
    expect(issue(at(4), 'requestedDate')).toEqual(['error:날짜 형식이 올바르지 않습니다. 예: 2026-11-14']);
    expect(issue(at(5), 'itemName')).toEqual(['error:필수 항목 누락: 품명을(를) 입력해야 합니다.']);
    expect(issue(at(6), 'quantity')[0]).toMatch(/^error:수량은 1 이상의 정수/);
    expect(issue(at(7), 'contactEmail')[0]).toMatch(/^error:이메일 형식/);
    expect(at(8).kind).toBe('new'); // 경고만
    expect(issue(at(8), 'requestedDate')[0]).toMatch(/^warn:검수 희망일이 오늘\(2026-11-09\)보다 이전/);
    expect(at(9).kind).toBe('new');
    expect(issue(at(9), 'category')[0]).toMatch(/^warn:설정의 검수 품목 목록에 없는/);
    expect(issue(at(10), 'requestedTime')[0]).toMatch(/^error:시간 형식/);
    expect(at(11).kind).toBe('error');
    expect(issue(at(11), 'id')[0]).toMatch(/^error:없는 신청번호/);
    expect(at(12).kind).toBe('new');
    expect(at(12).parsed).toMatchObject({ requestedDate: '2026-12-10', requestedTime: '14:00' });
    const r13 = at(13);
    for (const k of ['category', 'quantity', 'requestedDate', 'contactName', 'contactEmail']) expect(issue(r13, k)[0]).toMatch(/^error:필수 항목 누락/);
    // 오류 행은 기본 제외지만 포함 가능(막지 않음), 경고 행은 기본 포함
    expect(rows.filter((r) => r.defaultInclude).map((r) => r.rowNumber)).toEqual([3, 8, 9, 12]);
  });

  it('열은 위치가 아니라 머리글 이름으로 찾는다 (순서 바뀜·다른 이름·제목 행)', async () => {
    const buf = await freeformFile(
      ['품명', '업체', '이메일', '담당자', '희망일', '수량', '검수품목'],
      [['샤프트', '한솔정밀', 'kim@hansol.example.com', '김철수', '2026-12-01', 3, 'TANK']],
    );
    const r = await readWorkbook(buf, settings);
    expect(r.templateVersion).toBeUndefined();
    expect(r.warnings.join(' ')).toContain("'신청번호' 열이 없어");
    const rows = buildPreviewRows(r.rows, 0, 'x.xlsx', undefined, ctx);
    expect(rows[0].kind).toBe('new');
    expect(rows[0].parsed).toMatchObject({ itemName: '샤프트', vendorName: '한솔정밀', quantity: 3, requestedDate: '2026-12-01' });
  });

  it('필수 열이 아예 없으면 파일 단위로 알려준다', async () => {
    const buf = await freeformFile(['업체명', '검수 품목', '품명', '수량', '검수 희망일', '업체 담당자'], [['A', 'TANK', 'B', 1, '2026-12-01', 'C']]);
    await expect(readWorkbook(buf, settings)).rejects.toThrow("필수 열 '이메일'");
  });

  it('양식이 아닌 엑셀은 머리글을 찾을 수 없다고 알려준다', async () => {
    const buf = await freeformFile(['가', '나', '다'], [[1, 2, 3]]);
    await expect(readWorkbook(buf, settings)).rejects.toThrow('머리글을 찾을 수 없습니다');
  });
});

describe('신청번호', () => {
  it('연도별 4자리 순번', () => {
    expect(nextId([], 2026)).toBe('IN-2026-0001');
    expect(nextId(['IN-2026-0001', 'IN-2026-0012', 'IN-2025-0300'], 2026)).toBe('IN-2026-0013');
    expect(nextId(['IN-2025-0300'], 2026)).toBe('IN-2026-0001');
    expect(normalizeId(' in-2026-0003 ')).toBe('IN-2026-0003');
  });
});

describe('가져오기 전체 흐름 (폴더에 저장)', () => {
  let dir: string;
  let tmp: string;
  beforeEach(async () => {
    dir = await fs.mkdtemp(path.join(os.tmpdir(), 'imp-'));
    tmp = await fs.mkdtemp(path.join(os.tmpdir(), 'xls-'));
    await DataStore.initFolder(dir, HONG);
  });
  afterEach(async () => {
    await fs.rm(dir, { recursive: true, force: true });
    await fs.rm(tmp, { recursive: true, force: true });
  });
  const put = async (name: string, buf: Buffer) => {
    const p = path.join(tmp, name);
    await fs.writeFile(p, buf);
    return p;
  };

  it('신규 등록 → 신청번호 부여·원본 보관·업체 목록 반영 → 수정본 매칭', async () => {
    const store = new DataStore(dir, HONG, NOW);
    const svc = new ImportService(store, () => HONG, NOW);

    const p1 = await put('검수신청_1109.xlsx', await vendorReply(NORMAL_ROWS));
    const pv = await svc.preview([p1]);
    const sum = await svc.commit(pv.rows.filter((r) => r.defaultInclude).map((r) => r.key));
    expect(sum.created.map((c) => c.id)).toEqual(['IN-2026-0001', 'IN-2026-0002', 'IN-2026-0003']);
    expect(sum.originals).toEqual(['원본/2026-11/한솔정밀외2곳_20261109-140500_검수신청_1109.xlsx']);
    await fs.access(path.join(dir, sum.originals[0]));

    const insp = await store.load('inspections');
    const i1 = insp.items[0];
    expect(i1.status).toBe('registered');
    expect(i1.sources[0]).toMatchObject({ file: sum.originals[0], row: 3, importedBy: '홍길동' });
    expect(i1.history.map((h) => h.action)).toEqual(['created', 'status']);
    const vendors = await store.load('vendors');
    expect(vendors.items.map((v) => v.name)).toEqual(['한솔정밀', '미래기전', '청우소재']);
    expect(vendors.items[0].contactEmail).toBe('kim@hansol.example.com');

    // 수정본
    const p2 = await put('수정본.xlsx', await vendorReply(REVISION_ROWS));
    const pv2 = await svc.preview([p2]);
    const [r1, r3, r2, r4] = pv2.rows;
    expect(r1.kind).toBe('update');
    expect(r1.diff).toEqual({
      quantity: { from: '10', to: '12' },
      requestedDate: { from: '2026-12-01', to: '2026-12-02' },
      note: { from: '(비어 있음)', to: '수량 변경' },
    });
    expect(r3.parsed.id).toBe('IN-2026-0003'); // 소문자로 적어도 매칭
    expect(r3.diff).toEqual({ location: { from: '수입검사실 A', to: '수입검사실 B' } });
    expect(r2.diff).toEqual({});
    expect(r2.defaultInclude).toBe(false); // 바뀐 내용 없음
    expect(issue(r2, 'row')).toEqual(['warn:바뀐 내용이 없습니다.']);
    expect(r4.kind).toBe('new');

    const sum2 = await svc.commit(pv2.rows.filter((r) => r.defaultInclude).map((r) => r.key));
    expect(sum2.updated).toEqual(['IN-2026-0001', 'IN-2026-0003']);
    expect(sum2.created.map((c) => c.id)).toEqual(['IN-2026-0004']);
    expect(sum2.excluded).toBe(1);

    const after = await store.load('inspections');
    const a1 = after.items.find((i) => i.id === 'IN-2026-0001')!;
    expect(a1.quantity).toBe(12);
    expect(a1.original.quantity).toBe(10); // 원래 신청값은 그대로
    expect(a1.sources).toHaveLength(2);
    expect(a1.history.filter((h) => h.action === 'field').map((h) => [h.field, h.from, h.to])).toEqual([
      ['quantity', 10, 12], ['requestedDate', '2026-12-01', '2026-12-02'], ['schedule.date', '2026-12-01', '2026-12-02'], ['note', '', '수량 변경'],
    ]);
    expect(a1.schedule.date).toBe('2026-12-02'); // 확정 전이라 현재 일정도 따라감
    expect((await store.load('vendors')).items[0].contactEmail).toBe('kim.new@hansol.example.com');
  });

  it('오류 행도 사용자가 포함하면 가져온다 (없는 번호는 새 번호로)', async () => {
    const store = new DataStore(dir, HONG, NOW);
    const svc = new ImportService(store, () => HONG, NOW);
    const pv = await svc.preview([await put('mixed.xlsx', await vendorReply(MIXED_ROWS))]);
    const r11 = pv.rows.find((r) => r.rowNumber === 11)!;
    const sum = await svc.commit([r11.key]);
    expect(sum.created).toHaveLength(1);
    const i = (await store.load('inspections')).items[0];
    expect(i.id).toBe('IN-2026-0001');
    expect(i.history[0].reason).toContain('업체가 적은 신청번호 IN-2026-0099는 없는 번호여서 새로 부여');
  });

  it('미리보기 뒤 다른 사람이 가져오기를 했으면 저장하지 않고 원본 복사본도 지운다', async () => {
    const kim = new DataStore(dir, HONG, NOW);
    const svc = new ImportService(kim, () => HONG, NOW);
    const pv = await svc.preview([await put('a.xlsx', await vendorReply(NORMAL_ROWS))]);

    const LEE = { account: 'lee.sy', displayName: '이서연' };
    const lee = new DataStore(dir, LEE, NOW);
    const svc2 = new ImportService(lee, () => LEE, NOW);
    const pv2 = await svc2.preview([await put('b.xlsx', await vendorReply(NORMAL_ROWS.slice(0, 1)))]);
    await svc2.commit(pv2.rows.map((r) => r.key));

    await kim.loadAll(); // 화면에서 '새로 불러오기'를 눌러도
    let err: AppError | null = null;
    try {
      await svc.commit(pv.rows.map((r) => r.key)); // 미리보기 기준 데이터는 옛것
    } catch (e) {
      err = e as AppError;
    }
    expect(err?.code).toBe('CONFLICT');
    expect(err?.message).toContain('이서연');
    const items = (await kim.load('inspections')).items;
    expect(items.map((i) => i.id)).toEqual(['IN-2026-0001']);
    expect(await fs.readdir(path.join(dir, '원본', '2026-11'))).toHaveLength(1);
  });

  it('여러 파일을 한 번에: 파일별로 원본 보관, 잘못된 파일은 그 파일만 오류', async () => {
    const store = new DataStore(dir, HONG, NOW);
    const svc = new ImportService(store, () => HONG, NOW);
    const bad = await put('메모.xlsx', await freeformFile(['가'], [[1]]));
    const txt = await put('신청.csv', Buffer.from('a,b'));
    const pv = await svc.preview([
      await put('한솔.xlsx', await vendorReply(NORMAL_ROWS.slice(0, 1))),
      bad,
      await put('미래.xlsx', await vendorReply(NORMAL_ROWS.slice(1, 2))),
      txt,
    ]);
    expect(pv.files.map((f) => !!f.error)).toEqual([false, true, false, true]);
    expect(pv.files[3].error?.message).toContain('.xlsx 파일만');
    expect(pv.rows.map((r) => r.key)).toEqual(['0:3', '2:3']);
    const sum = await svc.commit(pv.rows.map((r) => r.key));
    expect(sum.originals).toHaveLength(2);
    expect(sum.created.map((c) => c.id)).toEqual(['IN-2026-0001', 'IN-2026-0002']);
  });

  it('같은 신청번호가 두 번 나오면 둘 다 경고', async () => {
    const store = new DataStore(dir, HONG, NOW);
    const svc = new ImportService(store, () => HONG, NOW);
    let pv = await svc.preview([await put('a.xlsx', await vendorReply(NORMAL_ROWS))]);
    await svc.commit(pv.rows.map((r) => r.key));
    pv = await svc.preview([await put('b.xlsx', await vendorReply([REVISION_ROWS[0], REVISION_ROWS[0]]))]);
    expect(issue(pv.rows[0], 'id')[0]).toMatch(/아래\(b\.xlsx 4행\)에도/);
    expect(issue(pv.rows[1], 'id')[0]).toMatch(/위\(b\.xlsx 3행\)에도/);
  });
});

describe('설정 구조 변환', () => {
  it('1단계에서 만든 settings(v1)는 v2로 바뀌며 양식·품목 기본값이 생긴다', () => {
    const v1 = { schemaVersion: 1, revision: 4, updatedBy: 'a', updatedByName: 'A', updatedAt: '', userMap: {}, lockStaleMinutes: 30, backupKeep: 50 };
    const v2 = migrate('settings', v1);
    expect(v2.schemaVersion).toBe(2);
    expect(v2.revision).toBe(4);
    expect(v2.inspectionItems).toEqual(DEFAULT_INSPECTION_ITEMS);
    expect(v2.excelTemplates[0].version).toBe('v1');
  });
});
