import { describe, it, expect, beforeAll, afterAll, beforeEach, afterEach } from 'vitest';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { AddressInfo } from 'node:net';
import ExcelJS from 'exceljs';
import { SMTPServer } from 'smtp-server';
import type { Inspection, Inspector } from '../src/shared/types';
import { matchInspectors, reminderHandled, reminderTargets, upcoming } from '../src/shared/reminders';
import { DataStore } from '../src/main/core/dataStore';
import { ReminderService } from '../src/main/core/reminders';
import { SmtpCredentialStore, type Crypto } from '../src/main/core/secrets';
import { buildListExport, buildVendorSchedule } from '../src/main/core/excel/export';
import { migrate } from '../src/main/core/schema';

const HONG = { account: 'hong.gd', displayName: '홍길동' };
const LEE = { account: 'lee.sy', displayName: '이서연' };
const CRYPTO: Crypto = { available: () => true, encrypt: (s) => Buffer.from(s), decrypt: (b) => b.toString() };
const NOW = new Date(2026, 10, 13, 8, 30); // 11-13(금)
const TODAY = '2026-11-13';
const P: Inspector[] = [
  { id: 'p1', name: '김민준', dept: '', email: 'mj@co.example.com', active: true, createdAt: '', account: 'kim.mj' },
  { id: 'p2', name: '이서연', dept: '', email: 'sy@co.example.com', active: true, createdAt: '' },
  { id: 'p3', name: '정하늘', dept: '', email: 'hn@co.example.com', active: false, createdAt: '' },
];

let n = 0;
function make(over: Partial<Inspection> & { date?: string; time?: string; who?: string | null } = {}): Inspection {
  const { date = '2026-11-14', time = '10:00', who = 'p1', ...rest } = over;
  const req = {
    vendorName: '한솔정밀', category: 'TANK', itemName: '베이스 플레이트', spec: 'BP-2205', quantity: 10,
    requestedDate: date, requestedTime: time, location: '수입검사실 A', contactName: '김철수', contactPhone: '', contactEmail: 'kim@hansol.example.com', note: '',
  };
  return {
    ...req, id: `IN-2026-${String(++n).padStart(4, '0')}`, status: 'confirmed', original: { ...req },
    schedule: { date, time, location: '수입검사실 A', inspectorId: who },
    confirmedAt: null, confirmedBy: null, sources: [], negotiations: [], result: null, attachments: [], history: [],
    createdAt: '', createdBy: '', ...rest,
  };
}

describe('알림 대상', () => {
  const items = [
    make({ date: '2026-11-13', time: '15:00' }),
    make({ date: '2026-11-14', time: '09:00' }),
    make({ date: '2026-11-15' }), // 이틀 뒤 → 1일 전 알림 범위 밖
    make({ date: '2026-11-14', status: 'negotiating' }), // 일정 미확정
    make({ date: '2026-11-14', status: 'canceled' }),
    make({ date: '2026-11-14', who: 'p2', status: 'changed' }),
  ];
  it('확정·변경된 건 중 오늘~N일 뒤, 시간순', () => {
    expect(upcoming(items, TODAY, 1).map((i) => `${i.schedule.date} ${i.schedule.time} ${i.schedule.inspectorId}`)).toEqual([
      '2026-11-13 15:00 p1', '2026-11-14 09:00 p1', '2026-11-14 10:00 p2',
    ]);
    expect(upcoming(items, TODAY, 0, 'p1')).toHaveLength(1);
  });
  it('이 PC 사용자와 검수자 연결: Windows 계정, 계정이 비어 있으면 이름', () => {
    expect(matchInspectors({ account: 'KIM.MJ', displayName: '아무개' }, P).map((p) => p.id)).toEqual(['p1']);
    expect(matchInspectors({ account: 'lee.sy', displayName: '이서연' }, P).map((p) => p.id)).toEqual(['p2']);
    expect(matchInspectors({ account: 'x', displayName: '정하늘' }, P)).toEqual([]); // 사용 안 함
  });
  it('메일 대상: 사용 중·이메일 있는 검수자별 목록', () => {
    expect(reminderTargets(items, P, TODAY, 1).map((t) => [t.inspector.name, t.items.length])).toEqual([['김민준', 2], ['이서연', 1]]);
  });
  it('처리 여부: 보냄=처리, 실패=다시, 오래된 보내는 중=다시', () => {
    const e = (result: string, at: string) => ({ kind: 'reminder', reminderDate: TODAY, inspectorId: 'p1', result, at }) as any;
    expect(reminderHandled([], TODAY, 'p1', NOW)).toBe(false);
    expect(reminderHandled([e('sent', NOW.toISOString())], TODAY, 'p1', NOW)).toBe(true);
    expect(reminderHandled([e('sent', NOW.toISOString()), e('failed', NOW.toISOString())], TODAY, 'p1', NOW)).toBe(false);
    expect(reminderHandled([e('claimed', new Date(NOW.getTime() - 5 * 60_000).toISOString())], TODAY, 'p1', NOW)).toBe(true);
    expect(reminderHandled([e('claimed', new Date(NOW.getTime() - 40 * 60_000).toISOString())], TODAY, 'p1', NOW)).toBe(false);
  });
});

describe('알림 메일 (시험용 메일 서버, PC 두 대)', () => {
  let server: SMTPServer;
  let port = 0;
  const got: string[][] = [];
  beforeAll(async () => {
    server = new SMTPServer({
      authOptional: true, allowInsecureAuth: true, disabledCommands: ['STARTTLS'], logger: false,
      onAuth: (_a, _s, cb) => cb(null, { user: 'x' }),
      onData(st, ses, cb) { st.on('data', () => {}); st.on('end', () => { got.push(ses.envelope.rcptTo.map((r) => r.address)); cb(); }); },
    });
    await new Promise<void>((r) => server.listen(0, '127.0.0.1', () => r()));
    port = (server.server.address() as AddressInfo).port;
  });
  afterAll(() => new Promise<void>((r) => server.close(() => r())));

  let dir: string;
  let sec1: string;
  let sec2: string;
  beforeEach(async () => {
    got.length = 0;
    dir = await fs.mkdtemp(path.join(os.tmpdir(), 'rem-'));
    sec1 = await fs.mkdtemp(path.join(os.tmpdir(), 'sec-'));
    sec2 = await fs.mkdtemp(path.join(os.tmpdir(), 'sec-'));
    await DataStore.initFolder(dir, HONG);
    const s = new DataStore(dir, HONG);
    const st = await s.load('settings');
    await s.save({ settings: { ...st, mail: { ...st.mail, method: 'smtp', smtp: { host: '127.0.0.1', port, security: 'none' } } } });
    const p = await s.load('inspectors');
    await s.save({ inspectors: { ...p, items: P } });
    const i = await s.load('inspections');
    await s.save({ inspections: { ...i, items: [make({ date: '2026-11-14' }), make({ date: '2026-11-13', who: 'p2' }), make({ date: '2026-11-20' })] } });
    await new SmtpCredentialStore(sec1, CRYPTO).save('hong@co.example.com', 'pw');
    await new SmtpCredentialStore(sec2, CRYPTO).save('lee@co.example.com', 'pw');
  });
  afterEach(async () => {
    for (const d of [dir, sec1, sec2]) await fs.rm(d, { recursive: true, force: true });
  });

  it('그날 처음 연 PC가 보내고, 다른 PC는 다시 보내지 않는다', async () => {
    const pc1 = new ReminderService(new DataStore(dir, HONG), new SmtpCredentialStore(sec1, CRYPTO), () => HONG, () => NOW);
    const pc2 = new ReminderService(new DataStore(dir, LEE), new SmtpCredentialStore(sec2, CRYPTO), () => LEE, () => NOW);
    const r1 = await pc1.auto();
    expect(r1).toEqual({ status: 'done', sent: ['mj@co.example.com', 'sy@co.example.com'], failed: [] });
    expect(await pc2.auto()).toEqual({ status: 'none' });
    expect(got).toEqual([['mj@co.example.com'], ['sy@co.example.com']]);
    const log = (await new DataStore(dir, HONG).load('mail_log')).items;
    expect(log.map((m) => `${m.inspectorId}:${m.result}`)).toEqual(['p1:claimed', 'p2:claimed', 'p1:sent', 'p2:sent']);
    expect(log[2]).toMatchObject({ kind: 'reminder', reminderDate: TODAY, inspectorId: 'p1', byName: '홍길동' });
  });

  it('두 PC가 동시에 열어도 한 통씩만 간다', async () => {
    const pc1 = new ReminderService(new DataStore(dir, HONG), new SmtpCredentialStore(sec1, CRYPTO), () => HONG, () => NOW);
    const pc2 = new ReminderService(new DataStore(dir, LEE), new SmtpCredentialStore(sec2, CRYPTO), () => LEE, () => NOW);
    await Promise.all([pc1.auto(), pc2.auto()]);
    expect(got.flat().sort()).toEqual(['mj@co.example.com', 'sy@co.example.com']);
  });

  it('미리보기 설정이면 초안만 돌려주고, 이 PC에 계정이 없으면 아무것도 하지 않는다', async () => {
    const s = new DataStore(dir, HONG);
    const st = await s.load('settings');
    await s.save({ settings: { ...st, notify: { ...st.notify, mailPreview: true } } });
    const pc1 = new ReminderService(new DataStore(dir, HONG), new SmtpCredentialStore(sec1, CRYPTO), () => HONG, () => NOW);
    const r = await pc1.auto();
    expect(r.status).toBe('preview');
    if (r.status === 'preview') {
      expect(r.drafts.map((d) => d.subject)).toEqual([
        '[검수 알림] 김민준 님 배정 검수 1건 (11-13(금)~11-14(토))',
        '[검수 알림] 이서연 님 배정 검수 1건 (11-13(금)~11-14(토))',
      ]);
      expect(r.drafts[0].body).toContain('- 11-14(토) 10:00 · 한솔정밀 · 베이스 플레이트 · 수입검사실 A (IN-');
    }
    expect(got).toHaveLength(0);
    const empty = await fs.mkdtemp(path.join(os.tmpdir(), 'sec-'));
    const pc3 = new ReminderService(new DataStore(dir, HONG), new SmtpCredentialStore(empty, CRYPTO), () => HONG, () => NOW);
    expect((await pc3.auto()).status).toBe('notConfigured');
    expect((await new DataStore(dir, HONG).load('mail_log')).items).toHaveLength(0); // 선점도 안 함
  });

  it('수동 방식이거나 알림 메일이 꺼져 있으면 보내지 않는다', async () => {
    const s = new DataStore(dir, HONG);
    const st = await s.load('settings');
    await s.save({ settings: { ...st, notify: { ...st.notify, mail: false } } });
    const pc1 = new ReminderService(new DataStore(dir, HONG), new SmtpCredentialStore(sec1, CRYPTO), () => HONG, () => NOW);
    expect(await pc1.auto()).toEqual({ status: 'off' });
  });
});

describe('엑셀 내보내기', () => {
  const nameOf = (id: string | null) => ({ p1: '김민준', p2: '이서연' } as Record<string, string>)[id ?? ''] ?? '';
  it('목록: 화면 순서 그대로, 상태·결과·불합격 사유 포함', async () => {
    const a = make({ date: '2026-11-14' });
    const b = make({ date: '2026-11-15', status: 'done', result: { verdict: 'fail', date: '2026-11-15', inspectorId: 'p1', failReason: '공차 초과', memo: '', at: '', by: '' } });
    const buf = await buildListExport([b, a], nameOf, { by: '홍길동', at: NOW, filter: '업체: 한솔정밀' });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf as unknown as ArrayBuffer);
    const ws = wb.getWorksheet('검수 목록')!;
    expect(ws.getRow(1).getCell(1).value).toBe('신청번호');
    expect(ws.getRow(2).getCell(1).value).toBe(b.id);
    expect([ws.getRow(2).getCell(11).value, ws.getRow(2).getCell(12).value, ws.getRow(2).getCell(13).value]).toEqual(['검수완료 · 불합격', '불합격', '공차 초과']);
    expect(ws.getRow(3).getCell(7).value).toBe('2026-11-14(토)');
    expect(wb.getWorksheet('정보')!.getCell('B3').value).toBe('업체: 한솔정밀');
  });
  it('업체별 일정 확인서: 업체·기간·작성자와 일정표, 취소 건 표시', async () => {
    const items = [make({ date: '2026-11-14' }), make({ date: '2026-11-20', status: 'canceled' })];
    const buf = await buildVendorSchedule('한솔정밀', items, nameOf, { from: '2026-11-01', to: '2026-11-30', by: '홍길동', today: TODAY });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf as unknown as ArrayBuffer);
    const ws = wb.getWorksheet('일정 확인서')!;
    expect(ws.getCell('A1').value).toBe('검수 일정 확인서');
    expect(ws.getCell('B3').value).toBe('한솔정밀');
    expect(ws.getCell('B4').value).toBe('2026-11-01(일) ~ 2026-11-30(월)');
    expect(ws.getCell('B6').value).toBe('홍길동');
    expect(ws.getRow(8).getCell(1).value).toBe('신청번호');
    expect(ws.getRow(9).getCell(8).value).toBe('김민준');
    expect(ws.getRow(10).getCell(10).value).toBe('취소됨');
    expect(ws.getCell('A12').value).toBe('총 2건');
  });
});

describe('설정 구조 변환 v7', () => {
  it('알림 설정 기본값: 1일 전, 팝업·메일 켜짐, 미리보기 없이', () => {
    const v = migrate('settings', { schemaVersion: 6, revision: 1, mailTemplates: {}, mail: {}, conflictRules: {} } as any);
    expect(v.notify).toEqual({ daysBefore: 1, popup: true, mail: true, mailPreview: false });
    expect(v.reminderTemplate.subject).toContain('{검수자}');
  });
});
