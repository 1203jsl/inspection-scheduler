import { describe, it, expect, beforeAll, afterAll, beforeEach, afterEach } from 'vitest';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { AddressInfo } from 'node:net';
import { SMTPServer } from 'smtp-server';
import type { Inspection, Inspector, MailLogEntry } from '../src/shared/types';
import { deadlineLabel, resultLine, resultNotice, resultTargets } from '../src/shared/reminders';
import { DataStore } from '../src/main/core/dataStore';
import { ReminderService } from '../src/main/core/reminders';
import { SmtpCredentialStore, type Crypto } from '../src/main/core/secrets';
import { migrate } from '../src/main/core/schema';

const TODAY = '2026-11-17'; // 화
const NOW = new Date(2026, 10, 17, 8, 30);
const P: Inspector[] = [
  { id: 'p1', name: '김민준', dept: '', email: 'mj@co.example.com', active: true, createdAt: '' },
  { id: 'p2', name: '이서연', dept: '', email: 'sy@co.example.com', active: true, createdAt: '' },
];
let n = 0;
function make(over: Partial<Inspection> & { date?: string } = {}): Inspection {
  const { date = '2026-11-10', ...rest } = over;
  const req = {
    vendorName: '대양산업', category: 'TANK', itemName: '필터', spec: '', quantity: 1, requestedDate: date, requestedTime: '10:00',
    location: 'A', contactName: '', contactPhone: '', contactEmail: 'j@d.example.com', note: '',
  };
  return {
    ...req, id: `IN-2026-${String(++n).padStart(4, '0')}`, status: 'confirmed', original: { ...req },
    schedule: { date, time: '10:00', location: 'A', inspectorId: 'p1' },
    confirmedAt: null, confirmedBy: null, sources: [], negotiations: [], result: null, attachments: [], history: [],
    createdAt: '', createdBy: '', ...rest,
  };
}
const done = (i: Inspection, verdict: 'pass' | 'fail', at = '2026-11-12T01:00:00.000Z'): Inspection => ({
  ...i, status: 'done',
  result: { verdict, date: i.schedule.date!, inspectorId: 'p2', failReason: verdict === 'fail' ? '공차' : '', memo: '', at, by: '' },
  history: [...i.history, { at, by: '', byName: '', action: 'status', field: 'status', from: 'confirmed', to: 'done' }],
});
const sent = (i: Inspection, kind: string, at = '2026-11-13T01:00:00.000Z', result = 'manual') =>
  ({ inspectionId: i.id, kind, at, result }) as unknown as MailLogEntry;

describe('결과 통보 기한 (검수일 다음날부터 7일)', () => {
  it('검수일이 지나야 대상, 기한 = 검수일 + 7일', () => {
    expect(resultNotice(make({ date: TODAY }), [], TODAY, 7)).toBeNull(); // 오늘 검수 → 아직
    const r = resultNotice(make({ date: '2026-11-10' }), [], TODAY, 7)!;
    expect(r).toMatchObject({ inspected: '2026-11-10', deadline: '2026-11-17', daysLeft: 0, stage: 'noResult', inspectorId: 'p1' });
    expect(deadlineLabel(r)).toBe('기한 11-17(화) · 오늘까지');
    expect(deadlineLabel(resultNotice(make({ date: '2026-11-16' }), [], TODAY, 7)!)).toBe('기한 11-23(월) · 6일 남음');
    expect(deadlineLabel(resultNotice(make({ date: '2026-11-05' }), [], TODAY, 7)!)).toBe('기한 11-12(목) · 5일 초과');
  });

  it('결과를 입력했어도 업체 통보 메일을 안 보냈으면 대상, 보내면(수동 기록 포함) 끝', () => {
    const d = done(make({ date: '2026-11-11' }), 'fail');
    const r = resultNotice(d, [], TODAY, 7)!;
    expect(r.stage).toBe('notSent');
    expect(r.inspectorId).toBe('p2'); // 결과에 적힌 검수자
    expect(resultLine(d, r)).toContain('결과 입력됨(불합격) · 통보 메일 미발송');
    expect(resultNotice(d, [sent(d, 'fail')], TODAY, 7)).toBeNull();
    expect(resultNotice(d, [sent(d, 'fail', '2026-11-13T01:00:00.000Z', 'failed')], TODAY, 7)).not.toBeNull(); // 발송 실패는 안 침
    expect(resultNotice(d, [sent(d, 'pass')], TODAY, 7)).not.toBeNull(); // 다른 종류 메일은 안 침
  });

  it('취소·협의중 건과 확정 회신 메일만 보낸 건', () => {
    expect(resultNotice(make({ status: 'canceled' }), [], TODAY, 7)).toBeNull();
    expect(resultNotice(make({ status: 'negotiating' }), [], TODAY, 7)).toBeNull();
    const c = make();
    expect(resultNotice(c, [sent(c, 'confirmed')], TODAY, 7)).not.toBeNull(); // 확정 회신은 결과 통보가 아님
  });

  it('검수자별 묶음, 기한 초과·임박 순', () => {
    const items = [make({ date: '2026-11-15' }), make({ date: '2026-11-05' }), make({ date: '2026-11-12', schedule: { date: '2026-11-12', time: '10:00', location: 'A', inspectorId: 'p2' } })];
    const t = resultTargets(items, P, [], TODAY, 7);
    expect(t.map((x) => [x.inspector.name, x.items.map((y) => y.n.daysLeft)])).toEqual([['김민준', [-5, 5]], ['이서연', [2]]]);
  });

  it('설정 구조 v8: 결과 통보 알림 켜짐·7일 기본, 기존 알림 설정은 유지', () => {
    const v = migrate('settings', { schemaVersion: 7, revision: 1, mailTemplates: {}, mail: {}, conflictRules: {}, notify: { daysBefore: 2, popup: false, mail: true, mailPreview: false } } as any);
    expect(v.notify).toEqual({ daysBefore: 2, popup: false, mail: true, mailPreview: false, resultAlert: true, resultDays: 7 });
    expect(v.resultReminderTemplate.subject).toContain('결과 통보 요청');
  });
});

describe('결과 통보 요청 메일 (시험용 메일 서버)', () => {
  const CRYPTO: Crypto = { available: () => true, encrypt: (s) => Buffer.from(s), decrypt: (b) => b.toString() };
  const HONG = { account: 'hong.gd', displayName: '홍길동' };
  let server: SMTPServer;
  let port = 0;
  const got: { to: string; subject: string }[] = [];
  beforeAll(async () => {
    server = new SMTPServer({ authOptional: true, allowInsecureAuth: true, disabledCommands: ['STARTTLS'], logger: false,
      onAuth: (_a, _s, cb) => cb(null, { user: 'x' }),
      onData(st, ses, cb) { let raw = ''; st.on('data', (c) => (raw += c)); st.on('end', () => {
        const subj = /Subject: (.*)/.exec(raw)?.[1] ?? '';
        got.push({ to: ses.envelope.rcptTo[0].address, subject: subj }); cb();
      }); } });
    await new Promise<void>((r) => server.listen(0, '127.0.0.1', () => r()));
    port = (server.server.address() as AddressInfo).port;
  });
  afterAll(() => new Promise<void>((r) => server.close(() => r())));
  let dir: string, sec: string;
  beforeEach(async () => {
    got.length = 0;
    dir = await fs.mkdtemp(path.join(os.tmpdir(), 'rn-'));
    sec = await fs.mkdtemp(path.join(os.tmpdir(), 'sec-'));
    await DataStore.initFolder(dir, HONG);
    const s = new DataStore(dir, HONG);
    const st = await s.load('settings');
    await s.save({ settings: { ...st, mail: { ...st.mail, method: 'smtp', smtp: { host: '127.0.0.1', port, security: 'none' } } } });
    const p = await s.load('inspectors');
    await s.save({ inspectors: { ...p, items: P } });
    const i = await s.load('inspections');
    await s.save({ inspections: { ...i, items: [make({ date: '2026-11-10' }), done(make({ date: '2026-11-12' }), 'pass'), make({ date: '2026-11-30' })] } });
    await new SmtpCredentialStore(sec, CRYPTO).save('hong@co.example.com', 'pw');
  });
  afterEach(async () => { for (const d of [dir, sec]) await fs.rm(d, { recursive: true, force: true }); });

  it('담당 검수자에게 하루 한 번 보내고, 같은 날 다시 열어도 안 보냄', async () => {
    const svc = () => new ReminderService(new DataStore(dir, HONG), new SmtpCredentialStore(sec, CRYPTO), () => HONG, () => NOW);
    const r = await svc().auto();
    expect(r).toMatchObject({ status: 'done', sent: ['mj@co.example.com', 'sy@co.example.com'] });
    expect(got.map((g) => g.to)).toEqual(['mj@co.example.com', 'sy@co.example.com']);
    const log = (await new DataStore(dir, HONG).load('mail_log')).items.filter((m) => m.result === 'sent');
    expect(log.map((m) => [m.kind, m.inspectorId])).toEqual([['resultReminder', 'p1'], ['resultReminder', 'p2']]);
    expect(log[0].body).toContain('기한 11-17(화) · 오늘까지 · 결과 미입력');
    expect(log[1].body).toContain('결과 입력됨(합격) · 통보 메일 미발송');
    expect(await svc().auto()).toEqual({ status: 'none' });
    expect(got).toHaveLength(2);
  });

  it('설정에서 끄면 보내지 않음', async () => {
    const s = new DataStore(dir, HONG);
    const st = await s.load('settings');
    await s.save({ settings: { ...st, notify: { ...st.notify, resultAlert: false } } });
    const r = await new ReminderService(new DataStore(dir, HONG), new SmtpCredentialStore(sec, CRYPTO), () => HONG, () => NOW).auto();
    expect(r).toEqual({ status: 'none' });
  });
});
