import { describe, it, expect, beforeAll, afterAll, beforeEach, afterEach } from 'vitest';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { AddressInfo } from 'node:net';
import { SMTPServer } from 'smtp-server';
import type { Inspection, MailDraft } from '../src/shared/types';
import { DEFAULT_TEMPLATES, fillTemplate, mailVars, parseAddresses, pendingMail } from '../src/shared/mail';
import { DataStore } from '../src/main/core/dataStore';
import { AppError } from '../src/main/core/errors';
import { appendMailLog, logEntry } from '../src/main/core/mailer';
import { MailService } from '../src/main/core/mailService';
import { SmtpCredentialStore, type Crypto } from '../src/main/core/secrets';
import { migrate } from '../src/main/core/schema';

const HONG = { account: 'hong.gd', displayName: '홍길동' };
const FAKE_CRYPTO: Crypto = {
  available: () => true,
  encrypt: (s) => Buffer.from(s.split('').reverse().join('') + '#enc'),
  decrypt: (b) => b.toString().replace(/#enc$/, '').split('').reverse().join(''),
};

function insp(over: Partial<Inspection> = {}): Inspection {
  const req = {
    vendorName: '한솔정밀', category: 'TANK', itemName: '베이스 플레이트', spec: 'BP-2205', quantity: 10,
    requestedDate: '2026-12-01', requestedTime: '10:30', location: '수입검사실 A',
    contactName: '김철수', contactPhone: '', contactEmail: 'kim@hansol.example.com', note: '',
  };
  return {
    ...req, id: 'IN-2026-0001', status: 'confirmed', original: { ...req },
    schedule: { date: '2026-12-02', time: '14:00', location: '수입검사실 B', inspectorId: 'p1' },
    confirmedAt: null, confirmedBy: null, sources: [], negotiations: [], result: null, attachments: [],
    history: [{ at: '2026-11-09T05:00:00.000Z', by: 'h', byName: '홍길동', action: 'status', field: 'status', from: 'negotiating', to: 'confirmed' }],
    createdAt: '2026-11-08T00:00:00.000Z', createdBy: '홍길동', ...over,
  };
}
const nameOf = (id: string | null) => (id === 'p1' ? '김민준' : '');

describe('템플릿 치환', () => {
  it('확정 템플릿에 검수 건 값이 채워진다', () => {
    const v = mailVars(insp(), nameOf, '홍길동');
    const subj = fillTemplate(DEFAULT_TEMPLATES.confirmed.subject, v);
    expect(subj.text).toBe('[검수 일정 확정 안내] IN-2026-0001 베이스 플레이트 (한솔정밀)');
    const body = fillTemplate(DEFAULT_TEMPLATES.confirmed.body, v).text;
    expect(body).toContain('김철수 님께');
    expect(body).toContain('- 검수 일시: 2026-12-02(수) 14:00');
    expect(body).toContain('- 검수 장소: 수입검사실 B');
    expect(body).toContain('- 담당 검수자: 김민준');
    expect(body).toContain('홍길동 드림');
  });

  it('디자인 표기({시간}{장소})도 받아 주고, 모르는 치환값은 그대로 두고 알려준다', () => {
    const r = fillTemplate('{시간} {장소} {검수 장소} {없는값}', mailVars(insp(), nameOf, '홍길동'));
    expect(r.text).toBe('14:00 수입검사실 B {검수 장소} {없는값}');
    expect(r.unknown).toEqual(['{검수 장소}', '{없는값}']);
  });

  it('취소·변경 사유는 마지막 상태 변경 사유에서 가져온다', () => {
    const i = insp({
      status: 'canceled',
      history: [{ at: 'x', by: '', byName: '', action: 'status', field: 'status', from: 'confirmed', to: 'canceled', reason: '업체 요청' }],
    });
    expect(mailVars(i, nameOf, '').취소사유).toBe('업체 요청');
    expect(mailVars(i, nameOf, '').변경사유).toBe('-');
  });

  it('받는 사람 주소 점검', () => {
    expect(parseAddresses('김철수 <kim@hansol.example.com>, lee@mirae.example.com')).toEqual({ ok: ['김철수 <kim@hansol.example.com>', 'lee@mirae.example.com'], bad: [] });
    expect(parseAddresses('kim@hansol, ').bad).toEqual(['kim@hansol']);
  });
});

describe('미발송 판단', () => {
  const sent = (kind: string, at: string, result = 'sent') =>
    ({ inspectionId: 'IN-2026-0001', kind, at, result }) as any;
  it('확정된 뒤 확정 메일이 없으면 미발송', () => {
    expect(pendingMail(insp(), [])).toBe('confirmed');
    expect(pendingMail(insp(), [sent('confirmed', '2026-11-09T04:00:00.000Z')])).toBe('confirmed'); // 확정 전에 보낸 건 아님
    expect(pendingMail(insp(), [sent('confirmed', '2026-11-09T06:00:00.000Z', 'failed')])).toBe('confirmed'); // 실패는 안 침
    expect(pendingMail(insp(), [sent('confirmed', '2026-11-09T06:00:00.000Z')])).toBeNull();
    expect(pendingMail(insp(), [sent('confirmed', '2026-11-09T06:00:00.000Z', 'manual')])).toBeNull();
  });
  it('일정협의중에는 보낼 메일이 없다', () => {
    expect(pendingMail(insp({ status: 'negotiating' }), [])).toBeNull();
  });
});

describe('비밀번호 저장 (이 PC)', () => {
  let dir: string;
  beforeEach(async () => { dir = await fs.mkdtemp(path.join(os.tmpdir(), 'sec-')); });
  afterEach(async () => { await fs.rm(dir, { recursive: true, force: true }); });

  it('암호화한 값만 파일에 쓰고, 화면용 정보에는 비밀번호가 없다', async () => {
    const s = new SmtpCredentialStore(dir, FAKE_CRYPTO);
    const info = await s.save('hong@company.example', 'Secr3t!pw');
    expect(info).toEqual({ account: 'hong@company.example', hasPassword: true, encryptionAvailable: true });
    const raw = await fs.readFile(path.join(dir, 'smtp-account.json'), 'utf8');
    expect(raw).not.toContain('Secr3t!pw');
    expect((await s.credentials()).password).toBe('Secr3t!pw');
    await s.save('hong@company.example'); // 비밀번호 칸 비워 두면 유지
    expect((await s.credentials()).password).toBe('Secr3t!pw');
    await s.save('hong@company.example', ''); // 지우기
    expect((await s.info()).hasPassword).toBe(false);
  });

  it('암호화를 쓸 수 없는 PC에서는 저장하지 않는다', async () => {
    const s = new SmtpCredentialStore(dir, { ...FAKE_CRYPTO, available: () => false });
    await expect(s.save('a@b.example', 'pw')).rejects.toThrow('비밀번호 암호화를 사용할 수 없어 비밀번호는 저장하지 않았습니다');
    const raw = await fs.readFile(path.join(dir, 'smtp-account.json'), 'utf8');
    expect(raw).not.toContain('"pw"'); // 평문 저장 안 함
    expect(await s.info()).toMatchObject({ account: 'a@b.example', hasPassword: false });
  });
});

describe('SMTP 발송 (시험용 메일 서버)', () => {
  let server: SMTPServer;
  let port = 0;
  const received: { to: string[]; raw: string }[] = [];
  beforeAll(async () => {
    server = new SMTPServer({
      authOptional: false,
      allowInsecureAuth: true,
      disabledCommands: ['STARTTLS'],
      onAuth(auth, _s, cb) {
        if (auth.username === 'hong@company.example' && auth.password === 'right-pw') cb(null, { user: auth.username });
        else cb(Object.assign(new Error('Invalid login'), { responseCode: 535 }));
      },
      onData(stream, session, cb) {
        let raw = '';
        stream.on('data', (c) => (raw += c.toString()));
        stream.on('end', () => {
          received.push({ to: session.envelope.rcptTo.map((r) => r.address), raw });
          cb();
        });
      },
      logger: false,
    });
    await new Promise<void>((r) => server.listen(0, '127.0.0.1', () => r()));
    port = (server.server.address() as AddressInfo).port;
  });
  afterAll(() => new Promise<void>((r) => server.close(() => r())));

  let dir: string;
  let secDir: string;
  let store: DataStore;
  let creds: SmtpCredentialStore;
  const draft = (over: Partial<MailDraft> = {}): MailDraft => ({
    inspectionId: 'IN-2026-0001', kind: 'confirmed', to: '김철수 <kim@hansol.example.com>', cc: 'team@company.example',
    subject: '[검수 일정 확정 안내] IN-2026-0001', body: '본문입니다', attachments: [], ...over,
  });
  beforeEach(async () => {
    received.length = 0;
    dir = await fs.mkdtemp(path.join(os.tmpdir(), 'mail-'));
    secDir = await fs.mkdtemp(path.join(os.tmpdir(), 'sec-'));
    await DataStore.initFolder(dir, HONG);
    store = new DataStore(dir, HONG);
    const s = await store.load('settings');
    await store.save({ settings: { ...s, mail: { ...s.mail, method: 'smtp', smtp: { host: '127.0.0.1', port, security: 'none' } } } });
    creds = new SmtpCredentialStore(secDir, FAKE_CRYPTO);
    await creds.save('hong@company.example', 'right-pw');
  });
  afterEach(async () => {
    await fs.rm(dir, { recursive: true, force: true });
    await fs.rm(secDir, { recursive: true, force: true });
  });

  it('보내고 발송 기록을 남긴다 (첨부 포함, 비밀번호는 기록에 없음)', async () => {
    await fs.mkdir(path.join(dir, '첨부', 'IN-2026-0001'), { recursive: true });
    await fs.writeFile(path.join(dir, '첨부', 'IN-2026-0001', '일정표.txt'), 'hello');
    const svc = new MailService(store, creds, () => HONG);
    const e = await svc.send(draft({ attachments: ['첨부/IN-2026-0001/일정표.txt'] }));
    expect(e).toMatchObject({ result: 'sent', method: 'smtp', inspectionId: 'IN-2026-0001', kind: 'confirmed', byName: '홍길동', attachments: ['일정표.txt'] });
    expect(received).toHaveLength(1);
    expect(received[0].to.sort()).toEqual(['kim@hansol.example.com', 'team@company.example']);
    expect(received[0].raw).toMatch(/Content-Disposition: attachment;\s*filename/);
    expect(received[0].raw).toContain(Buffer.from('hello').toString('base64'));
    const log = await store.load('mail_log');
    expect(log.items).toHaveLength(1);
    const files = (await fs.readdir(path.join(dir, 'data'))).map((f) => fs.readFile(path.join(dir, 'data', f), 'utf8'));
    for (const t of await Promise.all(files)) expect(t).not.toContain('right-pw');
  });

  it('비밀번호가 틀리면 고칠 방법을 알려주고, 실패도 기록한다', async () => {
    await creds.save('hong@company.example', 'wrong');
    const svc = new MailService(store, creds, () => HONG);
    await expect(svc.send(draft())).rejects.toThrow('메일 계정 또는 비밀번호가 맞지 않습니다');
    const log = await store.load('mail_log');
    expect(log.items[0]).toMatchObject({ result: 'failed' });
    expect(log.items[0].error).toContain('비밀번호가 맞지 않습니다');
  });

  it('서버에 연결할 수 없으면 주소·포트·사내 IT 확인을 안내한다', async () => {
    const s = await store.load('settings');
    await store.save({ settings: { ...s, mail: { ...s.mail, smtp: { host: '127.0.0.1', port: 1, security: 'none' } } } });
    const svc = new MailService(store, creds, () => HONG);
    await expect(svc.send(draft())).rejects.toThrow('메일 서버(127.0.0.1:1)에 연결할 수 없습니다');
  });

  it('주소가 잘못됐거나 첨부가 없으면 보내기 전에 막는다', async () => {
    const svc = new MailService(store, creds, () => HONG);
    await expect(svc.send(draft({ to: '' }))).rejects.toThrow('받는 사람 이메일을 입력하세요');
    await expect(svc.send(draft({ cc: 'abc' }))).rejects.toThrow('이메일 형식이 올바르지 않습니다: abc');
    await expect(svc.send(draft({ attachments: ['첨부/없는파일.pdf'] }))).rejects.toThrow('첨부파일을 찾을 수 없습니다');
    await expect(svc.send(draft({ attachments: ['../밖.txt'] }))).rejects.toThrow('데이터 폴더 밖');
    expect(received).toHaveLength(0);
  });

  it('수동 방식이면 자동 발송을 막고, [발송 완료로 기록]은 manual 로 남는다', async () => {
    const s = await store.load('settings');
    await store.save({ settings: { ...s, mail: { ...s.mail, method: 'manual' } } });
    const svc = new MailService(store, creds, () => HONG);
    await expect(svc.send(draft())).rejects.toThrow('발송 방식이 수동입니다');
    const e = await svc.recordManual(draft());
    expect(e).toMatchObject({ method: 'manual', result: 'manual' });
    expect(received).toHaveLength(0);
  });

  it('테스트 메일은 본인 계정으로 간다', async () => {
    const svc = new MailService(store, creds, () => HONG);
    const e = await svc.test();
    expect(e).toMatchObject({ kind: 'test', inspectionId: null, to: 'hong@company.example' });
    expect(received[0].to).toEqual(['hong@company.example']);
  });
});

describe('발송 기록 덧붙이기', () => {
  it('다른 PC가 먼저 기록해도 덮어쓰지 않고 둘 다 남는다', async () => {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'log-'));
    await DataStore.initFolder(dir, HONG);
    const a = new DataStore(dir, HONG);
    const b = new DataStore(dir, { account: 'lee', displayName: '이서연' });
    await a.load('mail_log'); // a 가 먼저 읽어 둔 상태
    const d: MailDraft = { inspectionId: 'X', kind: 'confirmed', to: 'a@b.example', cc: '', subject: 's', body: 'b', attachments: [] };
    await appendMailLog(b, logEntry(d, { account: 'lee', displayName: '이서연' }, 'manual', 'manual', new Date()));
    await appendMailLog(a, logEntry(d, HONG, 'manual', 'manual', new Date()));
    const log = await new DataStore(dir, HONG).load('mail_log');
    expect(log.items.map((m) => m.byName)).toEqual(['이서연', '홍길동']);
    await fs.rm(dir, { recursive: true, force: true });
  });
});

describe('설정 구조 변환', () => {
  it('settings v3 → v4: 메일 템플릿·발송 설정 기본값 (발송 방식 기본은 수동, SMTP 서버는 비워 둠)', () => {
    const v4 = migrate('settings', { schemaVersion: 3, revision: 1, conflictRules: {} } as any);
    expect(v4.schemaVersion).toBeGreaterThanOrEqual(4); // 이후 구조까지 이어서 바뀜
    expect(Object.keys(v4.mailTemplates)).toEqual(['registered', 'confirmed', 'changed', 'canceled', 'pass', 'fail']);
    expect(v4.mail).toEqual({ method: 'manual', smtp: { host: '', port: 587, security: 'starttls' }, webmailUrl: '', defaultCc: '' });
  });
});
