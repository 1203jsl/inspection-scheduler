import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { Inspection, Inspector } from '../src/shared/types';
import { can, checkResult, nextAction, recordResult, statusBadge } from '../src/shared/workflow';
import { kindForStatus, pendingMail } from '../src/shared/mail';
import { DataStore } from '../src/main/core/dataStore';
import { AttachmentService } from '../src/main/core/attachments';

const HONG = { account: 'hong.gd', displayName: '홍길동' };
const LEE = { account: 'lee.sy', displayName: '이서연' };
const NOW = new Date(2026, 10, 13, 15, 0);
const TODAY = '2026-11-13';
const INSPECTORS: Inspector[] = [
  { id: 'p1', name: '김민준', dept: '', email: '', active: true, createdAt: '' },
  { id: 'p2', name: '이서연', dept: '', email: '', active: false, createdAt: '' },
];

function make(over: Partial<Inspection> = {}): Inspection {
  const req = {
    vendorName: '대양산업', category: 'TANK', itemName: '필터 엘리먼트', spec: 'FE-100', quantity: 40,
    requestedDate: '2026-11-10', requestedTime: '10:00', location: 'A', contactName: '정대양', contactPhone: '', contactEmail: 'j@d.example.com', note: '',
  };
  return {
    ...req, id: 'IN-2026-0002', status: 'confirmed', original: { ...req },
    schedule: { date: '2026-11-10', time: '10:00', location: 'A', inspectorId: 'p1' },
    confirmedAt: '2026-11-09T01:00:00.000Z', confirmedBy: '홍길동', sources: [], negotiations: [], result: null, attachments: [], history: [],
    createdAt: '', createdBy: '', ...over,
  };
}
const ok = { verdict: 'pass' as const, date: '2026-11-10', inspectorId: 'p1', failReason: '', memo: '' };

describe('검수 결과 입력', () => {
  it('검수일이 되면 결과 입력이 다음 할 일', () => {
    expect(nextAction(make(), TODAY)).toBe('result');
    expect(nextAction(make({ schedule: { ...make().schedule, date: '2026-11-20' } }), TODAY)).toBeNull();
    expect(can(make({ status: 'negotiating' }), 'result')).toBe(false);
  });

  it('합격: 상태가 검수완료·합격이 되고 이력이 남으며, 합격 통보 메일이 미발송으로 잡힌다', () => {
    const r = recordResult(make(), { ...ok, memo: ' 이상 없음 ' }, INSPECTORS, HONG, NOW, TODAY);
    expect(r.status).toBe('done');
    expect(r.result).toMatchObject({ verdict: 'pass', date: '2026-11-10', inspectorId: 'p1', failReason: '', memo: '이상 없음', by: '홍길동' });
    expect(statusBadge(r)).toEqual({ cls: 'st-pass', label: '검수완료 · 합격' });
    expect(r.history.map((h) => [h.field, h.from, h.to])).toEqual([
      ['result.verdict', null, '합격'], ['result.date', null, '2026-11-10'], ['result.inspectorId', null, 'p1'],
      ['result.memo', null, '이상 없음'], ['status', 'confirmed', 'done'],
    ]);
    expect(kindForStatus(r)).toBe('pass');
    expect(pendingMail(r, [])).toBe('pass');
  });

  it('불합격은 사유 필수, 배지는 불합격·통보 메일은 불합격', () => {
    expect(() => recordResult(make(), { ...ok, verdict: 'fail' }, INSPECTORS, HONG, NOW, TODAY)).toThrow('불합격 사유를 입력하세요');
    const r = recordResult(make(), { ...ok, verdict: 'fail', failReason: '외경 공차 초과' }, INSPECTORS, HONG, NOW, TODAY);
    expect(statusBadge(r).cls).toBe('st-fail');
    expect(kindForStatus(r)).toBe('fail');
    expect(r.result?.failReason).toBe('외경 공차 초과');
  });

  it('결과 수정: 바뀐 항목만 "결과 수정" 사유로 남고, 합격으로 바꾸면 불합격 사유는 지워진다', () => {
    const f = recordResult(make(), { ...ok, verdict: 'fail', failReason: '외경 공차 초과' }, INSPECTORS, HONG, NOW, TODAY);
    const p = recordResult(f, { ...ok, failReason: '외경 공차 초과' }, INSPECTORS, HONG, NOW, TODAY);
    const added = p.history.slice(f.history.length);
    expect(added.map((h) => [h.field, h.from, h.to, h.reason])).toEqual([
      ['result.verdict', '불합격', '합격', '결과 수정'],
      ['result.failReason', '외경 공차 초과', null, '결과 수정'],
    ]);
    expect(p.result?.failReason).toBe('');
  });

  it('점검: 판정·검수일·검수자 필수, 미래 날짜는 경고', () => {
    const c = checkResult({ verdict: null, date: null, inspectorId: null, failReason: '', memo: '' }, INSPECTORS, TODAY);
    expect(Object.keys(c.errors).sort()).toEqual(['date', 'inspectorId', 'verdict']);
    expect(checkResult({ ...ok, date: '2026-11-20' }, INSPECTORS, TODAY).warnings[0]).toContain('오늘보다 뒤입니다');
    // 사용 안 함 검수자도 결과에는 적을 수 있음 (이미 검수한 사람)
    expect(checkResult({ ...ok, inspectorId: 'p2' }, INSPECTORS, TODAY).errors).toEqual({});
  });
});

describe('첨부파일', () => {
  let dir: string;
  let src: string;
  beforeEach(async () => {
    dir = await fs.mkdtemp(path.join(os.tmpdir(), 'att-'));
    src = await fs.mkdtemp(path.join(os.tmpdir(), 'src-'));
    await DataStore.initFolder(dir, HONG);
    const s = new DataStore(dir, HONG);
    const f = await s.load('inspections');
    await s.save({ inspections: { ...f, items: [make()] } });
    await fs.writeFile(path.join(src, '치수측정.xlsx'), 'data1');
    await fs.writeFile(path.join(src, '회의록.pdf'), 'data2');
  });
  afterEach(async () => {
    await fs.rm(dir, { recursive: true, force: true });
    await fs.rm(src, { recursive: true, force: true });
  });

  it('첨부/{신청번호}/ 에 복사하고 종류·등록자·등록일과 이력을 남긴다, 같은 이름은 (2)', async () => {
    const store = new DataStore(dir, HONG);
    const svc = new AttachmentService(store, () => HONG, () => NOW);
    await svc.add('IN-2026-0002', [path.join(src, '치수측정.xlsx')], '검수자료');
    const added = await svc.add('IN-2026-0002', [path.join(src, '치수측정.xlsx'), path.join(src, '회의록.pdf')], '회의록');
    expect(added.map((a) => a.file)).toEqual(['첨부/IN-2026-0002/치수측정 (2).xlsx', '첨부/IN-2026-0002/회의록.pdf']);
    const i = (await store.load('inspections')).items[0];
    expect(i.attachments.map((a) => [a.kind, a.name, a.addedBy, a.size])).toEqual([
      ['검수자료', '치수측정.xlsx', '홍길동', 5], ['회의록', '치수측정.xlsx', '홍길동', 5], ['회의록', '회의록.pdf', '홍길동', 5],
    ]);
    expect(i.history.filter((h) => h.action === 'attach').map((h) => h.to)).toEqual(['치수측정.xlsx (검수자료)', '치수측정 (2).xlsx (회의록)', '회의록.pdf (회의록)']);
    expect(await fs.readFile(path.join(dir, '첨부/IN-2026-0002/치수측정 (2).xlsx'), 'utf8')).toBe('data1');
  });

  it('삭제하면 목록에서 빠지고 이력이 남으며, 파일은 _삭제됨 폴더에 보관', async () => {
    const store = new DataStore(dir, HONG);
    const svc = new AttachmentService(store, () => HONG, () => NOW);
    const [a] = await svc.add('IN-2026-0002', [path.join(src, '회의록.pdf')], '회의록');
    await svc.remove('IN-2026-0002', a.id);
    const i = (await store.load('inspections')).items[0];
    expect(i.attachments).toEqual([]);
    expect(i.history.at(-1)).toMatchObject({ action: 'detach', from: '회의록.pdf (회의록)', byName: '홍길동' });
    await expect(fs.access(path.join(dir, a.file))).rejects.toThrow();
    expect(await fs.readdir(path.join(dir, '첨부/IN-2026-0002/_삭제됨'))).toEqual(['20261113-150000_회의록.pdf']);
    await expect(svc.remove('IN-2026-0002', a.id)).rejects.toThrow('이미 삭제된 첨부파일');
  });

  it('다른 사람이 그 사이 저장했어도 그 변경을 덮어쓰지 않고 첨부를 더한다', async () => {
    const me = new DataStore(dir, HONG);
    const svc = new AttachmentService(me, () => HONG, () => NOW);
    await me.load('inspections');
    const other = new DataStore(dir, LEE);
    const f = await other.load('inspections');
    await other.save({ inspections: { ...f, items: f.items.map((i) => ({ ...i, note: '이서연이 고침' })) } });
    await svc.add('IN-2026-0002', [path.join(src, '회의록.pdf')], '회의록');
    const i = (await me.load('inspections')).items[0];
    expect(i.note).toBe('이서연이 고침');
    expect(i.attachments).toHaveLength(1);
  });

  it('없는 파일·폴더는 거부하고 복사한 파일도 남기지 않는다', async () => {
    const svc = new AttachmentService(new DataStore(dir, HONG), () => HONG, () => NOW);
    await expect(svc.add('IN-2026-0002', [path.join(src, '회의록.pdf'), path.join(src, '없음.pdf')], '기타')).rejects.toThrow('파일을 찾을 수 없습니다: 없음.pdf');
    expect(await fs.readdir(path.join(dir, '첨부/IN-2026-0002'))).toEqual([]);
    await expect(svc.add('IN-2026-0002', [src], '기타')).rejects.toThrow('폴더는 첨부할 수 없습니다');
    await expect(svc.add('../x', [path.join(src, '회의록.pdf')], '기타')).rejects.toThrow('신청번호가 올바르지 않습니다');
  });
});
