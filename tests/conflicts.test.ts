import { describe, it, expect } from 'vitest';
import type { ConflictRules, Inspection, Inspector, PreviewRow } from '../src/shared/types';
import { conflictMap, conflictsFor, DEFAULT_CONFLICT_RULES, type Slot } from '../src/shared/conflicts';
import { checkForce, confirm, change } from '../src/shared/workflow';
import { importConflicts } from '../src/main/core/importConflicts';
import { defaultSettings, migrate } from '../src/main/core/schema';

const NAMES: Record<string, string> = { p1: '김민준', p2: '이서연' };
const nameOf = (id: string) => NAMES[id] ?? id;
const rules = (over: Partial<ConflictRules> = {}): ConflictRules => ({ ...structuredClone(DEFAULT_CONFLICT_RULES), ...over });
const OFF = { enabled: false, value: 99 };
const only = (k: keyof ConflictRules, v: any): ConflictRules => rules({
  maxPerDay: OFF, maxPerInspector: OFF, maxPerVendor: OFF, maxQtyPerDay: OFF,
  sameSlot: { enabled: false, minutes: 60 }, holidays: { enabled: false, dates: [] }, [k]: v,
});

let n = 0;
const slot = (o: Partial<Slot>): Slot => ({
  id: `IN-2026-${String(++n).padStart(4, '0')}`, date: '2026-11-10', time: '10:00', location: '수입검사실 A',
  inspectorId: null, vendorName: '한솔정밀', quantity: 10, status: 'confirmed', ...o,
});
const msgs = (cs: { message: string }[]) => cs.map((c) => c.message);

describe('충돌 규칙별 검사', () => {
  it('하루 최대 검수 건수: 넘으면 그날 모든 건이 걸림, 취소 건은 세지 않음', () => {
    const r = only('maxPerDay', { enabled: true, value: 3 });
    const day = [slot({ time: '09:00' }), slot({ time: '11:00' }), slot({ time: '13:00' }), slot({ time: '15:00', status: 'canceled' })];
    expect(conflictMap(day, r, nameOf).size).toBe(0);
    const t = slot({ time: '16:00' });
    expect(msgs(conflictsFor(t, day, r, nameOf))).toEqual(['하루 최대 검수 건수(3건) 초과 — 11-10(화) 4건']);
    expect(conflictsFor({ ...t, date: '2026-11-11' }, day, r, nameOf)).toEqual([]);
  });

  it('자기 자신의 기존 일정은 빼고 계산 (같은 날 안에서 시간만 바꿔도 충돌로 안 봄)', () => {
    const r = only('maxPerDay', { enabled: true, value: 2 });
    const a = slot({ time: '09:00' }), b = slot({ time: '11:00' });
    expect(conflictsFor({ ...a, time: '14:00' }, [a, b], r, nameOf)).toEqual([]);
  });

  it('검수자 1인당 하루 최대 건수', () => {
    const r = only('maxPerInspector', { enabled: true, value: 2 });
    const day = [slot({ inspectorId: 'p1' }), slot({ inspectorId: 'p1', time: '13:00' }), slot({ inspectorId: 'p2' })];
    expect(msgs(conflictsFor(slot({ inspectorId: 'p1', time: '15:00' }), day, r, nameOf))).toEqual(['검수자 1인당 하루 최대 2건 초과 — 김민준 11-10(화) 3건']);
    expect(conflictsFor(slot({ inspectorId: 'p2', time: '15:00' }), day, r, nameOf)).toEqual([]);
    expect(conflictsFor(slot({ inspectorId: null, time: '15:00' }), day, r, nameOf)).toEqual([]); // 미배정은 대상 아님
  });

  it('같은 시간대·같은 장소: 시작 시각 차이가 설정 분 미만이면 중복', () => {
    const r = only('sameSlot', { enabled: true, minutes: 60 });
    const a = slot({ id: 'IN-A', time: '10:00', location: '수입검사실 A' });
    expect(msgs(conflictsFor(slot({ id: 'IN-B', time: '10:30' }), [a], r, nameOf))).toEqual(['같은 시간대·같은 장소 중복 — 11-10(화) 수입검사실 A 10:00(IN-A) / 10:30(IN-B)']);
    expect(conflictsFor(slot({ time: '11:00' }), [a], r, nameOf)).toEqual([]); // 60분 차이는 허용
    expect(conflictsFor(slot({ time: '10:00', location: '수입검사실 B' }), [a], r, nameOf)).toEqual([]);
    expect(conflictsFor(slot({ time: null }), [a], r, nameOf)).toEqual([]); // 시간 모르면 판단 안 함
  });

  it('하루 최대 검수 수량 합계', () => {
    const r = only('maxQtyPerDay', { enabled: true, value: 1000 });
    const day = [slot({ quantity: 600 }), slot({ quantity: 300, time: '13:00' })];
    expect(conflictsFor(slot({ quantity: 100, time: '15:00' }), day, r, nameOf)).toEqual([]);
    expect(msgs(conflictsFor(slot({ quantity: 101, time: '15:00' }), day, r, nameOf))).toEqual(['하루 최대 검수 수량(1,000개) 초과 — 11-10(화) 합계 1,001개']);
  });

  it('같은 업체 하루 최대 건수', () => {
    const r = only('maxPerVendor', { enabled: true, value: 2 });
    const day = [slot({}), slot({ time: '13:00' }), slot({ vendorName: '대양산업' })];
    expect(msgs(conflictsFor(slot({ time: '15:00' }), day, r, nameOf))).toEqual(['같은 업체 하루 최대 2건 초과 — 한솔정밀 11-10(화) 3건']);
    expect(conflictsFor(slot({ vendorName: '대양산업', time: '15:00' }), day, r, nameOf)).toEqual([]);
  });

  it('검수 불가일', () => {
    const r = only('holidays', { enabled: true, dates: [{ date: '2026-12-25', name: '성탄절' }] });
    expect(msgs(conflictsFor(slot({ date: '2026-12-25' }), [], r, nameOf))).toEqual(['검수 불가일입니다 — 12-25(금) 성탄절']);
    expect(conflictsFor(slot({ date: '2026-12-24' }), [], r, nameOf)).toEqual([]);
  });

  it('꺼진 규칙은 검사하지 않는다', () => {
    const r = only('holidays', { enabled: false, dates: [{ date: '2026-12-25', name: '성탄절' }] });
    const many = Array.from({ length: 10 }, (_, k) => slot({ time: `${String(8 + k).padStart(2, '0')}:00` }));
    expect(conflictMap(many, r, nameOf).size).toBe(0);
  });

  it('여러 규칙이 함께 걸리면 모두 알려준다', () => {
    const day = [slot({ inspectorId: 'p1', time: '09:00' }), slot({ inspectorId: 'p1', time: '13:00' }), slot({ time: '15:00' })];
    const cs = conflictsFor(slot({ inspectorId: 'p1', time: '13:30' }), day, rules(), nameOf);
    expect(cs.map((c) => c.rule).sort()).toEqual(['maxPerDay', 'maxPerInspector', 'maxPerVendor', 'sameSlot']);
  });
});

describe('강제 진행', () => {
  const INSPECTORS: Inspector[] = [{ id: 'p1', name: '김민준', dept: '', email: '', active: true, createdAt: '' }];
  const insp = {
    id: 'IN-2026-0001', status: 'negotiating', vendorName: '한솔정밀', itemName: 'x', quantity: 1, history: [], negotiations: [], sources: [],
    schedule: { date: '2026-11-10', time: '10:00', location: 'A', inspectorId: 'p1' },
  } as unknown as Inspection;
  const HONG = { account: 'hong.gd', displayName: '홍길동' };

  it('충돌이 없으면 그냥 통과', () => {
    expect(checkForce('IN-1', [], false, '')).toBeUndefined();
  });
  it('강제 진행이 꺼져 있으면 막는다', () => {
    expect(() => checkForce('IN-1', ['하루 최대 초과'], false, '사유')).toThrow('설정에서 강제 진행이 꺼져 있어 진행할 수 없습니다');
  });
  it('허용돼도 사유가 없으면 막는다', () => {
    expect(() => checkForce('IN-1', ['하루 최대 초과'], true, '  ')).toThrow('강제로 진행하려면 사유를 입력하세요');
  });
  it('사유를 넣으면 진행되고, 걸린 충돌과 사유가 변경 이력에 남는다', () => {
    const f = checkForce(insp.id, ['하루 최대 검수 건수(3건) 초과 — 11-10(화) 4건'], true, ' 팀장 승인 ');
    const c = confirm(insp, insp.schedule, INSPECTORS, HONG, new Date(), undefined, f);
    expect(c.history[0]).toMatchObject({ action: 'force', field: 'conflict', to: '하루 최대 검수 건수(3건) 초과 — 11-10(화) 4건', reason: '팀장 승인', byName: '홍길동' });
    expect(c.status).toBe('confirmed');
    const ch = change(c, { ...c.schedule, date: '2026-11-11' }, INSPECTORS, '연기', HONG, new Date(), { conflicts: ['x'], reason: 'y' });
    expect(ch.history.filter((h) => h.action === 'force')).toHaveLength(2);
  });
});

describe('가져오기 충돌 경고', () => {
  const settings = defaultSettings('a', 'A', '');
  const existing = [0, 1, 2].map((k) => ({
    id: `IN-2026-000${k + 1}`, status: 'confirmed', vendorName: `업체${k}`, quantity: 1,
    schedule: { date: '2026-12-01', time: `${10 + k * 2}:00`, location: 'A', inspectorId: null },
  })) as unknown as Inspection[];
  const row = (n: number, over: Partial<PreviewRow['parsed']> = {}, kind: PreviewRow['kind'] = 'new'): PreviewRow => ({
    key: `0:${n}`, fileIndex: 0, fileName: 'a.xlsx', rowNumber: n, kind, target: 'new', display: {} as any, issues: {}, defaultInclude: true,
    parsed: { id: '', vendorName: '새업체', category: '', itemName: '', spec: '', quantity: 1, requestedDate: '2026-12-01', requestedTime: '17:00', location: 'B', contactName: '', contactPhone: '', contactEmail: '', note: '', ...over },
  });

  it('기존 일정과 합쳐 하루 최대를 넘으면 희망일 칸에 경고 (막지 않음)', () => {
    const m = importConflicts([row(3), row(4, { requestedDate: '2026-12-02' })], existing, settings, []);
    expect(m.get('0:3')?.requestedDate?.[0]).toEqual({ level: 'warn', message: '충돌: 하루 최대 검수 건수(3건) 초과 — 12-01(화) 4건 — 가져온 뒤 일정 조정 필요' });
    expect(m.has('0:4')).toBe(false);
  });

  it('가져오는 행끼리도 검사하고, 오류 행은 셈에서 뺀다', () => {
    const m = importConflicts([row(3, { requestedDate: '2026-12-05', time: undefined } as any), row(4, { requestedDate: '2026-12-05', location: 'B', requestedTime: '17:30' }), row(5, { requestedDate: '2026-12-05' }, 'error')], [], settings, []);
    expect(m.get('0:4')?.requestedDate?.map((i) => i.message)).toEqual([
      '충돌: 같은 시간대·같은 장소 중복 — 12-05(토) B 17:00(a.xlsx 3행) / 17:30(a.xlsx 4행) — 가져온 뒤 일정 조정 필요',
    ]);
  });
});

describe('설정 구조 변환', () => {
  it('settings v2 → v3: 충돌 조건 기본값 추가', () => {
    const v2 = { schemaVersion: 2, revision: 5, userMap: {}, lockStaleMinutes: 30, backupKeep: 50, inspectionItems: [], excelTemplates: [], currentTemplate: 'v1' };
    const v3 = migrate('settings', v2);
    expect(v3.schemaVersion).toBeGreaterThanOrEqual(3); // 이후 구조까지 이어서 바뀜
    expect(v3.conflictRules).toEqual(DEFAULT_CONFLICT_RULES);
    expect(v3.conflictRules).not.toBe(DEFAULT_CONFLICT_RULES); // 복사본
  });
});
