import { describe, it, expect } from 'vitest';
import type { Inspection, Inspector } from '../src/shared/types';
import {
  addNote, assignInspector, can, cancel, change, checkSchedule, confirm, editRequest, nextAction,
  resultMissing, startNegotiation, statusBadge, WorkflowError,
} from '../src/shared/workflow';
import { applyImport } from '../src/main/core/importer';
import { migrate } from '../src/main/core/schema';
import { normalizeTime } from '../src/shared/dates';

const HONG = { account: 'hong.gd', displayName: '홍길동' };
const T = (h: number, m = 0) => new Date(2026, 10, 9, h, m);
const INSPECTORS: Inspector[] = [
  { id: 'p1', name: '김민준', dept: '품질관리팀', email: 'mj@co.example', active: true, createdAt: '' },
  { id: 'p2', name: '이서연', dept: '품질관리팀', email: 'sy@co.example', active: true, createdAt: '' },
  { id: 'p3', name: '정하늘', dept: '품질관리팀', email: 'hn@co.example', active: false, createdAt: '' },
];

function make(over: Partial<Inspection> = {}): Inspection {
  const req = {
    vendorName: '한솔정밀', category: 'TANK', itemName: '베이스 플레이트', spec: 'BP-2205', quantity: 10,
    requestedDate: '2026-12-01', requestedTime: '10:30', location: '수입검사실 A',
    contactName: '김철수', contactPhone: '', contactEmail: 'kim@hansol.example.com', note: '',
  };
  return {
    ...req, id: 'IN-2026-0001', status: 'registered', original: { ...req },
    schedule: { date: '2026-12-01', time: '10:30', location: '수입검사실 A', inspectorId: null },
    confirmedAt: null, confirmedBy: null, sources: [], negotiations: [], result: null, attachments: [], history: [],
    createdAt: '', createdBy: '', ...over,
  };
}
const fields = (i: Inspection) => i.history.map((h) => [h.field, h.from, h.to]);

describe('상태 흐름', () => {
  it('등록 → 일정협의중 → 확정 (검수자 필수) → 변경 → 재확정', () => {
    let i = make();
    expect(nextAction(i)).toBe('negotiate');
    i = startNegotiation(i, HONG, T(9));
    expect(i.status).toBe('negotiating');
    expect(fields(i)).toEqual([['status', 'registered', 'negotiating']]);

    expect(() => confirm(i, { ...i.schedule }, INSPECTORS, HONG, T(10))).toThrow('확정하려면 검수자를 배정해야 합니다.');
    i = confirm(i, { ...i.schedule, time: '14:00', inspectorId: 'p1' }, INSPECTORS, HONG, T(10));
    expect(i.status).toBe('confirmed');
    expect(i.confirmedBy).toBe('홍길동');
    expect(fields(i).slice(1)).toEqual([
      ['schedule.time', '10:30', '14:00'], ['schedule.inspectorId', null, 'p1'], ['status', 'negotiating', 'confirmed'],
    ]);
    expect(i.history.every((h) => h.byName === '홍길동' && h.at)).toBe(true);

    expect(() => change(i, { ...i.schedule }, INSPECTORS, '', HONG, T(11))).toThrow('바뀐 내용이 없습니다');
    i = change(i, { ...i.schedule, date: '2026-12-03' }, INSPECTORS, '업체 입고 지연', HONG, T(11));
    expect(i.status).toBe('changed');
    expect(i.history.at(-1)).toMatchObject({ field: 'status', from: 'confirmed', to: 'changed', reason: '업체 입고 지연' });
    expect(nextAction(i)).toBe('confirm');

    i = confirm(i, { ...i.schedule }, INSPECTORS, HONG, T(12));
    expect(i.status).toBe('confirmed');
    expect(i.history.at(-1)).toMatchObject({ from: 'changed', to: 'confirmed', reason: '재확정' });
  });

  it('원본은 바꾸지 않는다', () => {
    const i = make();
    startNegotiation(i, HONG, T(9));
    expect(i.status).toBe('registered');
    expect(i.history).toHaveLength(0);
  });

  it('취소는 사유 필수, 검수완료·취소 건은 취소할 수 없다', () => {
    expect(() => cancel(make(), ' ', HONG, T(9))).toThrow('취소 사유를 입력하세요');
    const c = cancel(make({ status: 'confirmed' }), '업체 요청', HONG, T(9));
    expect(c.status).toBe('canceled');
    expect(c.history[0]).toMatchObject({ from: 'confirmed', to: 'canceled', reason: '업체 요청' });
    expect(() => cancel(c, '다시', HONG, T(9))).toThrow(WorkflowError);
    expect(can(make({ status: 'done' }), 'cancel')).toBe(false);
  });

  it('허용되지 않은 단계 이동은 막는다', () => {
    expect(() => change(make(), make().schedule, INSPECTORS, '', HONG, T(9))).toThrow("'등록' 상태에서는 변경할 수 없습니다");
    expect(() => startNegotiation(make({ status: 'confirmed' }), HONG, T(9))).toThrow('협의 시작할 수 없습니다');
    expect(can(make({ status: 'canceled' }), 'confirm')).toBe(false);
  });

  it('확정 값 점검: 날짜·시간·장소·검수자, 사용 안 함 검수자 거부', () => {
    expect(checkSchedule({ date: null, time: '', location: ' ', inspectorId: 'p3' }, INSPECTORS)).toEqual({
      date: '검수일을 입력하세요.',
      time: '검수 시간을 입력하세요. 예: 13:30',
      location: '검수 장소를 입력하세요.',
      inspectorId: "정하늘은(는) '사용 안 함' 상태입니다. 다른 검수자를 고르세요.",
    });
  });

  it('시간은 24시간 HH:MM 으로 맞추고, 범위를 벗어나면 거부', () => {
    expect(normalizeTime('9:30')).toBe('09:30');
    expect(normalizeTime('0930')).toBe('09:30');
    expect(normalizeTime('13시 30분')).toBe('13:30');
    expect(normalizeTime('25:00')).toBeNull();
    expect(normalizeTime('2:00 PM')).toBeNull();
    const base = { date: '2026-12-01', location: 'A', inspectorId: 'p1' };
    expect(checkSchedule({ ...base, time: '25:00' }, INSPECTORS).time).toBe('시간은 24시간 형식으로 입력하세요. 예: 09:30, 13:30');
    expect(checkSchedule({ ...base, time: '9:30' }, INSPECTORS).time).toBeDefined(); // 저장 전에 화면에서 09:30 으로 맞춤
    expect(checkSchedule({ ...base, time: '09:30' }, INSPECTORS)).toEqual({});
  });

  it('검수자 배정·기본정보 수정·협의 기록도 이력/기록이 남는다', () => {
    let i = assignInspector(make(), 'p2', INSPECTORS, HONG, T(9));
    expect(fields(i)).toEqual([['schedule.inspectorId', null, 'p2']]);
    i = editRequest(i, { quantity: 12, itemName: '베이스 플레이트' }, HONG, T(9), '전화로 확인');
    expect(fields(i).at(-1)).toEqual(['quantity', 10, 12]);
    expect(i.history).toHaveLength(2); // 같은 값은 기록 안 함
    i = addNote(i, '  업체 요청: 오후 희망 ', HONG, T(9));
    expect(i.negotiations).toEqual([{ at: T(9).toISOString(), by: 'hong.gd', byName: '홍길동', text: '업체 요청: 오후 희망' }]);
    expect(() => addNote(i, '', HONG, T(9))).toThrow('협의 내용을 입력하세요.');
  });

  it('상태 배지·결과 미입력 표시', () => {
    expect(statusBadge(make()).label).toBe('등록');
    expect(statusBadge(make({ status: 'done', result: { verdict: 'fail' } as any })).cls).toBe('st-fail');
    const c = make({ status: 'confirmed', schedule: { date: '2026-11-06', time: '10:00', location: 'A', inspectorId: 'p1' } });
    expect(resultMissing(c, '2026-11-09')).toBe(true);
    expect(resultMissing(c, '2026-11-06')).toBe(false);
  });
});

describe('구조 변환·가져오기 연동', () => {
  it('inspections v1 → v2: 현재 일정을 업체 희망값으로 채움', () => {
    const v1 = { schemaVersion: 1, revision: 2, items: [{ id: 'IN-2026-0001', requestedDate: '2026-12-01', requestedTime: '10:30', location: 'A', history: [] }] };
    const v2 = migrate('inspections', v1);
    expect(v2.schemaVersion).toBe(2);
    expect(v2.items[0].schedule).toEqual({ date: '2026-12-01', time: '10:30', location: 'A', inspectorId: null });
    expect(v2.items[0].confirmedAt).toBeNull();
  });

  it('확정된 건에 업체 수정본이 오면 신청 내용만 바뀌고 현재 일정·상태는 그대로', () => {
    const confirmed = make({ status: 'confirmed', schedule: { date: '2026-12-01', time: '14:00', location: '수입검사실 A', inspectorId: 'p1' } });
    const row: any = {
      key: '0:3', fileIndex: 0, fileName: 'r.xlsx', rowNumber: 3, kind: 'update', target: 'update',
      parsed: { ...confirmed, id: 'IN-2026-0001', requestedDate: '2026-12-05', location: '수입검사실 B' },
      display: {}, issues: {}, defaultInclude: true,
    };
    const out = applyImport({
      rows: [row], includeKeys: new Set(['0:3']), inspections: [confirmed], vendors: [],
      sourcePaths: ['원본/x.xlsx'], user: HONG, now: T(9),
    });
    const u = out.inspections[0];
    expect(u.requestedDate).toBe('2026-12-05');
    expect(u.location).toBe('수입검사실 B');
    expect(u.schedule).toEqual(confirmed.schedule);
    expect(u.status).toBe('confirmed');
  });
});
