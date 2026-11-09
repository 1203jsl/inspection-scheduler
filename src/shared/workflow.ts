/**
 * 검수 건 상태 흐름과 변경 이력.
 *
 *   신청접수 → 등록(신청번호 부여) → 일정협의중 → 확정 → 검수완료(6단계)
 *                                       ↘ 변경(재확정 필요) ↗
 *                   어느 단계에서든(검수완료 제외) → 취소
 *
 * 모든 함수는 원본을 바꾸지 않고 새 객체를 돌려줍니다.
 * 규칙에 맞지 않으면 WorkflowError(화면에 그대로 보여줄 한국어 문구)를 던집니다.
 */
import type {
  HistoryEntry, Inspection, InspectionStatus, Inspector, RequestFields, Schedule,
} from './types';
import { REQUEST_FIELDS, STATUS_LABEL } from './types';
import { normalizeTime } from './dates';

export class WorkflowError extends Error {}

export interface Actor {
  account: string;
  displayName: string;
}

export type Action = 'negotiate' | 'confirm' | 'change' | 'cancel' | 'assign' | 'edit' | 'note';

const ALLOWED: Record<Action, InspectionStatus[]> = {
  negotiate: ['received', 'registered'],
  confirm: ['received', 'registered', 'negotiating', 'changed'],
  change: ['confirmed', 'changed'],
  cancel: ['received', 'registered', 'negotiating', 'confirmed', 'changed'],
  assign: ['received', 'registered', 'negotiating', 'confirmed', 'changed'],
  edit: ['received', 'registered', 'negotiating', 'confirmed', 'changed', 'done', 'canceled'],
  note: ['received', 'registered', 'negotiating', 'confirmed', 'changed', 'done', 'canceled'],
};

const ACTION_LABEL: Record<Action, string> = {
  negotiate: '협의 시작', confirm: '확정', change: '변경', cancel: '취소',
  assign: '검수자 배정', edit: '기본정보 수정', note: '협의 기록',
};

export function can(insp: Pick<Inspection, 'status'>, action: Action): boolean {
  return ALLOWED[action].includes(insp.status);
}

function ensure(insp: Inspection, action: Action) {
  if (!can(insp, action)) {
    throw new WorkflowError(`${insp.id}: '${STATUS_LABEL[insp.status]}' 상태에서는 ${ACTION_LABEL[action]}할 수 없습니다.`);
  }
}

/** 변경 이력에 쓰는 항목 이름 */
export const FIELD_LABEL: Record<string, string> = {
  vendorName: '업체명', category: '검수 품목', itemName: '품명', spec: '품번/규격', quantity: '수량',
  requestedDate: '검수 희망일', requestedTime: '희망 시간', location: '검수 장소(신청)',
  contactName: '업체 담당자', contactPhone: '연락처', contactEmail: '이메일', note: '비고',
  'schedule.date': '검수일', 'schedule.time': '검수 시간', 'schedule.location': '검수 장소',
  'schedule.inspectorId': '검수자', status: '상태', conflict: '충돌 강제 진행',
};

function clone(i: Inspection): Inspection {
  return { ...i, schedule: { ...i.schedule }, history: [...i.history], negotiations: [...i.negotiations] };
}

function entry(actor: Actor, now: Date, e: Omit<HistoryEntry, 'at' | 'by' | 'byName'>): HistoryEntry {
  return { at: now.toISOString(), by: actor.account, byName: actor.displayName, ...e };
}

function setStatus(i: Inspection, to: InspectionStatus, actor: Actor, now: Date, reason?: string) {
  if (i.status === to) return;
  i.history.push(entry(actor, now, { action: 'status', field: 'status', from: i.status, to, reason }));
  i.status = to;
}

function setSchedule(i: Inspection, next: Schedule, actor: Actor, now: Date, reason?: string) {
  for (const k of ['date', 'time', 'location', 'inspectorId'] as const) {
    const a = i.schedule[k] ?? null;
    const b = next[k] ?? null;
    if ((a ?? '') === (b ?? '')) continue;
    i.history.push(entry(actor, now, { action: 'field', field: `schedule.${k}`, from: a, to: b, reason }));
    (i.schedule as any)[k] = b;
  }
}

// ───────── 입력 점검 ─────────

export type ScheduleErrors = Partial<Record<keyof Schedule, string>>;

/** 확정·변경 때 필요한 값 점검. 검수자 배정은 필수입니다. */
export function checkSchedule(s: Schedule, inspectors: Inspector[]): ScheduleErrors {
  const e: ScheduleErrors = {};
  if (!s.date || !/^\d{4}-\d{2}-\d{2}$/.test(s.date)) e.date = '검수일을 입력하세요.';
  if (!s.time) e.time = '검수 시간을 입력하세요. 예: 13:30';
  else if (normalizeTime(s.time) !== s.time) e.time = '시간은 24시간 형식으로 입력하세요. 예: 09:30, 13:30';
  if (!s.location?.trim()) e.location = '검수 장소를 입력하세요.';
  if (!s.inspectorId) e.inspectorId = '확정하려면 검수자를 배정해야 합니다.';
  else {
    const p = inspectors.find((x) => x.id === s.inspectorId);
    if (!p) e.inspectorId = '검수자 목록에 없는 사람입니다. 다시 고르세요.';
    else if (!p.active) e.inspectorId = `${p.name}은(는) '사용 안 함' 상태입니다. 다른 검수자를 고르세요.`;
  }
  return e;
}

function assertSchedule(i: Inspection, s: Schedule, inspectors: Inspector[]) {
  const e = checkSchedule(s, inspectors);
  const msgs = Object.values(e);
  if (msgs.length) throw new WorkflowError(`${i.id}: ${msgs.join(' ')}`);
}

// ───────── 충돌 강제 진행 ─────────

/** 충돌이 있는데 진행할 때: 걸린 충돌과 사유 */
export interface ForceInfo {
  conflicts: string[];
  reason: string;
}

/**
 * 충돌이 있으면 강제 진행 허용 여부와 사유를 확인합니다.
 * 통과하면 이력에 남길 ForceInfo(충돌 없으면 undefined)를 돌려줍니다.
 */
export function checkForce(id: string, conflicts: string[], allowForce: boolean, reason: string | undefined): ForceInfo | undefined {
  if (conflicts.length === 0) return undefined;
  if (!allowForce) {
    throw new WorkflowError(`${id}: 충돌 조건에 걸립니다 (${conflicts.join(' / ')}). 설정에서 강제 진행이 꺼져 있어 진행할 수 없습니다. 날짜나 검수자를 바꾸세요.`);
  }
  if (!reason?.trim()) {
    throw new WorkflowError(`${id}: 충돌 조건에 걸립니다 (${conflicts.join(' / ')}). 강제로 진행하려면 사유를 입력하세요.`);
  }
  return { conflicts, reason: reason.trim() };
}

function recordForce(i: Inspection, f: ForceInfo | undefined, actor: Actor, now: Date) {
  if (!f) return;
  i.history.push(entry(actor, now, { action: 'force', field: 'conflict', to: f.conflicts.join(' / '), reason: f.reason }));
}

// ───────── 동작 ─────────

export function startNegotiation(insp: Inspection, actor: Actor, now: Date): Inspection {
  ensure(insp, 'negotiate');
  const i = clone(insp);
  setStatus(i, 'negotiating', actor, now);
  return i;
}

/** 확정 (재확정 포함). reason: 충돌 강제 진행 사유 등 (4단계) */
export function confirm(insp: Inspection, s: Schedule, inspectors: Inspector[], actor: Actor, now: Date, reason?: string, force?: ForceInfo): Inspection {
  ensure(insp, 'confirm');
  assertSchedule(insp, s, inspectors);
  const i = clone(insp);
  recordForce(i, force, actor, now);
  setSchedule(i, s, actor, now, reason);
  setStatus(i, 'confirmed', actor, now, reason ?? (insp.status === 'changed' ? '재확정' : undefined));
  i.confirmedAt = now.toISOString();
  i.confirmedBy = actor.displayName;
  return i;
}

/** 확정 뒤 일정 변경 → '변경됨'(재확정 필요) */
export function change(insp: Inspection, s: Schedule, inspectors: Inspector[], reason: string, actor: Actor, now: Date, force?: ForceInfo): Inspection {
  ensure(insp, 'change');
  assertSchedule(insp, s, inspectors);
  const same = (['date', 'time', 'location', 'inspectorId'] as const).every((k) => (insp.schedule[k] ?? '') === (s[k] ?? ''));
  if (same) throw new WorkflowError(`${insp.id}: 바뀐 내용이 없습니다. 검수일·시간·장소·검수자 중 하나 이상을 바꾸세요.`);
  const i = clone(insp);
  const r = reason.trim() || undefined;
  recordForce(i, force, actor, now);
  setSchedule(i, s, actor, now, r);
  setStatus(i, 'changed', actor, now, r);
  return i;
}

export function cancel(insp: Inspection, reason: string, actor: Actor, now: Date): Inspection {
  ensure(insp, 'cancel');
  if (!reason.trim()) throw new WorkflowError('취소 사유를 입력하세요. 변경 이력에 남습니다.');
  const i = clone(insp);
  setStatus(i, 'canceled', actor, now, reason.trim());
  return i;
}

export function assignInspector(insp: Inspection, inspectorId: string, inspectors: Inspector[], actor: Actor, now: Date): Inspection {
  ensure(insp, 'assign');
  const p = inspectors.find((x) => x.id === inspectorId);
  if (!p || !p.active) throw new WorkflowError('사용 중인 검수자를 고르세요.');
  const i = clone(insp);
  setSchedule(i, { ...i.schedule, inspectorId }, actor, now);
  return i;
}

/** 업체 신청 내용(기본정보) 수정 */
export function editRequest(insp: Inspection, patch: Partial<RequestFields>, actor: Actor, now: Date, reason?: string): Inspection {
  ensure(insp, 'edit');
  const i = clone(insp);
  for (const k of REQUEST_FIELDS) {
    if (!(k in patch)) continue;
    const next = patch[k] ?? null;
    const cur = i[k] ?? null;
    if (String(cur ?? '') === String(next ?? '')) continue;
    i.history.push(entry(actor, now, { action: 'field', field: k, from: cur, to: next, reason }));
    (i as any)[k] = next;
  }
  return i;
}

export function addNote(insp: Inspection, text: string, actor: Actor, now: Date): Inspection {
  ensure(insp, 'note');
  if (!text.trim()) throw new WorkflowError('협의 내용을 입력하세요.');
  const i = clone(insp);
  i.negotiations.push({ at: now.toISOString(), by: actor.account, byName: actor.displayName, text: text.trim() });
  return i;
}

// ───────── 표시 ─────────

export type BadgeClass = 'st-received' | 'st-negotiating' | 'st-confirmed' | 'st-changed' | 'st-canceled' | 'st-pass' | 'st-fail';

/** 상태 배지 (디자인의 StatusBadge 규칙) */
export function statusBadge(i: Pick<Inspection, 'status' | 'result'>): { cls: BadgeClass; label: string } {
  switch (i.status) {
    case 'received': return { cls: 'st-received', label: '신청접수' };
    case 'registered': return { cls: 'st-received', label: '등록' };
    case 'negotiating': return { cls: 'st-negotiating', label: '일정협의중' };
    case 'confirmed': return { cls: 'st-confirmed', label: '확정' };
    case 'changed': return { cls: 'st-changed', label: '변경됨' };
    case 'canceled': return { cls: 'st-canceled', label: '취소' };
    case 'done':
      return i.result?.verdict === 'fail'
        ? { cls: 'st-fail', label: '검수완료 · 불합격' }
        : { cls: 'st-pass', label: '검수완료 · 합격' };
  }
}

/** 확정됐는데 검수일이 지나도록 결과가 없음 */
export function resultMissing(i: Inspection, today: string): boolean {
  return (i.status === 'confirmed' || i.status === 'changed') && !!i.schedule.date && i.schedule.date < today && !i.result;
}

/** 다음에 할 일 (상세 화면의 강조 버튼) */
export function nextAction(i: Inspection): Action | null {
  if (i.status === 'received' || i.status === 'registered') return 'negotiate';
  if (i.status === 'negotiating' || i.status === 'changed') return 'confirm';
  return null;
}
