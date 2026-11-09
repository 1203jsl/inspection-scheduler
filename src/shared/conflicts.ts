/**
 * 충돌 조건 검사.
 * 날짜별로 그날 잡힌 검수(취소 제외)를 모아 설정의 규칙을 넘는지 봅니다.
 * 확정·변경·가져오기 때 쓰고, 달력·목록의 충돌 표시에도 씁니다.
 */
import type { ConflictRules, Inspection, InspectionStatus } from './types';
import { fmtMD } from './dates';

export const DEFAULT_CONFLICT_RULES: ConflictRules = {
  maxPerDay: { enabled: true, value: 3 },
  maxPerInspector: { enabled: true, value: 2 },
  sameSlot: { enabled: true, minutes: 60 },
  maxQtyPerDay: { enabled: false, value: 1000 },
  maxPerVendor: { enabled: true, value: 2 },
  holidays: { enabled: true, dates: [] },
  allowForce: true,
};

export type RuleKey = 'maxPerDay' | 'maxPerInspector' | 'sameSlot' | 'maxQtyPerDay' | 'maxPerVendor' | 'holiday';

export interface Conflict {
  rule: RuleKey;
  date: string;
  message: string;
  /** 이 충돌에 걸린 건 */
  ids: string[];
}

/** 충돌 검사에 필요한 값만 */
export interface Slot {
  id: string;
  date: string | null;
  time: string | null;
  location: string;
  inspectorId: string | null;
  vendorName: string;
  quantity: number | null;
  status?: InspectionStatus;
}

export const toSlot = (i: Inspection): Slot => ({
  id: i.id, date: i.schedule.date, time: i.schedule.time, location: i.schedule.location,
  inspectorId: i.schedule.inspectorId, vendorName: i.vendorName, quantity: i.quantity, status: i.status,
});

/** 취소된 건은 일정을 차지하지 않음 */
const counts = (s: Slot) => !!s.date && s.status !== 'canceled';

const minutes = (t: string | null) => {
  const m = t?.match(/^(\d{2}):(\d{2})$/);
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
};

function group<T>(list: T[], key: (x: T) => string | null): Map<string, T[]> {
  const m = new Map<string, T[]>();
  for (const x of list) {
    const k = key(x);
    if (!k) continue;
    (m.get(k) ?? m.set(k, []).get(k)!).push(x);
  }
  return m;
}

/** 하루치 충돌 */
export function dayConflicts(date: string, slots: Slot[], r: ConflictRules, nameOf: (id: string) => string): Conflict[] {
  const day = slots.filter((s) => counts(s) && s.date === date);
  const out: Conflict[] = [];
  const md = fmtMD(date);
  if (r.holidays.enabled) {
    const h = r.holidays.dates.find((x) => x.date === date);
    if (h && day.length) out.push({ rule: 'holiday', date, ids: day.map((s) => s.id), message: `검수 불가일입니다 — ${md}${h.name ? ` ${h.name}` : ''}` });
  }
  if (r.maxPerDay.enabled && day.length > r.maxPerDay.value) {
    out.push({ rule: 'maxPerDay', date, ids: day.map((s) => s.id), message: `하루 최대 검수 건수(${r.maxPerDay.value}건) 초과 — ${md} ${day.length}건` });
  }
  if (r.maxPerInspector.enabled) {
    for (const [pid, list] of group(day, (s) => s.inspectorId)) {
      if (list.length > r.maxPerInspector.value) {
        out.push({ rule: 'maxPerInspector', date, ids: list.map((s) => s.id), message: `검수자 1인당 하루 최대 ${r.maxPerInspector.value}건 초과 — ${nameOf(pid)} ${md} ${list.length}건` });
      }
    }
  }
  if (r.maxPerVendor.enabled) {
    for (const [v, list] of group(day, (s) => s.vendorName || null)) {
      if (list.length > r.maxPerVendor.value) {
        out.push({ rule: 'maxPerVendor', date, ids: list.map((s) => s.id), message: `같은 업체 하루 최대 ${r.maxPerVendor.value}건 초과 — ${v} ${md} ${list.length}건` });
      }
    }
  }
  if (r.maxQtyPerDay.enabled) {
    const sum = day.reduce((a, s) => a + (s.quantity ?? 0), 0);
    if (sum > r.maxQtyPerDay.value) {
      out.push({ rule: 'maxQtyPerDay', date, ids: day.map((s) => s.id), message: `하루 최대 검수 수량(${r.maxQtyPerDay.value.toLocaleString()}개) 초과 — ${md} 합계 ${sum.toLocaleString()}개` });
    }
  }
  if (r.sameSlot.enabled) {
    for (const [loc, list] of group(day, (s) => s.location.trim() || null)) {
      const timed = list.map((s) => ({ s, m: minutes(s.time) })).filter((x) => x.m !== null).sort((a, b) => a.m! - b.m!);
      for (let a = 0; a < timed.length; a++) {
        for (let b = a + 1; b < timed.length && timed[b].m! - timed[a].m! < r.sameSlot.minutes; b++) {
          const x = timed[a].s, y = timed[b].s;
          out.push({ rule: 'sameSlot', date, ids: [x.id, y.id], message: `같은 시간대·같은 장소 중복 — ${md} ${loc} ${x.time}(${x.id}) / ${y.time}(${y.id})` });
        }
      }
    }
  }
  return out;
}

/** 날짜별 충돌 (달력·목록 표시용) */
export function conflictMap(slots: Slot[], r: ConflictRules, nameOf: (id: string) => string): Map<string, Conflict[]> {
  const m = new Map<string, Conflict[]>();
  const dates = new Set(slots.filter(counts).map((s) => s.date!));
  for (const d of dates) {
    const c = dayConflicts(d, slots, r, nameOf);
    if (c.length) m.set(d, c);
  }
  return m;
}

/** 이 건을 이 일정으로 두면 생기는 충돌 (자기 자신의 기존 일정은 빼고 계산) */
export function conflictsFor(target: Slot, all: Slot[], r: ConflictRules, nameOf: (id: string) => string): Conflict[] {
  if (!target.date || target.status === 'canceled') return [];
  const slots = [...all.filter((s) => s.id !== target.id), target];
  return dayConflicts(target.date, slots, r, nameOf).filter((c) => c.ids.includes(target.id));
}
