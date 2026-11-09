import type { CellIssue, FieldKey, Inspection, Inspector, PreviewRow, Settings } from '../../shared/types';
import { conflictsFor, toSlot, type Slot } from '../../shared/conflicts';

const PRE_CONFIRM: Inspection['status'][] = ['received', 'registered', 'negotiating'];

/**
 * 가져오기 미리보기의 충돌 조건 검사.
 * 기존 일정 + 이번에 가져올 행(오류 행 제외)을 합쳐서, 각 행이 규칙을 넘는지 봅니다.
 * 경고만 하고 가져오기는 막지 않습니다.
 */
export function importConflicts(
  rows: PreviewRow[], existing: Inspection[], settings: Settings, inspectors: Inspector[],
): Map<string, Partial<Record<FieldKey, CellIssue[]>>> {
  const out = new Map<string, Partial<Record<FieldKey, CellIssue[]>>>();
  const rules = settings.conflictRules;
  if (!rules) return out;
  const names = new Map(inspectors.map((p) => [p.id, p.name]));
  const nameOf = (id: string) => names.get(id) ?? id;
  const byId = new Map(existing.map((i) => [i.id, i]));

  const rowSlot = (r: PreviewRow): Slot | null => {
    if (r.kind === 'error') return null;
    const p = r.parsed;
    if (r.target === 'update') {
      const cur = byId.get(p.id);
      // 확정된 건은 수정본이 와도 일정이 그대로이므로 새 충돌이 생기지 않음
      if (!cur || !PRE_CONFIRM.includes(cur.status)) return null;
      return { ...toSlot(cur), date: p.requestedDate ?? cur.schedule.date, time: p.requestedTime ?? cur.schedule.time, location: p.location || cur.schedule.location, quantity: p.quantity ?? cur.quantity };
    }
    return {
      id: `${r.fileName} ${r.rowNumber}행`, date: p.requestedDate, time: p.requestedTime, location: p.location,
      inspectorId: null, vendorName: p.vendorName, quantity: p.quantity, status: 'registered',
    };
  };

  const pairs = rows.map((r) => [r, rowSlot(r)] as const).filter((x): x is readonly [PreviewRow, Slot] => !!x[1]);
  const replaced = new Set(pairs.filter(([r]) => r.target === 'update').map(([r]) => r.parsed.id));
  const all = [...existing.filter((i) => !replaced.has(i.id)).map(toSlot), ...pairs.map(([, s]) => s)];

  for (const [r, s] of pairs) {
    const cs = conflictsFor(s, all, rules, nameOf);
    if (cs.length) {
      out.set(r.key, {
        requestedDate: cs.map((c) => ({ level: 'warn' as const, message: `충돌: ${c.message} — 가져온 뒤 일정 조정 필요` })),
      });
    }
  }
  return out;
}
