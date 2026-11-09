import type {
  CellIssue, FieldKey, HistoryEntry, Inspection, PreviewRow, RequestFields, Settings, SourceRef, Vendor,
} from '../../shared/types';
import { REQUEST_FIELDS } from '../../shared/types';
import { cellText, isEmail, parseDate, parseQuantity, parseTime } from './excel/cells';
import type { RawRow } from './excel/read';
import { templateDef } from './schema';

// ───────── 신청번호 ─────────

export const ID_RE = /^IN-(\d{4})-(\d{4,})$/;

/** 신청번호 표기를 통일 (소문자·공백 허용) */
export function normalizeId(s: string): string {
  return s.trim().toUpperCase().replace(/\s+/g, '');
}

/** 그 해의 다음 신청번호: IN-2026-0001 부터 */
export function nextId(existing: string[], year: number): string {
  let max = 0;
  for (const id of existing) {
    const m = ID_RE.exec(id);
    if (m && +m[1] === year) max = Math.max(max, +m[2]);
  }
  return `IN-${year}-${String(max + 1).padStart(4, '0')}`;
}

// ───────── 판정 ─────────

export interface ValidateContext {
  settings: Settings;
  existing: Inspection[];
  /** 오늘 (YYYY-MM-DD) */
  today: string;
  /**
   * 충돌 조건 검사 (4단계에서 연결). 행 목록을 받아 행 key → 칸별 경고를 돌려줍니다.
   */
  checkConflicts?: (rows: PreviewRow[]) => Map<string, Partial<Record<FieldKey, CellIssue[]>>>;
}

const EMPTY_FIELDS = (): RequestFields => ({
  vendorName: '', category: '', itemName: '', spec: '', quantity: null, requestedDate: null, requestedTime: null,
  location: '', contactName: '', contactPhone: '', contactEmail: '', note: '',
});

const show = (v: unknown) => (v === null || v === undefined || v === '' ? '(비어 있음)' : String(v));

/** 한 파일에서 읽은 행들을 판정합니다. 오류가 있어도 막지 않고 칸별로 표시만 합니다. */
export function buildPreviewRows(
  raw: RawRow[], fileIndex: number, fileName: string, templateVersion: string | undefined, ctx: ValidateContext,
): PreviewRow[] {
  const def = templateDef(ctx.settings, templateVersion);
  const header = (k: FieldKey) => def.columns.find((c) => c.key === k)?.header ?? k;
  const required = new Set(def.columns.filter((c) => c.required).map((c) => c.key));
  const byId = new Map(ctx.existing.map((i) => [i.id, i]));
  const items = new Set(ctx.settings.inspectionItems ?? []);

  return raw.map((r) => {
    const issues: PreviewRow['issues'] = {};
    const add = (k: FieldKey | 'row', level: CellIssue['level'], message: string) => {
      (issues[k] ??= []).push({ level, message });
    };
    const display = {} as Record<FieldKey, string>;
    for (const c of def.columns) display[c.key] = cellText(r.values[c.key]);

    const p = { ...EMPTY_FIELDS(), id: '' };
    const text = (k: keyof RequestFields | 'id') => cellText(r.values[k as FieldKey]);
    p.id = text('id') ? normalizeId(text('id')) : '';
    p.vendorName = text('vendorName');
    p.category = text('category');
    p.itemName = text('itemName');
    p.spec = text('spec');
    p.location = text('location');
    p.contactName = text('contactName');
    p.contactPhone = text('contactPhone');
    p.contactEmail = text('contactEmail');
    p.note = text('note');

    // 필수값
    for (const k of required) {
      if (display[k] === '') add(k, 'error', `필수 항목 누락: ${header(k)}을(를) 입력해야 합니다.`);
    }

    // 형식
    const q = parseQuantity(r.values.quantity);
    if (q === null) add('quantity', 'error', '수량은 1 이상의 정수로 입력해야 합니다. 예: 10');
    else p.quantity = q === '' ? null : q;

    const d = parseDate(r.values.requestedDate);
    if (d === null) add('requestedDate', 'error', '날짜 형식이 올바르지 않습니다. 예: 2026-11-14');
    else {
      p.requestedDate = d || null;
      if (d && d < ctx.today) add('requestedDate', 'warn', `검수 희망일이 오늘(${ctx.today})보다 이전입니다.`);
    }

    const t = parseTime(r.values.requestedTime);
    if (t === null) add('requestedTime', 'error', '시간 형식이 올바르지 않습니다. 24시간 형식으로 적어야 합니다. 예: 13:30');
    else p.requestedTime = t || null;

    if (p.contactEmail && !isEmail(p.contactEmail)) {
      add('contactEmail', 'error', '이메일 형식이 올바르지 않습니다. 예: name@company.co.kr');
    }

    if (p.category && items.size && !items.has(p.category)) {
      add('category', 'warn', `설정의 검수 품목 목록에 없는 품목입니다. (목록: ${[...items].join(', ')})`);
    }

    // 신청번호 매칭
    let target: PreviewRow['target'] = 'new';
    let diff: PreviewRow['diff'];
    if (p.id) {
      const cur = byId.get(p.id);
      if (!cur) {
        add('id', 'error', `없는 신청번호입니다. 신청번호가 맞는지 업체에 확인하세요. 포함하면 새 신청번호로 등록합니다.`);
      } else {
        target = 'update';
        diff = {};
        for (const k of REQUEST_FIELDS) {
          const next = p[k];
          // 비운 칸은 기존 값을 지우지 않음
          if (next === '' || next === null) continue;
          if (String(cur[k] ?? '') !== String(next)) diff[k] = { from: show(cur[k]), to: show(next) };
        }
        if (cur.status === 'canceled') add('id', 'warn', '취소된 건입니다. 포함하면 내용만 바뀌고 상태는 취소로 남습니다.');
        if (cur.status === 'done') add('id', 'warn', '이미 검수가 끝난 건입니다.');
        if (cur.vendorName && p.vendorName && cur.vendorName !== p.vendorName) {
          add('vendorName', 'warn', `이 신청번호의 업체는 '${cur.vendorName}'입니다. 다른 업체의 번호를 적었는지 확인하세요.`);
        }
      }
    }

    const hasError = Object.values(issues).some((l) => l!.some((i) => i.level === 'error'));
    const unchanged = target === 'update' && diff && Object.keys(diff).length === 0;
    if (unchanged) add('row', 'warn', '바뀐 내용이 없습니다.');

    return {
      key: `${fileIndex}:${r.rowNumber}`,
      fileIndex,
      fileName,
      rowNumber: r.rowNumber,
      kind: hasError ? 'error' : target,
      target,
      display,
      parsed: p,
      issues,
      diff,
      defaultInclude: !hasError && !unchanged,
    } satisfies PreviewRow;
  });
}

/** 여러 파일의 행을 합친 뒤 파일끼리 걸리는 문제(같은 번호 중복)와 충돌 조건을 붙입니다. */
export function finalizePreview(rows: PreviewRow[], ctx: ValidateContext): PreviewRow[] {
  const seen = new Map<string, PreviewRow>();
  for (const r of rows) {
    if (!r.parsed.id || r.target !== 'update') continue;
    const prev = seen.get(r.parsed.id);
    if (prev) {
      (prev.issues.id ??= []).push({ level: 'warn', message: `같은 신청번호가 아래(${r.fileName} ${r.rowNumber}행)에도 있습니다. 둘 다 포함하면 나중 행이 반영됩니다.` });
      (r.issues.id ??= []).push({ level: 'warn', message: `같은 신청번호가 위(${prev.fileName} ${prev.rowNumber}행)에도 있습니다.` });
    }
    seen.set(r.parsed.id, r);
  }
  const conflicts = ctx.checkConflicts?.(rows);
  if (conflicts) {
    for (const r of rows) {
      const c = conflicts.get(r.key);
      if (!c) continue;
      for (const [k, list] of Object.entries(c)) (r.issues[k as FieldKey] ??= []).push(...list!);
    }
  }
  return rows;
}

// ───────── 적용 ─────────

export interface ApplyInput {
  rows: PreviewRow[];
  includeKeys: Set<string>;
  inspections: Inspection[];
  vendors: Vendor[];
  /** fileIndex → 원본 보관 상대 경로 */
  sourcePaths: string[];
  user: { account: string; displayName: string };
  now: Date;
}

export interface ApplyOutput {
  inspections: Inspection[];
  vendors: Vendor[];
  created: Inspection[];
  updated: Inspection[];
  excluded: number;
}

/** 선택한 행을 검수 건·업체 목록에 반영합니다 (저장은 하지 않음). */
export function applyImport(inp: ApplyInput): ApplyOutput {
  const iso = inp.now.toISOString();
  const by = { at: iso, by: inp.user.account, byName: inp.user.displayName };
  const inspections = inp.inspections.map((i) => ({ ...i, history: [...i.history], sources: [...i.sources] }));
  const index = new Map(inspections.map((i) => [i.id, i]));
  const vendors = inp.vendors.map((v) => ({ ...v }));
  const created: Inspection[] = [];
  const updated = new Map<string, Inspection>();
  let excluded = 0;

  for (const row of inp.rows) {
    if (!inp.includeKeys.has(row.key)) {
      excluded++;
      continue;
    }
    const p = row.parsed;
    const source: SourceRef = { file: inp.sourcePaths[row.fileIndex], row: row.rowNumber, importedAt: iso, importedBy: inp.user.displayName };
    const fields: RequestFields = {
      vendorName: p.vendorName, category: p.category, itemName: p.itemName, spec: p.spec, quantity: p.quantity,
      requestedDate: p.requestedDate, requestedTime: p.requestedTime, location: p.location,
      contactName: p.contactName, contactPhone: p.contactPhone, contactEmail: p.contactEmail, note: p.note,
    };
    const cur = row.target === 'update' ? index.get(p.id) : undefined;

    if (cur) {
      const hist: HistoryEntry[] = [];
      for (const k of REQUEST_FIELDS) {
        const next = fields[k];
        if (next === '' || next === null) continue;
        if (String(cur[k] ?? '') === String(next)) continue;
        hist.push({ ...by, action: 'field', field: k, from: cur[k], to: next, reason: `업체 수정본 가져오기 (${row.fileName} ${row.rowNumber}행)` });
        (cur as any)[k] = next;
      }
      cur.sources.push(source);
      cur.history.push(...hist);
      updated.set(cur.id, cur);
    } else {
      const id = nextId(inspections.map((i) => i.id), inp.now.getFullYear());
      const memo = p.id ? ` · 업체가 적은 신청번호 ${p.id}는 없는 번호여서 새로 부여` : '';
      const insp: Inspection = {
        ...fields,
        id,
        status: 'registered',
        original: { ...fields },
        sources: [source],
        negotiations: [],
        result: null,
        attachments: [],
        history: [
          { ...by, action: 'created', reason: `엑셀 가져오기 (${row.fileName} ${row.rowNumber}행)${memo}` },
          { ...by, action: 'status', from: 'received', to: 'registered', reason: `신청번호 ${id} 부여` },
        ],
        createdAt: iso,
        createdBy: inp.user.displayName,
      };
      inspections.push(insp);
      index.set(id, insp);
      created.push(insp);
    }

    // 업체 목록 자동 반영 (적힌 값만 덮어씀)
    if (fields.vendorName) {
      let v = vendors.find((x) => x.name === fields.vendorName);
      if (!v) {
        v = { name: fields.vendorName, contactName: '', contactPhone: '', contactEmail: '', updatedAt: iso };
        vendors.push(v);
      }
      if (fields.contactName) v.contactName = fields.contactName;
      if (fields.contactPhone) v.contactPhone = fields.contactPhone;
      if (fields.contactEmail && isEmail(fields.contactEmail)) v.contactEmail = fields.contactEmail;
      v.updatedAt = iso;
    }
  }
  return { inspections, vendors, created, updated: [...updated.values()], excluded };
}

/** 원본 보관 파일 이름: 업체명_가져온일시_원파일명.xlsx */
export function originalFileName(vendorNames: string[], now: Date, fileName: string): string {
  const pad = (n: number) => String(n).padStart(2, '0');
  const uniq = [...new Set(vendorNames.filter(Boolean))];
  const vendor = uniq.length === 0 ? '업체미상' : uniq.length <= 2 ? uniq.join('_') : `${uniq[0]}외${uniq.length - 1}곳`;
  const stamp = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
  const safe = (s: string) => s.replace(/[\\/:*?"<>|]+/g, '_').trim();
  const base = safe(fileName).replace(/\.xlsx$/i, '');
  return `${safe(vendor)}_${stamp}_${base}.xlsx`;
}
