import { useMemo, useState } from 'react';
import type { AppErrorInfo, Inspection, InspectionStatus } from '../shared/types';
import { addDays, addMonths, fmtMD, monthEnd, monthStart, weekStart } from '../shared/dates';
import { assignInspector, cancel, confirm, itemLabel, resultMissing, startNegotiation, statusBadge, WorkflowError } from '../shared/workflow';
import { conflictMap, conflictsFor, toSlot } from '../shared/conflicts';
import { Alert, ErrorAlert, Icon } from './components';
import { api } from './api';
import { useApp, type ListPreset } from './ctx';
import { CancelDialog } from './dialogs';

type SortKey = 'id' | 'vendorName' | 'projectName' | 'category' | 'itemName' | 'quantity' | 'date' | 'time' | 'location' | 'inspector' | 'status' | 'result';
type QuickPeriod = 'today' | 'week' | 'month' | 'nextMonth';
const QUICK: { v: QuickPeriod; label: string }[] = [
  { v: 'today', label: '오늘' }, { v: 'week', label: '이번 주' }, { v: 'month', label: '이번 달' }, { v: 'nextMonth', label: '다음 달' },
];
/** 빠른 선택 → 시작일·종료일 */
function quickRange(p: QuickPeriod, today: string): [string, string] {
  switch (p) {
    case 'today': return [today, today];
    case 'week': { const s = weekStart(today); return [s, addDays(s, 6)]; }
    case 'month': return [monthStart(today), monthEnd(today)];
    case 'nextMonth': { const s = addMonths(today, 1); return [s, monthEnd(s)]; }
  }
}

const STATUS_OPTIONS: { v: InspectionStatus; label: string }[] = [
  { v: 'registered', label: '등록' },
  { v: 'negotiating', label: '일정협의중' },
  { v: 'confirmed', label: '확정' },
  { v: 'changed', label: '변경됨' },
  { v: 'done', label: '검수완료' },
  { v: 'canceled', label: '취소' },
];
const PAGE = 50;

export function StatusBadge({ i }: { i: Pick<Inspection, 'status' | 'result'> }) {
  const b = statusBadge(i);
  return <span className={`badge ${b.cls}`}>{b.label}</span>;
}

export function ListPage({ preset }: { preset?: ListPreset }) {
  const { data, today, go, inspectorName, saveInspections, user } = useApp();
  const items = data.inspections.items;
  const inspectors = data.inspectors.items;

  const [q, setQ] = useState(preset?.q ?? '');
  const init: [string, string] = preset?.day ? [preset.day, preset.day] : preset?.period ? quickRange(preset.period, today) : ['', ''];
  const [from, setFrom] = useState(init[0]);
  const [to, setTo] = useState(init[1]);
  const [vendor, setVendor] = useState(preset?.vendor ?? '');
  const [project, setProject] = useState(preset?.project ?? '');
  const [category, setCategory] = useState(preset?.category ?? '');
  const [insp, setInsp] = useState(preset?.unassigned ? '__none' : preset?.inspectorId ?? '');
  const [status, setStatus] = useState<string>(preset?.status ?? '');
  const [result, setResult] = useState(preset?.result ?? '');
  const [exportMsg, setExportMsg] = useState<{ ok: boolean; text: string } | null>(null);
  const [sort, setSort] = useState<{ key: SortKey; dir: 1 | -1 }>({ key: 'date', dir: 1 });
  const [page, setPage] = useState(0);
  const [sel, setSel] = useState<Set<string>>(new Set());
  const [bulkInsp, setBulkInsp] = useState('');
  const [bulkStatus, setBulkStatus] = useState('');
  const [bulkMsg, setBulkMsg] = useState<{ ok: string; fails: string[] } | null>(null);
  const [bulkErr, setBulkErr] = useState<AppErrorInfo | null>(null);
  const [cancelOpen, setCancelOpen] = useState(false);

  // 기간: 시작일·종료일 중 하나만 있어도 됨. 거꾸로 넣으면 바꿔서 적용
  const range = useMemo((): [string, string] | null => {
    if (!from && !to) return null;
    const a = from || '0000-00-00', b = to || '9999-12-31';
    return a <= b ? [a, b] : [b, a];
  }, [from, to]);
  const period = !from && !to ? '' : QUICK.find((x) => { const [a, b] = quickRange(x.v, today); return a === from && b === to; })?.v ?? 'custom';
  const periodText = range ? `${from ? range[0].replace(/-/g, '.') : ''} ~ ${to ? range[1].replace(/-/g, '.') : ''}` : '';

  const vendors = useMemo(() => [...new Set(items.map((i) => i.vendorName).filter(Boolean))].sort(), [items]);
  const projects = useMemo(() => [...new Set(items.map((i) => i.projectName).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'ko')), [items]);
  // 검수 품목: 설정 목록 순서 + 실제 건에만 있는 품목
  const categories = useMemo(() => [...new Set([...(data.settings.inspectionItems ?? []), ...items.map((i) => i.category).filter(Boolean)])], [items, data.settings.inspectionItems]);
  const rules = data.settings.conflictRules;
  const cmap = useMemo(() => conflictMap(items.map(toSlot), rules, inspectorName), [items, rules, inspectorName]);
  const conflictOf = (i: Inspection) => (i.schedule.date ? (cmap.get(i.schedule.date) ?? []).filter((c) => c.ids.includes(i.id)).map((c) => c.message) : []);

  const rows = useMemo(() => {
    const kw = q.trim().toLowerCase();
    const f = items.filter((i) => {
      if (kw && ![i.id, i.vendorName, i.projectName, i.itemName, i.spec, i.category].some((s) => (s ?? '').toLowerCase().includes(kw))) return false;
      if (range && (!i.schedule.date || i.schedule.date < range[0] || i.schedule.date > range[1])) return false;
      if (vendor && i.vendorName !== vendor) return false;
      if (project && (project === '__none' ? !!i.projectName : i.projectName !== project)) return false;
      if (category && i.category !== category) return false;
      if (insp === '__none' ? !!i.schedule.inspectorId || ['done', 'canceled'].includes(i.status) : insp && i.schedule.inspectorId !== insp) return false;
      if (status && i.status !== status) return false;
      if (result === 'pass' && i.result?.verdict !== 'pass') return false;
      if (result === 'fail' && i.result?.verdict !== 'fail') return false;
      if (result === 'missing' && !resultMissing(i, today)) return false;
      return true;
    });
    const val = (i: Inspection): string | number => {
      switch (sort.key) {
        case 'date': return i.schedule.date ?? '9999';
        case 'time': return i.schedule.time ?? '';
        case 'location': return i.schedule.location ?? '';
        case 'inspector': return inspectorName(i.schedule.inspectorId);
        case 'status': return statusBadge(i).label;
        case 'result': return i.result?.verdict ?? '';
        case 'quantity': return i.quantity ?? -1;
        default: return (i[sort.key] as string) ?? '';
      }
    };
    return f.sort((a, b) => {
      const x = val(a), y = val(b);
      const c = typeof x === 'number' && typeof y === 'number' ? x - y : String(x).localeCompare(String(y), 'ko');
      return (c || a.id.localeCompare(b.id)) * sort.dir;
    });
  }, [items, q, range, vendor, project, category, insp, status, result, sort, today, inspectorName]);

  const pages = Math.max(1, Math.ceil(rows.length / PAGE));
  const cur = Math.min(page, pages - 1);
  const shown = rows.slice(cur * PAGE, cur * PAGE + PAGE);
  const selected = items.filter((i) => sel.has(i.id));

  const reset = () => {
    setQ(''); setFrom(''); setTo(''); setVendor(''); setProject(''); setCategory(''); setInsp(''); setStatus(''); setResult(''); setPage(0);
  };
  const sortBy = (key: SortKey) => setSort((s) => ({ key, dir: s.key === key ? (s.dir === 1 ? -1 : 1) : 1 }));
  const Th = ({ k, children, num }: { k: SortKey; children: string; num?: boolean }) => (
    <th className={num ? 'num' : undefined} aria-sort={sort.key === k ? (sort.dir === 1 ? 'ascending' : 'descending') : undefined}>
      <button type="button" className="sort" onClick={() => sortBy(k)}
        style={{ border: 0, background: 'none', padding: 0, font: 'inherit', color: 'inherit', cursor: 'pointer' }}>
        {children}{sort.key === k ? (sort.dir === 1 ? ' ▲' : ' ▼') : ''}
      </button>
    </th>
  );

  /** 일괄 처리: 검수자 배정 → 상태 변경 순서로, 안 되는 건은 건너뛰고 알려줌 */
  const applyBulk = async (cancelReason?: string) => {
    setBulkMsg(null);
    setBulkErr(null);
    const actor = { account: user.account, displayName: user.displayName };
    const now = new Date();
    const fails: string[] = [];
    const changed = new Map<string, Inspection>();
    for (const it of selected) {
      let x = it;
      try {
        if (bulkInsp) x = assignInspector(x, bulkInsp, inspectors, actor, now);
        if (bulkStatus === 'negotiating') x = startNegotiation(x, actor, now);
        if (bulkStatus === 'confirmed') {
          // 일괄 확정은 강제 진행 없이: 충돌이 있으면 건너뜀 (이미 처리한 건을 반영한 상태로 검사)
          const working = items.map((w) => changed.get(w.id) ?? w).filter((w) => w.id !== x.id).map(toSlot);
          const cs = conflictsFor(toSlot(x), working, rules, inspectorName).map((c) => c.message);
          if (cs.length) throw new WorkflowError(`${x.id}: 충돌 조건에 걸려 건너뜀 (${cs.join(' / ')}). 상세 화면에서 사유를 적고 강제 진행할 수 있습니다.`);
          x = confirm(x, { ...x.schedule }, inspectors, actor, now);
        }
        if (bulkStatus === 'canceled') x = cancel(x, cancelReason ?? '', actor, now);
        if (x !== it) changed.set(x.id, x);
      } catch (e) {
        if (e instanceof WorkflowError) fails.push(e.message.startsWith(it.id) ? e.message : `${it.id}: ${e.message}`);
        else throw e;
      }
    }
    if (changed.size === 0) {
      setBulkMsg({ ok: '', fails: fails.length ? fails : ['바뀐 건이 없습니다. 검수자나 상태를 고르세요.'] });
      return null;
    }
    const r = await saveInspections(items.map((i) => changed.get(i.id) ?? i));
    if (r && r !== 'handled') setBulkErr(r);
    if (r === null) {
      setBulkMsg({ ok: `${changed.size}건을 처리했습니다.`, fails });
      setSel(new Set());
      setBulkInsp('');
      setBulkStatus('');
    }
    return r;
  };

  const allOnPage = shown.length > 0 && shown.every((i) => sel.has(i.id));
  const pickInspectors = inspectors.filter((p) => p.active);

  return (
    <>
      <div className="toolbar">
        <input className="input" style={{ width: 220 }} placeholder="목록 안에서 검색" aria-label="목록 안에서 검색" value={q} onChange={(e) => { setQ(e.target.value); setPage(0); }} />
        <span className="lbl muted" style={{ marginLeft: 8 }}>검수일</span>
        <select className="sel" style={{ width: 120 }} aria-label="기간" value={period} onChange={(e) => {
          const v = e.target.value;
          if (!v) { setFrom(''); setTo(''); } else if (v !== 'custom') { const [a, b] = quickRange(v as QuickPeriod, today); setFrom(a); setTo(b); }
          setPage(0);
        }}>
          <option value="">전체 기간</option>
          {QUICK.map((x) => <option key={x.v} value={x.v}>{x.label}</option>)}
          <option value="custom" disabled={period !== 'custom'}>직접 선택</option>
        </select>
        <input className="input" type="date" style={{ width: 150 }} aria-label="시작일" value={from} max={to || undefined}
          onChange={(e) => { setFrom(e.target.value); setPage(0); }} />
        <span className="muted">~</span>
        <input className="input" type="date" style={{ width: 150 }} aria-label="종료일" value={to} min={from || undefined}
          onChange={(e) => { setTo(e.target.value); setPage(0); }} />
        <div className="sp" />
        <button className="btn" type="button" disabled={rows.length === 0} onClick={async () => {
          setExportMsg(null);
          const sel = (id: string, label: string) => (document.querySelector(`select[aria-label="${id}"]`) as HTMLSelectElement | null)?.selectedOptions[0]?.text ?? label;
          const filter = [q.trim() && `검색: ${q.trim()}`, periodText && `검수일: ${periodText}`, vendor && `업체: ${vendor}`, project && `프로젝트: ${project === '__none' ? '(없음)' : project}`, category && `검수 품목: ${category}`, insp && sel('검수자', ''), status && sel('상태', ''), result && sel('결과', '')]
            .filter(Boolean).join(' · ');
          const r = await api.exportList(rows.map((i) => i.id), filter);
          if (!r.ok) setExportMsg({ ok: false, text: r.error.message });
          else if (r.value) setExportMsg({ ok: true, text: `${rows.length}건을 내보냈습니다: ${r.value}` });
        }}><Icon name="download" />엑셀 내보내기</button>
      </div>
      <div className="toolbar">
        <select className="sel" style={{ width: 128 }} aria-label="업체" value={vendor} onChange={(e) => { setVendor(e.target.value); setPage(0); }}>
          <option value="">업체: 전체</option>
          {vendors.map((v) => <option key={v}>{v}</option>)}
        </select>
        <select className="sel" style={{ width: 150 }} aria-label="프로젝트" value={project} onChange={(e) => { setProject(e.target.value); setPage(0); }}>
          <option value="">프로젝트: 전체</option>
          <option value="__none">(프로젝트 없음)</option>
          {projects.map((v) => <option key={v}>{v}</option>)}
        </select>
        <select className="sel" style={{ width: 140 }} aria-label="검수 품목" value={category} onChange={(e) => { setCategory(e.target.value); setPage(0); }}>
          <option value="">검수 품목: 전체</option>
          {categories.map((v) => <option key={v}>{v}</option>)}
        </select>
        <select className="sel" style={{ width: 128 }} aria-label="검수자" value={insp} onChange={(e) => { setInsp(e.target.value); setPage(0); }}>
          <option value="">검수자: 전체</option>
          <option value="__none">(미배정)</option>
          {inspectors.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
        </select>
        <select className="sel" style={{ width: 128 }} aria-label="상태" value={status} onChange={(e) => { setStatus(e.target.value); setPage(0); }}>
          <option value="">상태: 전체</option>
          {STATUS_OPTIONS.map((s) => <option key={s.v} value={s.v}>{s.label}</option>)}
        </select>
        <select className="sel" style={{ width: 112 }} aria-label="결과" value={result} onChange={(e) => { setResult(e.target.value); setPage(0); }}>
          <option value="">결과: 전체</option>
          <option value="pass">합격</option>
          <option value="fail">불합격</option>
          <option value="missing">미입력</option>
        </select>
        <button className="btn btn-ghost" type="button" onClick={reset}>필터 초기화</button>
      </div>

      {exportMsg && <Alert kind={exportMsg.ok ? 'ok' : 'danger'}>{exportMsg.text}</Alert>}
      {bulkMsg && (
        <>
          {bulkMsg.ok && <Alert kind="ok">{bulkMsg.ok}</Alert>}
          {bulkMsg.fails.length > 0 && (
            <Alert kind="warn">
              <b>처리하지 못한 건 {bulkMsg.fails.length}건</b>
              {bulkMsg.fails.map((f) => <span key={f} style={{ display: 'block' }}>· {f}</span>)}
            </Alert>
          )}
        </>
      )}
      {bulkErr && <ErrorAlert error={bulkErr} />}

      {sel.size > 0 && (
        <div className="alert alert-info" style={{ padding: '8px 12px', alignItems: 'center' }}>
          <span style={{ display: 'flex', alignItems: 'center', gap: 8, flex: 1 }}>
            <Icon name="info" /><b>{sel.size}건 선택됨</b>
            <button className="btn btn-ghost btn-sm" type="button" onClick={() => setSel(new Set())}>선택 해제</button>
          </span>
          <select className="sel" style={{ width: 150 }} aria-label="검수자 배정" value={bulkInsp} onChange={(e) => setBulkInsp(e.target.value)}>
            <option value="">검수자 배정…</option>
            {pickInspectors.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
          </select>
          <select className="sel" style={{ width: 150 }} aria-label="상태 변경" value={bulkStatus} onChange={(e) => setBulkStatus(e.target.value)}>
            <option value="">상태 변경…</option>
            <option value="negotiating">일정협의중</option>
            <option value="confirmed">확정 (현재 일정으로)</option>
            <option value="canceled">취소</option>
          </select>
          <button className="btn btn-primary btn-sm" type="button" disabled={!bulkInsp && !bulkStatus}
            onClick={() => (bulkStatus === 'canceled' ? setCancelOpen(true) : applyBulk())}>선택 건 적용</button>
        </div>
      )}

      {items.length === 0 ? (
        <Alert kind="info" action={<button className="btn btn-sm" type="button" onClick={() => go({ page: 'import' })}>엑셀 가져오기</button>}>
          아직 검수 건이 없습니다. 업체가 보낸 엑셀을 가져오세요.
        </Alert>
      ) : (
        <div className="tbl-wrap" style={{ overflow: 'auto' }}>
          <table className="tbl">
            <thead>
              <tr>
                <th style={{ width: 32 }}>
                  <input type="checkbox" className="chk" aria-label="이 페이지 전체 선택" checked={allOnPage}
                    onChange={() => {
                      const s = new Set(sel);
                      shown.forEach((i) => (allOnPage ? s.delete(i.id) : s.add(i.id)));
                      setSel(s);
                    }} />
                </th>
                <Th k="id">신청번호</Th>
                <Th k="vendorName">업체명</Th>
                <Th k="projectName">프로젝트명</Th>
                <Th k="category">검수 품목</Th>
                <Th k="itemName">품명/품번</Th>
                <Th k="quantity" num>수량</Th>
                <Th k="date">검수일</Th>
                <Th k="time">시간</Th>
                <Th k="location">장소</Th>
                <Th k="inspector">검수자</Th>
                <Th k="status">상태</Th>
                <Th k="result">결과</Th>
              </tr>
            </thead>
            <tbody>
              {shown.length === 0 && (
                <tr><td colSpan={13} className="muted" style={{ textAlign: 'center', padding: 24 }}>조건에 맞는 검수 건이 없습니다. [초기화]로 필터를 지우세요.</td></tr>
              )}
              {shown.map((i) => (
                <tr key={i.id} className={sel.has(i.id) ? 'sel' : undefined}>
                  <td>
                    <input type="checkbox" className="chk" aria-label={`${i.id} 선택`} checked={sel.has(i.id)}
                      onChange={() => { const s = new Set(sel); s.has(i.id) ? s.delete(i.id) : s.add(i.id); setSel(s); }} />
                  </td>
                  <td><a href="#" onClick={(e) => { e.preventDefault(); go({ page: 'detail', id: i.id, from: { page: 'list' } }); }}>{i.id}</a></td>
                  <td>{i.vendorName}</td>
                  <td title={i.projectName}>{i.projectName}</td>
                  <td>{i.category}</td>
                  <td title={itemLabel(i)}>{itemLabel(i)}</td>
                  <td className="num">{i.quantity ?? ''}</td>
                  {(() => {
                    const cs = conflictOf(i);
                    if (!cs.length) return <td>{fmtMD(i.schedule.date)}</td>;
                    return (
                      <td className="cell-warn">
                        <span className="cell-flag" title={cs.map((c) => `충돌: ${c}`).join('\n')}><Icon name="warn" />{fmtMD(i.schedule.date)} 충돌</span>
                        <span className="sr-only">{cs.join(' / ')}</span>
                      </td>
                    );
                  })()}
                  <td>{i.schedule.time ?? ''}</td>
                  <td>{i.schedule.location}</td>
                  <td>{i.schedule.inspectorId ? inspectorName(i.schedule.inspectorId) : <span className="muted">(미배정)</span>}</td>
                  <td><StatusBadge i={i} /></td>
                  <td>
                    {i.result ? <b>{i.result.verdict === 'pass' ? '합격' : '불합격'}</b>
                      : resultMissing(i, today) ? <span className="cell-flag" style={{ color: 'var(--danger-fg)', fontWeight: 700 }}><Icon name="error" />결과 미입력</span>
                      : <span className="muted">—</span>}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {items.length > 0 && (
        <div className="toolbar">
          <span className="help">총 {rows.length}건{rows.length > PAGE ? ` 중 ${cur * PAGE + 1}–${Math.min(rows.length, cur * PAGE + PAGE)}건` : ''}</span>
          <div className="sp" />
          {pages > 1 && <>
            <button className="btn btn-sm" type="button" disabled={cur === 0} onClick={() => setPage(cur - 1)}>이전</button>
            <button className="btn btn-sm" type="button" disabled={cur >= pages - 1} onClick={() => setPage(cur + 1)}>다음</button>
          </>}
        </div>
      )}

      {cancelOpen && (
        <CancelDialog label={`선택한 ${sel.size}건`} onClose={() => setCancelOpen(false)} onSubmit={(reason) => applyBulk(reason)} />
      )}
    </>
  );
}
