import { useMemo, useState } from 'react';
import type { Inspection } from '../shared/types';
import { addDays, addMonths, monthEnd, monthStart, parseYmd, weekStart } from '../shared/dates';
import { statusBadge } from '../shared/workflow';
import { conflictMap, toSlot } from '../shared/conflicts';
import { Icon } from './components';
import { useApp } from './ctx';

const HEAD = ['월', '화', '수', '목', '금', '토', '일'];
const MAX_IN_MONTH = 3;
const LEGEND: { cls: string; label: string }[] = [
  { cls: 'st-received', label: '신청접수·등록' },
  { cls: 'st-negotiating', label: '일정협의중' },
  { cls: 'st-confirmed', label: '확정' },
  { cls: 'st-changed', label: '변경됨' },
  { cls: 'st-canceled', label: '취소' },
  { cls: 'st-pass', label: '검수완료 · 합격' },
  { cls: 'st-fail', label: '검수완료 · 불합격' },
];
const evBtn: React.CSSProperties = { border: 0, width: '100%', textAlign: 'left', cursor: 'pointer', font: 'inherit', fontSize: 13, lineHeight: '20px' };
const plainBtn: React.CSSProperties = { border: 0, background: 'none', padding: 0, font: 'inherit', color: 'inherit', cursor: 'pointer' };

export function CalendarPage({ initialDate }: { initialDate?: string }) {
  const { data, today, go, inspectorName } = useApp();
  const [view, setView] = useState<'month' | 'week'>('month');
  const [anchor, setAnchor] = useState(initialDate ?? today);
  const [vendor, setVendor] = useState('');
  const [insp, setInsp] = useState('');
  const [status, setStatus] = useState('');
  const items = data.inspections.items;
  const vendors = useMemo(() => [...new Set(items.map((i) => i.vendorName).filter(Boolean))].sort(), [items]);

  const byDay = useMemo(() => {
    const m = new Map<string, Inspection[]>();
    for (const i of items) {
      if (!i.schedule.date) continue;
      if (vendor && i.vendorName !== vendor) continue;
      if (insp && i.schedule.inspectorId !== insp) continue;
      if (status && i.status !== status) continue;
      (m.get(i.schedule.date) ?? m.set(i.schedule.date, []).get(i.schedule.date)!).push(i);
    }
    for (const l of m.values()) l.sort((a, b) => (a.schedule.time ?? '').localeCompare(b.schedule.time ?? ''));
    return m;
  }, [items, vendor, insp, status]);

  // 충돌 판정은 필터와 상관없이 전체 일정 기준
  const cmap = useMemo(() => conflictMap(items.map(toSlot), data.settings.conflictRules, inspectorName), [items, data.settings.conflictRules, inspectorName]);

  const days = useMemo(() => {
    if (view === 'week') {
      const s = weekStart(anchor);
      return Array.from({ length: 7 }, (_, k) => addDays(s, k));
    }
    const first = weekStart(monthStart(anchor));
    const last = addDays(weekStart(monthEnd(anchor)), 6);
    const out: string[] = [];
    for (let d = first; d <= last; d = addDays(d, 1)) out.push(d);
    return out;
  }, [view, anchor]);

  const month = anchor.slice(0, 7);
  const a = parseYmd(anchor);
  const title = view === 'month'
    ? `${a.getFullYear()}년 ${a.getMonth() + 1}월`
    : (() => { const s = parseYmd(days[0]), e = parseYmd(days[6]); return `${s.getFullYear()}년 ${s.getMonth() + 1}월 ${s.getDate()}일 – ${e.getMonth() + 1}월 ${e.getDate()}일`; })();
  const move = (n: number) => setAnchor(view === 'month' ? addMonths(anchor, n) : addDays(anchor, 7 * n));
  const openDay = (d: string) => go({ page: 'list', preset: { day: d } });
  const evLabel = (i: Inspection) => `${i.vendorName} · ${i.itemName}`;

  return (
    <>
      <div className="toolbar">
        <div className="seg" role="group" aria-label="보기 전환">
          <button type="button" aria-pressed={view === 'month'} onClick={() => setView('month')}>월</button>
          <button type="button" aria-pressed={view === 'week'} onClick={() => setView('week')}>주</button>
        </div>
        <button className="btn btn-sm" type="button" aria-label={view === 'month' ? '이전 달' : '이전 주'} onClick={() => move(-1)}>이전</button>
        <b style={{ font: '700 18px/26px var(--font-sans)' }} aria-live="polite">{title}</b>
        <button className="btn btn-sm" type="button" aria-label={view === 'month' ? '다음 달' : '다음 주'} onClick={() => move(1)}>다음</button>
        <button className="btn btn-sm" type="button" onClick={() => setAnchor(today)}>오늘</button>
        <div className="sp" />
        <select className="sel" style={{ width: 128 }} aria-label="업체" value={vendor} onChange={(e) => setVendor(e.target.value)}>
          <option value="">업체: 전체</option>
          {vendors.map((v) => <option key={v}>{v}</option>)}
        </select>
        <select className="sel" style={{ width: 128 }} aria-label="검수자" value={insp} onChange={(e) => setInsp(e.target.value)}>
          <option value="">검수자: 전체</option>
          {data.inspectors.items.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
        </select>
        <select className="sel" style={{ width: 128 }} aria-label="상태" value={status} onChange={(e) => setStatus(e.target.value)}>
          <option value="">상태: 전체</option>
          <option value="registered">등록</option>
          <option value="negotiating">일정협의중</option>
          <option value="confirmed">확정</option>
          <option value="changed">변경됨</option>
          <option value="done">검수완료</option>
          <option value="canceled">취소</option>
        </select>
      </div>

      <div className="legend">
        {LEGEND.map((l) => <span key={l.cls} className={`badge ${l.cls}`}>{l.label}</span>)}
      </div>

      <div className="cal" role="grid" aria-label={title}>
        {HEAD.map((h) => <div key={h} className="cal-h" role="columnheader">{h}</div>)}
        {days.map((d) => {
          const list = byDay.get(d) ?? [];
          const out = view === 'month' && d.slice(0, 7) !== month;
          const shown = view === 'month' ? list.slice(0, MAX_IN_MONTH) : list;
          const conf = cmap.get(d) ?? [];
          const confIds = new Set(conf.flatMap((c) => c.ids));
          const cls = ['cal-c', out ? 'out' : '', d === today ? 'today' : '', conf.length ? 'conf' : ''].filter(Boolean).join(' ');
          const n = parseYmd(d).getDate();
          return (
            <div key={d} className={cls} role="gridcell" style={view === 'week' ? { minHeight: 480 } : undefined}>
              <div className="d">
                <button type="button" style={plainBtn} onClick={() => openDay(d)} aria-label={`${d} 검수 목록 보기 (${list.length}건)`}>
                  {d === today ? <b>{n}</b> : n}
                </button>
                {conf.length > 0 && (
                  <span className="cell-flag" title={conf.map((c) => c.message).join('\n')}>
                    <Icon name="warn" />충돌 {confIds.size}건
                    <span className="sr-only">{conf.map((c) => c.message).join(' / ')}</span>
                  </span>
                )}
              </div>
              {shown.map((i) => {
                const b = statusBadge(i);
                const text = view === 'week'
                  ? `${i.schedule.time ?? ''} ${evLabel(i)}${i.schedule.inspectorId ? ` · ${inspectorName(i.schedule.inspectorId)}` : ''}`
                  : evLabel(i);
                return (
                  <button key={i.id} type="button" className={`ev ${b.cls}`} style={{ ...evBtn, whiteSpace: view === 'week' ? 'normal' : 'nowrap' }}
                    title={`${i.id} · ${i.schedule.time ?? ''} · ${evLabel(i)} · ${b.label}`}
                    onClick={() => go({ page: 'detail', id: i.id, from: { page: 'calendar', date: anchor } })}>
                    {text}
                  </button>
                );
              })}
              {view === 'month' && list.length > MAX_IN_MONTH && (
                <button type="button" className="more-ev" style={{ ...plainBtn, color: 'var(--primary)', textAlign: 'left' }} onClick={() => openDay(d)}>+{list.length - MAX_IN_MONTH}건 더보기</button>
              )}
            </div>
          );
        })}
      </div>
    </>
  );
}
