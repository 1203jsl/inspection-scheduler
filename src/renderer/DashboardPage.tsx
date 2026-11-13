import { useMemo } from 'react';
import type { Inspection, InspectionStatus } from '../shared/types';
import { MAIL_KIND_LABEL } from '../shared/types';
import { addDays, fmtMD, monthEnd, monthStart, parseYmd, weekStart } from '../shared/dates';
import { conflictMap, toSlot } from '../shared/conflicts';
import { pendingMail } from '../shared/mail';
import { upcoming } from '../shared/reminders';
import { resultMissing } from '../shared/workflow';
import { Icon, type IconName } from './components';
import { useApp, type ListPreset } from './ctx';

const dayDiff = (from: string, to: string) => Math.round((parseYmd(to).getTime() - parseYmd(from).getTime()) / 86_400_000);
const NEGO_LATE_DAYS = 3;

function DDay({ date, today }: { date: string; today: string }) {
  const d = dayDiff(today, date);
  const cls = d <= 0 ? 'dday' : d === 1 ? 'dday dday-soon' : 'dday dday-far';
  return <span className={cls}>{d <= 0 ? 'D-Day' : `D-${d}`}</span>;
}

export function DashboardPage() {
  const { data, today, go, inspectorName } = useApp();
  const items = data.inspections.items;
  const log = data.mail_log.items;
  const live = items.filter((i) => i.status !== 'canceled');

  const w0 = weekStart(today), w1 = addDays(w0, 6);
  const m0 = monthStart(today), m1 = monthEnd(today);
  const todayItems = live.filter((i) => i.schedule.date === today);
  const weekItems = live.filter((i) => i.schedule.date && i.schedule.date >= w0 && i.schedule.date <= w1);
  const monthItems = items.filter((i) => i.schedule.date && i.schedule.date >= m0 && i.schedule.date <= m1);

  const conflicts = useMemo(() => {
    const m = conflictMap(items.map(toSlot), data.settings.conflictRules, inspectorName);
    return [...m.entries()].filter(([d]) => d >= today).sort(([a], [b]) => a.localeCompare(b));
  }, [items, data.settings.conflictRules, inspectorName, today]);

  // 처리 필요
  const statusSince = (i: Inspection) => [...i.history].reverse().find((h) => h.field === 'status' && h.to === i.status)?.at;
  const registered = items.filter((i) => i.status === 'registered' || i.status === 'received');
  const negoLate = items
    .filter((i) => i.status === 'negotiating')
    .map((i) => ({ i, days: dayDiff(statusSince(i)?.slice(0, 10) ?? today, today) }))
    .filter((x) => x.days >= NEGO_LATE_DAYS);
  const noResult = items.filter((i) => resultMissing(i, today));
  const unsent = items.map((i) => ({ i, kind: pendingMail(i, log) })).filter((x) => x.kind);
  const todo = registered.length + negoLate.length + noResult.length + unsent.length;

  const next = upcoming(items, today, 7);

  const chips: { label: string; cls: string; n: number; preset: ListPreset }[] = [
    { label: '신청접수·등록', cls: 'st-received', n: monthItems.filter((i) => i.status === 'registered' || i.status === 'received').length, preset: { status: 'registered' } },
    { label: '일정협의중', cls: 'st-negotiating', n: monthItems.filter((i) => i.status === 'negotiating').length, preset: { status: 'negotiating' } },
    { label: '확정', cls: 'st-confirmed', n: monthItems.filter((i) => i.status === 'confirmed').length, preset: { status: 'confirmed' } },
    { label: '변경됨', cls: 'st-changed', n: monthItems.filter((i) => i.status === 'changed').length, preset: { status: 'changed' } },
    { label: '취소', cls: 'st-canceled', n: monthItems.filter((i) => i.status === 'canceled').length, preset: { status: 'canceled' } },
    { label: '검수완료 · 합격', cls: 'st-pass', n: monthItems.filter((i) => i.status === 'done' && i.result?.verdict !== 'fail').length, preset: { status: 'done', result: 'pass' } },
    { label: '검수완료 · 불합격', cls: 'st-fail', n: monthItems.filter((i) => i.status === 'done' && i.result?.verdict === 'fail').length, preset: { status: 'done', result: 'fail' } },
  ];
  const count = (s: InspectionStatus) => weekItems.filter((i) => i.status === s).length;
  const link = (label: string, onClick: () => void) => (
    <a className="link" href="#" onClick={(e) => { e.preventDefault(); onClick(); }}>{label}</a>
  );
  const detail = (i: Inspection) => go({ page: 'detail', id: i.id, from: { page: 'dashboard' } });

  const Group = <T,>({ title, list, icon, row }: { title: string; list: T[]; icon: IconName; row: (x: T) => [Inspection, string, React.ReactNode] }) => (
    <div>
      <div className="lbl muted" style={{ padding: '4px 0' }}>{title} ({list.length})</div>
      {list.length === 0 && <div className="help" style={{ padding: '4px 0 8px' }}>없음</div>}
      {list.slice(0, 4).map((x) => {
        const [i, text, act] = row(x);
        return (
          <div key={i.id} className="list-i">
            <Icon name={icon} />
            <span><b>{i.vendorName}</b> · {i.itemName} <span className="muted">· {text}</span></span>
            <span className="sp" />
            {act}
          </div>
        );
      })}
      {list.length > 4 && <div className="help" style={{ padding: '4px 0' }}>외 {list.length - 4}건</div>}
    </div>
  );

  return (
    <>
      <div className="cards4">
        <div className="card stat">
          <span className="lbl muted">오늘 검수</span>
          <span className="n">{todayItems.length}건</span>
          <span className="help">완료 {todayItems.filter((i) => i.status === 'done').length} · 예정 {todayItems.filter((i) => i.status !== 'done').length}</span>
        </div>
        <div className="card stat">
          <span className="lbl muted">이번 주 검수 ({fmtMD(w0).slice(0, 5).replace('-', '/')}~{fmtMD(w1).slice(0, 5).replace('-', '/')})</span>
          <span className="n">{weekItems.length}건</span>
          <span className="help">확정 {count('confirmed')} · 협의중 {count('negotiating')}{count('changed') ? ` · 변경 ${count('changed')}` : ''}</span>
        </div>
        <div className="card stat">
          <span className="lbl muted">충돌 경고 날짜</span>
          <span className="n" style={conflicts.length ? { color: 'var(--warn-fg)' } : undefined}>{conflicts.length}일</span>
          <span className="help">{conflicts.length ? `${fmtMD(conflicts[0][0])} · ${new Set(conflicts[0][1].flatMap((c) => c.ids)).size}건` : '오늘 이후 충돌 없음'}</span>
        </div>
        <div className="card stat">
          <span className="lbl muted">처리 필요</span>
          <span className="n">{todo}건</span>
          <span className="help">{todo ? '아래 목록 확인' : '모두 처리됨'}</span>
        </div>
      </div>

      <div className="card">
        <div className="card-h">상태별 건수 <span className="help" style={{ fontWeight: 400 }}>이번 달 {monthItems.length}건 · 누르면 해당 상태 목록으로 이동</span></div>
        <div className="chips">
          {chips.map((c) => (
            <a key={c.label} className="chip" href="#" onClick={(e) => { e.preventDefault(); go({ page: 'list', preset: { ...c.preset, period: 'month' } }); }}>
              <span className={`badge ${c.cls}`}>{c.label}</span><b>{c.n}건</b>
            </a>
          ))}
        </div>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1.6fr 1fr', gap: 16, minHeight: 0 }}>
        <div className="card">
          <div className="card-h"><Icon name="warn" />처리 필요<span className="more">{link('전체 목록', () => go({ page: 'list' }))}</span></div>
          {conflicts.length > 0 && (
            <div className="alert alert-warn" style={{ marginBottom: 8 }}>
              <Icon name="warn" />
              <span><b>충돌 경고</b> · {fmtMD(conflicts[0][0])} — {conflicts[0][1][0].message}{conflicts.length > 1 ? ` 외 ${conflicts.length - 1}일` : ''}</span>
              <span className="act">{link('달력에서 보기', () => go({ page: 'calendar', date: conflicts[0][0] }))}</span>
            </div>
          )}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: '0 32px' }}>
            <div>
              <Group title="등록 후 협의 전" list={registered} icon="upload"
                row={(i) => [i, `희망 ${fmtMD(i.requestedDate)}`, link('상세', () => detail(i))]} />
              <Group title={`협의 지연 (${NEGO_LATE_DAYS}일 이상)`} list={negoLate} icon="clock"
                row={(x) => [x.i, `협의 ${x.days}일째`, <span className="badge st-negotiating">일정협의중</span>]} />
            </div>
            <div>
              <Group title="결과 미입력" list={noResult} icon="error"
                row={(i) => [i, `검수일 ${fmtMD(i.schedule.date)} (${dayDiff(i.schedule.date!, today)}일 경과)`, link('결과 입력', () => detail(i))]} />
              <Group title="메일 미발송" list={unsent} icon="mail"
                row={(x) => [x.i, `${MAIL_KIND_LABEL[x.kind!]} 안내`, link('메일 보내기', () => go({ page: 'mail', id: x.i.id, from: { page: 'dashboard' } }))]} />
            </div>
          </div>
        </div>
        <div className="card">
          <div className="card-h"><Icon name="calendar" />다가오는 검수 <span className="help" style={{ fontWeight: 400 }}>7일 이내 · 확정</span>
            <span className="more">{link('달력 보기', () => go({ page: 'calendar' }))}</span>
          </div>
          {next.length === 0 && <span className="help">7일 안에 확정된 검수가 없습니다.</span>}
          {next.slice(0, 8).map((i) => (
            <div key={i.id} className="list-i">
              <DDay date={i.schedule.date!} today={today} />
              <span>
                <a className="link" href="#" onClick={(e) => { e.preventDefault(); detail(i); }}><b>{i.vendorName}</b> · {i.itemName}</a>
                <span className="muted"> · {fmtMD(i.schedule.date)} {i.schedule.time ?? ''} · {i.schedule.location}{i.schedule.inspectorId ? ` · ${inspectorName(i.schedule.inspectorId)}` : ''}</span>
              </span>
            </div>
          ))}
          {next.length > 8 && <div className="help" style={{ paddingTop: 4 }}>외 {next.length - 8}건</div>}
        </div>
      </div>
    </>
  );
}
