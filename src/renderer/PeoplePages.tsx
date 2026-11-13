import { useMemo, useState } from 'react';
import { api } from './api';
import { Dialog } from './components';
import type { AppErrorInfo, Inspector, Vendor } from '../shared/types';
import { fmtMD, monthEnd, monthStart } from '../shared/dates';
import { Alert, ErrorAlert, Icon } from './components';
import { useApp } from './ctx';
import { InspectorDialog, VendorDialog } from './dialogs';
import { StatusBadge } from './ListPage';

// ───────── 업체 관리 ─────────

export function VendorsPage({ initial }: { initial?: string }) {
  const { data, today, go, save } = useApp();
  const [q, setQ] = useState('');
  const [selName, setSelName] = useState(initial ?? data.vendors.items[0]?.name ?? '');
  const [dlg, setDlg] = useState<null | 'add' | 'edit'>(null);
  const [sched, setSched] = useState<null | { from: string; to: string }>(null);
  const [msg, setMsg] = useState<{ ok: boolean; text: string } | null>(null);
  const items = data.inspections.items;
  const m0 = monthStart(today), m1 = monthEnd(today);

  const rows = useMemo(() => {
    const kw = q.trim().toLowerCase();
    return data.vendors.items
      .filter((v) => !kw || [v.name, v.contactName, v.contactEmail].some((s) => s.toLowerCase().includes(kw)))
      .map((v) => {
        const mine = items.filter((i) => i.vendorName === v.name);
        const dates = mine.map((i) => i.schedule.date).filter(Boolean) as string[];
        return {
          v,
          last: dates.sort().at(-1) ?? null,
          month: mine.filter((i) => i.schedule.date && i.schedule.date >= m0 && i.schedule.date <= m1 && i.status !== 'canceled').length,
          total: mine.length,
        };
      })
      .sort((a, b) => a.v.name.localeCompare(b.v.name, 'ko'));
  }, [data.vendors.items, items, q, m0, m1]);

  const selected = data.vendors.items.find((v) => v.name === selName) ?? null;
  const history = items
    .filter((i) => i.vendorName === selName)
    .sort((a, b) => (b.schedule.date ?? '').localeCompare(a.schedule.date ?? ''));

  const submit = async (v: Vendor): Promise<AppErrorInfo | 'handled' | null> => {
    const now = new Date().toISOString();
    const list = dlg === 'edit' && selected
      ? data.vendors.items.map((x) => (x === selected ? { ...v, updatedAt: now } : x))
      : [...data.vendors.items, { ...v, updatedAt: now }];
    const r = await save({ vendors: { ...data.vendors, items: list } }, `업체 '${v.name}'을(를) 저장했습니다.`);
    if (r === null) setSelName(v.name);
    return r;
  };

  return (
    <>
      <div className="toolbar">
        <input className="input" style={{ width: 240 }} placeholder="업체명·담당자 검색" aria-label="업체 검색" value={q} onChange={(e) => setQ(e.target.value)} />
        <div className="sp" />
        <button className="btn btn-primary" type="button" onClick={() => setDlg('add')}>업체 추가</button>
      </div>
      {data.vendors.items.length === 0 ? (
        <Alert kind="info">등록된 업체가 없습니다. 엑셀을 가져오면 업체가 자동으로 추가됩니다.</Alert>
      ) : (
        <div className="split" style={{ gridTemplateColumns: 'minmax(0, 1fr) 440px' }}>
          <div className="tbl-wrap" style={{ alignSelf: 'start', overflow: 'auto' }}>
            <table className="tbl">
              <thead><tr><th>업체명</th><th>담당자</th><th>이메일</th><th>연락처</th><th>최근 검수일</th><th className="num">검수 건수(이번 달)</th></tr></thead>
              <tbody>
                {rows.map(({ v, last, month }) => (
                  <tr key={v.name} className={v.name === selName ? 'sel' : undefined}>
                    <td><a href="#" onClick={(e) => { e.preventDefault(); setSelName(v.name); }}>{v.name}</a></td>
                    <td>{v.contactName}</td>
                    <td>{v.contactEmail || <span className="muted">(없음)</span>}</td>
                    <td>{v.contactPhone}</td>
                    <td>{fmtMD(last)}</td>
                    <td className="num">{month}건</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {selected && (
            <div className="card" style={{ alignSelf: 'start' }}>
              <div className="card-h">{selected.name} · 검수 이력
                <span className="more btn-row">
                  <a className="link" href="#" onClick={(e) => { e.preventDefault(); go({ page: 'list', preset: { vendor: selected.name } }); }}>목록에서 보기</a>
                  <button className="btn btn-sm" type="button" onClick={() => { setMsg(null); setSched({ from: m0, to: m1 }); }}>일정 확인서</button>
                  <button className="btn btn-sm" type="button" onClick={() => setDlg('edit')}>수정</button>
                </span>
              </div>
              <div className="help" style={{ marginBottom: 8 }}>
                담당자 {selected.contactName || '—'} · {selected.contactEmail || '이메일 없음'} · 전체 {history.length}건
              </div>
              {history.length === 0 && <span className="help">검수 이력이 없습니다.</span>}
              {history.map((i) => (
                <div key={i.id} className="list-i">
                  <a className="link" href="#" onClick={(e) => { e.preventDefault(); go({ page: 'detail', id: i.id, from: { page: 'vendors', vendor: selected.name } }); }}>{i.id}</a>
                  <span><b>{i.itemName}</b></span>
                  <span className="muted">{fmtMD(i.schedule.date)} {i.schedule.time ?? ''}</span>
                  <span className="sp" />
                  <StatusBadge i={i} />
                </div>
              ))}
            </div>
          )}
        </div>
      )}
      {msg && <Alert kind={msg.ok ? 'ok' : 'danger'}>{msg.text}</Alert>}
      {sched && selected && (
        <Dialog title={`일정 확인서 · ${selected.name}`} footer={<>
          <button className="btn" type="button" onClick={() => setSched(null)}>닫기</button>
          <button className="btn btn-primary" type="button" disabled={!sched.from || !sched.to || sched.from > sched.to} onClick={async () => {
            const r = await api.exportVendor({ vendor: selected.name, ...sched });
            if (!r.ok) setMsg({ ok: false, text: r.error.message });
            else if (r.value) setMsg({ ok: true, text: `일정 확인서를 저장했습니다: ${r.value}` });
            if (!r.ok || r.value) setSched(null);
          }}><Icon name="download" />엑셀로 저장</button>
        </>}>
          <span className="help">이 기간에 잡힌 {selected.name}의 검수 일정을 확인서(엑셀)로 만듭니다. 업체에 보내거나 출력할 때 씁니다.</span>
          <div className="form-row">
            <div className="field"><label htmlFor="sf">시작일</label><input id="sf" type="date" className="input" value={sched.from} onChange={(e) => setSched({ ...sched, from: e.target.value })} /></div>
            <div className="field"><label htmlFor="st2">종료일</label><input id="st2" type="date" className={`input${sched.from > sched.to ? ' invalid' : ''}`} value={sched.to} onChange={(e) => setSched({ ...sched, to: e.target.value })} /></div>
          </div>
          {sched.from > sched.to && <span className="help help-err"><Icon name="error" />종료일이 시작일보다 앞입니다.</span>}
        </Dialog>
      )}
      {dlg && (
        <VendorDialog
          vendor={dlg === 'edit' ? selected : null}
          locked={dlg === 'edit' && !!selected && items.some((i) => i.vendorName === selected.name)}
          onClose={() => setDlg(null)}
          onSubmit={submit}
        />
      )}
    </>
  );
}

// ───────── 검수자 관리 ─────────

const ACTIVE_STATUSES = ['received', 'registered', 'negotiating', 'confirmed', 'changed'];

export function InspectorsPage() {
  const { data, today, go, save } = useApp();
  const [dlg, setDlg] = useState<null | { p: Inspector | null }>(null);
  const [error, setError] = useState<AppErrorInfo | null>(null);
  const items = data.inspections.items;
  const m0 = monthStart(today), m1 = monthEnd(today);
  const [, mm] = today.split('-');

  const monthItems = items.filter((i) => i.schedule.date && i.schedule.date >= m0 && i.schedule.date <= m1 && i.status !== 'canceled');
  const count = (id: string) => monthItems.filter((i) => i.schedule.inspectorId === id).length;
  const max = Math.max(1, ...data.inspectors.items.map((p) => count(p.id)));
  const unassigned = items.filter((i) => ACTIVE_STATUSES.includes(i.status) && !i.schedule.inspectorId).length;
  const assignedMonth = monthItems.filter((i) => i.schedule.inspectorId).length;

  const write = (list: Inspector[], msg: string) => save({ inspectors: { ...data.inspectors, items: list } }, msg);

  const toggle = async (p: Inspector) => {
    setError(null);
    const r = await write(data.inspectors.items.map((x) => (x.id === p.id ? { ...x, active: !x.active } : x)),
      `${p.name}: ${p.active ? '사용 안 함' : '사용'}으로 바꿨습니다.`);
    if (r && r !== 'handled') setError(r);
  };

  const submit = async (p: Inspector) => {
    const list = p.id
      ? data.inspectors.items.map((x) => (x.id === p.id ? p : x))
      : [...data.inspectors.items, { ...p, id: `P-${Date.now().toString(36)}`, createdAt: new Date().toISOString() }];
    return write(list, `검수자 '${p.name}'을(를) 저장했습니다.`);
  };

  return (
    <>
      {unassigned > 0 && (
        <Alert kind="warn" action={<a className="link" href="#" onClick={(e) => { e.preventDefault(); go({ page: 'list', preset: { unassigned: true } }); }}>목록에서 보기</a>}>
          <b>미배정 {unassigned}건</b> · 검수자가 정해지지 않은 건이 있습니다.
        </Alert>
      )}
      {error && <ErrorAlert error={error} />}
      <div className="toolbar">
        <span className="muted">이번 달({today.slice(0, 4)}년 {Number(mm)}월) 배정 {assignedMonth}건 · 미배정 {unassigned}건</span>
        <div className="sp" />
        <button className="btn btn-primary" type="button" onClick={() => setDlg({ p: null })}>검수자 추가</button>
      </div>
      {data.inspectors.items.length === 0 ? (
        <Alert kind="info">등록된 검수자가 없습니다. [검수자 추가]로 사내 검수자를 등록하세요. 확정하려면 검수자 배정이 필요합니다.</Alert>
      ) : (
        <div className="tbl-wrap">
          <table className="tbl">
            <thead><tr><th>이름</th><th>부서</th><th>이메일</th><th>사용 여부</th><th className="num">이번 달 배정</th><th>배정 현황</th><th /></tr></thead>
            <tbody>
              {data.inspectors.items.map((p) => (
                <tr key={p.id}>
                  <td><b>{p.name}</b></td>
                  <td>{p.dept}</td>
                  <td>{p.email}</td>
                  <td>
                    <label className="switch-row">
                      <input type="checkbox" className="switch" role="switch" checked={p.active} aria-label={`${p.name} 사용 여부`} onChange={() => toggle(p)} />
                      <span>{p.active ? '사용' : '사용 안 함'}</span>
                    </label>
                  </td>
                  <td className="num">{count(p.id)}건</td>
                  <td><span className="bar" aria-hidden="true"><i style={{ width: `${(count(p.id) / max) * 100}%` }} /></span></td>
                  <td><button className="btn btn-sm" type="button" onClick={() => setDlg({ p })}>수정</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {dlg && <InspectorDialog inspector={dlg.p} onClose={() => setDlg(null)} onSubmit={submit} />}
    </>
  );
}
