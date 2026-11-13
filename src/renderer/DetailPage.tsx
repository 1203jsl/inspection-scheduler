import { useState } from 'react';
import type { AppErrorInfo, HistoryEntry, Inspection, InspectionStatus, MailLogEntry } from '../shared/types';
import { MAIL_KIND_LABEL, STATUS_LABEL } from '../shared/types';
import { mailCovers, pendingMail } from '../shared/mail';
import { fmtStamp, fmtYMD } from '../shared/dates';
import {
  addNote, can, cancel, change, confirm, editRequest, FIELD_LABEL, nextAction, resultMissing, startNegotiation, WorkflowError,
} from '../shared/workflow';
import { api } from './api';
import { Alert, ErrorAlert, Icon } from './components';
import { useApp, type Route } from './ctx';
import { CancelDialog, EditInfoDialog, ScheduleDialog } from './dialogs';
import { StatusBadge } from './ListPage';
import { FilesPanel, ResultPanel } from './DetailTabs';

type Tab = 'info' | 'nego' | 'result' | 'files' | 'mail' | 'history';
type Dlg = null | 'confirm' | 'change' | 'cancel' | 'edit';

const dash = (v: unknown) => (v === null || v === undefined || v === '' ? '—' : String(v));

export function DetailPage({ id, from }: { id: string; from?: Route }) {
  const app = useApp();
  const { data, today, go, inspectorName, saveInspections, user } = app;
  const insp = data.inspections.items.find((i) => i.id === id);
  const [tab, setTab] = useState<Tab>('info');
  const [dlg, setDlg] = useState<Dlg>(null);
  const [note, setNote] = useState('');
  const [error, setError] = useState<AppErrorInfo | null>(null);

  if (!insp) {
    return (
      <Alert kind="warn" action={<button className="btn btn-sm" type="button" onClick={() => go({ page: 'list' })}>목록으로</button>}>
        신청번호 {id}를 찾을 수 없습니다. 다른 사람이 백업을 복원했을 수 있습니다. 목록에서 다시 찾으세요.
      </Alert>
    );
  }

  const actor = { account: user.account, displayName: user.displayName };
  /** 이 건만 바꿔 저장 */
  const commit = (next: Inspection, msg: string) =>
    saveInspections(data.inspections.items.map((i) => (i.id === next.id ? next : i)), msg);
  const act = async (fn: () => Inspection, msg: string) => {
    setError(null);
    try {
      const r = await commit(fn(), msg);
      if (r && r !== 'handled') setError(r);
    } catch (e) {
      if (e instanceof WorkflowError) setError({ code: 'UNKNOWN', message: e.message });
      else throw e;
    }
  };

  const mails = data.mail_log.items.filter((m) => mailCovers(m, insp.id));
  const pending = pendingMail(insp, data.mail_log.items);
  // 보낼 안내 메일이 남아 있으면 [메일 보내기]가 다음 할 일
  const primary = pending ? null : nextAction(insp, today);
  const here: Route = { page: 'detail', id: insp.id, from };
  const btn = (key: 'negotiate' | 'confirm' | 'change', label: string, onClick: () => void) =>
    can(insp, key) && (
      <button key={key} className={`btn${primary === key ? ' btn-primary' : ''}`} type="button" onClick={onClick}>{label}</button>
    );

  const TABS: { id: Tab; label: string }[] = [
    { id: 'info', label: '기본정보' },
    { id: 'nego', label: `일정 협의${insp.negotiations.length ? ` (${insp.negotiations.length})` : ''}` },
    { id: 'result', label: '검수 결과' },
    { id: 'files', label: `첨부파일${(insp.attachments ?? []).length ? ` (${insp.attachments.length})` : ''}` },
    { id: 'mail', label: `메일 이력${mails.length ? ` (${mails.length})` : ''}` },
    { id: 'history', label: `변경 이력 (${insp.history.length})` },
  ];

  return (
    <>
      <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
          <button className="btn btn-ghost btn-sm" type="button" onClick={() => go(from ?? { page: 'list' })} aria-label="이전 화면으로">
            <Icon name="chev" style={{ transform: 'rotate(180deg)' }} />돌아가기
          </button>
          <div>
            <div className="help">신청번호 {insp.id}</div>
            <div style={{ font: '700 22px/30px var(--font-sans)', display: 'flex', alignItems: 'center', gap: 8 }}>
              {insp.vendorName} · {insp.itemName} <StatusBadge i={insp} />
            </div>
          </div>
          <div className="sp" />
          <div className="btn-row">
            {btn('negotiate', '협의 시작', () => act(() => startNegotiation(insp, actor, new Date()), `${insp.id}: 일정협의중으로 바꿨습니다.`))}
            {btn('confirm', insp.status === 'changed' ? '재확정' : '확정', () => setDlg('confirm'))}
            {btn('change', '변경', () => setDlg('change'))}
            {can(insp, 'result') && (
              <button className={`btn${primary === 'result' ? ' btn-primary' : ''}`} type="button" onClick={() => setTab('result')}>
                {insp.result ? '결과 수정' : '결과 입력'}
              </button>
            )}
            {can(insp, 'cancel') && <button className="btn btn-danger" type="button" onClick={() => setDlg('cancel')}>취소</button>}
            <button className={`btn${pending ? ' btn-primary' : ''}`} type="button" onClick={() => go({ page: 'mail', id: insp.id, from: here })}>
              <Icon name="mail" />메일 보내기
            </button>
          </div>
        </div>
        {pending && <Alert kind="info"><b>안내</b> · {MAIL_KIND_LABEL[pending]} 안내 메일이 아직 발송되지 않았습니다.</Alert>}
        {insp.status === 'changed' && <Alert kind="info">일정이 변경되었습니다. 업체와 다시 맞춘 뒤 [재확정]하세요.</Alert>}
        {resultMissing(insp, today) && tab !== 'result' && (
          <Alert kind="danger" action={<button className="btn btn-sm" type="button" onClick={() => setTab('result')}>결과 입력</button>}>
            검수일({fmtYMD(insp.schedule.date)})이 지났는데 결과가 없습니다.
          </Alert>
        )}
        {error && <ErrorAlert error={error} />}
      </div>

      <div className="tabs" role="tablist">
        {TABS.map((t) => (
          <button key={t.id} className="tab" role="tab" type="button" aria-selected={tab === t.id} onClick={() => setTab(t.id)}>{t.label}</button>
        ))}
      </div>

      <div role="tabpanel">
        {tab === 'info' && (
          <div style={{ display: 'grid', gridTemplateColumns: '1.5fr 1fr', gap: 16, alignItems: 'start' }}>
            <div className="card">
              <div className="card-h">업체 신청 내용 vs 현재 내용
                <button className="btn btn-sm more" type="button" onClick={() => setDlg('edit')}>기본정보 수정</button>
              </div>
              <InfoTable insp={insp} inspectorName={inspectorName} />
              {insp.confirmedAt && <p className="help" style={{ margin: '8px 0 0' }}>마지막 확정: {insp.confirmedBy} · {fmtStamp(insp.confirmedAt)}</p>}
            </div>
            <div className="card">
              <div className="card-h">일정 협의 기록
                <a className="more link" href="#" onClick={(e) => { e.preventDefault(); setTab('nego'); }}>전체 보기</a>
              </div>
              {insp.negotiations.length === 0 ? <span className="help">아직 기록이 없습니다.</span>
                : insp.negotiations.slice(-5).map((n, k) => <NoteRow key={k} n={n} />)}
            </div>
          </div>
        )}

        {tab === 'nego' && (
          <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            <div className="card-h" style={{ margin: 0 }}>일정 협의 기록</div>
            {insp.negotiations.length === 0 ? <span className="help">아직 기록이 없습니다. 업체와 통화·메일로 협의한 내용을 남겨 두세요.</span>
              : insp.negotiations.map((n, k) => <NoteRow key={k} n={n} />)}
            <div className="field">
              <label htmlFor="note">협의 내용 추가</label>
              <textarea id="note" className="input" value={note} onChange={(e) => setNote(e.target.value)} placeholder="예: 업체 요청 — 11/10 입고 지연으로 11/11 검수 희망" />
            </div>
            <div className="btn-row">
              <button className="btn btn-primary" type="button" disabled={!note.trim()}
                onClick={async () => { await act(() => addNote(insp, note, actor, new Date()), '협의 기록을 추가했습니다.'); setNote(''); }}>
                기록 추가
              </button>
            </div>
          </div>
        )}

        {tab === 'result' && (
          <div className="split" style={{ gridTemplateColumns: 'minmax(0, 1fr) 520px' }}>
            <ResultPanel key={insp.result?.at ?? 'new'} insp={insp} here={here} />
            <FilesPanel insp={insp} compact />
          </div>
        )}

        {tab === 'files' && <FilesPanel insp={insp} />}

        {tab === 'mail' && <MailHistory mails={mails} />}

        {tab === 'history' && <HistoryTable insp={insp} inspectorName={inspectorName} />}
      </div>

      {(dlg === 'confirm' || dlg === 'change') && (
        <ScheduleDialog
          mode={dlg}
          insp={insp}
          onClose={() => setDlg(null)}
          onSubmit={async (s, reason, force) => commit(
            dlg === 'confirm'
              ? confirm(insp, s, data.inspectors.items, actor, new Date(), undefined, force)
              : change(insp, s, data.inspectors.items, reason, actor, new Date(), force),
            dlg === 'confirm' ? `${insp.id}: 확정했습니다.` : `${insp.id}: 일정을 변경했습니다. 다시 확정해야 합니다.`,
          )}
        />
      )}
      {dlg === 'cancel' && (
        <CancelDialog label={insp.id} onClose={() => setDlg(null)}
          onSubmit={async (reason) => commit(cancel(insp, reason, actor, new Date()), `${insp.id}: 취소했습니다.`)} />
      )}
      {dlg === 'edit' && (
        <EditInfoDialog insp={insp} onClose={() => setDlg(null)}
          onSubmit={async (patch, reason) => commit(editRequest(insp, patch, actor, new Date(), reason.trim() || undefined), '기본정보를 수정했습니다.')} />
      )}
    </>
  );
}

function NoteRow({ n }: { n: Inspection['negotiations'][number] }) {
  return (
    <div style={{ display: 'flex', gap: 12, padding: '8px 0', borderBottom: '1px solid var(--border)', fontSize: 13 }}>
      <span className="muted" style={{ width: 96, flex: 'none' }}>{fmtStamp(n.at)}</span>
      <span style={{ width: 64, flex: 'none' }}><b>{n.byName}</b></span>
      <span style={{ whiteSpace: 'pre-wrap' }}>{n.text}</span>
    </div>
  );
}

function InfoTable({ insp, inspectorName }: { insp: Inspection; inspectorName: (id: string | null) => string }) {
  const o = insp.original;
  const when = (d: string | null, t: string | null) => (d ? `${fmtYMD(d)}${t ? ` ${t}` : ''}` : '—');
  const contact = (n: string, e: string) => [n, e].filter(Boolean).join(' · ') || '—';
  const rows: [string, string, string][] = [
    ['희망일 / 검수일', when(o.requestedDate, o.requestedTime), when(insp.schedule.date, insp.schedule.time)],
    ['검수 품목', dash(o.category), dash(insp.category)],
    ['품명', dash(o.itemName), dash(insp.itemName)],
    ['품번/규격', dash(o.spec), dash(insp.spec)],
    ['수량', dash(o.quantity), dash(insp.quantity)],
    ['장소', dash(o.location), dash(insp.schedule.location)],
    ['검수자', '—', insp.schedule.inspectorId ? `${inspectorName(insp.schedule.inspectorId)} (배정)` : '(미배정)'],
    ['업체 담당자', contact(o.contactName, o.contactEmail), contact(insp.contactName, insp.contactEmail)],
    ['연락처', dash(o.contactPhone), dash(insp.contactPhone)],
    ['비고', dash(o.note), dash(insp.note)],
  ];
  return (
    <div className="kv">
      <div className="h">항목</div><div className="h">업체 신청(원본)</div><div className="h">현재</div>
      {rows.map(([k, a, b]) => {
        const diff = k !== '검수자' && a !== b;
        return [
          <div key={k + 'k'} className="k">{k}</div>,
          <div key={k + 'a'}>{a}</div>,
          <div key={k + 'b'} className={diff ? 'chg' : undefined}>
            {diff ? <><Icon name="warn" /> {b} · 변경됨</> : b}
          </div>,
        ];
      })}
    </div>
  );
}

function HistoryTable({ insp, inspectorName }: { insp: Inspection; inspectorName: (id: string | null) => string }) {
  const show = (h: HistoryEntry, v: unknown) => {
    if (v === null || v === undefined || v === '') return '—';
    if (h.field === 'status' || h.action === 'status') return STATUS_LABEL[v as InspectionStatus] ?? String(v);
    if (h.field === 'schedule.inspectorId' || h.field === 'result.inspectorId') return inspectorName(String(v));
    return String(v);
  };
  const what = (h: HistoryEntry) =>
    h.action === 'created' ? '등록' : h.action === 'status' ? '상태' : h.action === 'attach' ? '첨부 추가' : h.action === 'detach' ? '첨부 삭제'
      : FIELD_LABEL[h.field ?? ''] ?? h.field ?? h.action;
  const list = [...insp.history].reverse();
  return (
    <div className="tbl-wrap" style={{ overflow: 'auto' }}>
      <table className="tbl">
        <thead><tr><th>일시</th><th>작성자</th><th>항목</th><th>이전값</th><th>새 값</th><th>사유</th></tr></thead>
        <tbody>
          {list.map((h, k) => (
            <tr key={k}>
              <td>{fmtStamp(h.at)}</td>
              <td>{h.byName || h.by}</td>
              <td>{what(h)}</td>
              <td className="muted">{h.action === 'created' ? '' : show(h, h.from)}</td>
              <td><b>{h.action === 'created' ? '' : show(h, h.to)}</b></td>
              <td className="diff" title={h.reason}>{h.reason ?? ''}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

const RESULT: Record<MailLogEntry['result'], { icon: 'ok' | 'error'; label: string; color: string }> = {
  sent: { icon: 'ok', label: '발송 성공', color: 'var(--ok-fg)' },
  manual: { icon: 'ok', label: '수동 발송 기록', color: 'var(--ok-fg)' },
  failed: { icon: 'error', label: '발송 실패', color: 'var(--danger-fg)' },
};

function MailHistory({ mails }: { mails: MailLogEntry[] }) {
  const [open, setOpen] = useState<string | null>(null);
  if (mails.length === 0) return <Alert kind="info">아직 보낸 메일이 없습니다. [메일 보내기]로 미리보기를 연 뒤 보내세요.</Alert>;
  const list = [...mails].reverse();
  return (
    <div className="tbl-wrap" style={{ overflow: 'auto' }}>
      <table className="tbl">
        <thead><tr><th>발송 일시</th><th>종류</th><th>받는 사람</th><th>제목</th><th>방식</th><th>결과</th><th>보낸 사람</th><th /></tr></thead>
        <tbody>
          {list.map((m) => {
            const r = RESULT[m.result];
            return [
              <tr key={m.id}>
                <td>{fmtStamp(m.at)}</td>
                <td>{m.kind === 'test' ? '테스트' : MAIL_KIND_LABEL[m.kind]}</td>
                <td title={m.to}>{m.to}</td>
                <td title={m.subject}>{m.subject}</td>
                <td>{m.method === 'smtp' ? '자동' : '수동'}</td>
                <td><span className="cell-flag" style={{ color: r.color, fontWeight: 700 }}><Icon name={r.icon} />{r.label}</span></td>
                <td>{m.byName}</td>
                <td><button className="btn btn-sm btn-ghost" type="button" aria-expanded={open === m.id} onClick={() => setOpen(open === m.id ? null : m.id)}>{open === m.id ? '닫기' : '내용 보기'}</button></td>
              </tr>,
              open === m.id && (
                <tr key={m.id + ':b'}>
                  <td colSpan={8} className="diff" style={{ background: 'var(--surface-alt)' }}>
                    {m.error && <div style={{ color: 'var(--danger-fg)', marginBottom: 8 }}><Icon name="error" /> {m.error}</div>}
                    {m.cc && <div className="help">참조: {m.cc}</div>}
                    {m.attachments.length > 0 && <div className="help">첨부: {m.attachments.join(', ')}</div>}
                    <div style={{ whiteSpace: 'pre-wrap', marginTop: 8 }}>{m.body}</div>
                  </td>
                </tr>
              ),
            ];
          })}
        </tbody>
      </table>
    </div>
  );
}
