import { useState, type DragEvent } from 'react';
import type { AppErrorInfo, Attachment, AttachmentKind, Inspection } from '../shared/types';
import { ATTACHMENT_KINDS } from '../shared/types';
import { fmtStamp, fmtYMD } from '../shared/dates';
import { can, checkResult, recordResult, resultMissing, WorkflowError, type ResultInput } from '../shared/workflow';
import { api } from './api';
import { Alert, Dialog, ErrorAlert, Icon } from './components';
import { useApp, useEditLock, type Route } from './ctx';

function FieldErr({ msg }: { msg?: string }) {
  return msg ? <span className="help help-err"><Icon name="error" />{msg}</span> : null;
}

const daysBetween = (a: string, b: string) => Math.round((Date.parse(b) - Date.parse(a)) / 86_400_000);

// ───────── 검수 결과 ─────────

export function ResultPanel({ insp, here }: { insp: Inspection; here: Route }) {
  const { data, today, user, go, saveInspections } = useApp();
  const inspectors = data.inspectors.items;
  const r = insp.result;
  const initial: ResultInput = {
    verdict: r?.verdict ?? null,
    date: r?.date ?? (insp.schedule.date && insp.schedule.date <= today ? insp.schedule.date : today),
    inspectorId: r?.inspectorId ?? insp.schedule.inspectorId,
    failReason: r?.failReason ?? '',
    memo: r?.memo ?? '',
  };
  const [v, setV] = useState<ResultInput>(initial);
  const [tried, setTried] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<AppErrorInfo | null>(null);
  const dirty = JSON.stringify(v) !== JSON.stringify(initial);
  useEditLock(dirty);

  if (!can(insp, 'result')) {
    return <Alert kind="info">결과는 일정이 확정된 건에만 입력할 수 있습니다. 지금 상태: {insp.status === 'canceled' ? '취소' : '확정 전'}.</Alert>;
  }

  const { errors, warnings } = checkResult(v, inspectors, today);
  const show = (k: keyof ResultInput) => (tried ? errors[k] : undefined);
  const pickable = inspectors.filter((p) => p.active || p.id === v.inspectorId);
  const verdictLabel = v.verdict === 'pass' ? '검수완료 · 합격' : v.verdict === 'fail' ? '검수완료 · 불합격' : '';

  const save = async (thenMail: boolean) => {
    setTried(true);
    setError(null);
    if (Object.keys(errors).length) return;
    setBusy(true);
    try {
      const next = recordResult(insp, v, inspectors, { account: user.account, displayName: user.displayName }, new Date(), today);
      const res = await saveInspections(data.inspections.items.map((i) => (i.id === next.id ? next : i)),
        `${insp.id}: 결과(${v.verdict === 'pass' ? '합격' : '불합격'})를 저장했습니다.${thenMail ? '' : ' 업체에 결과를 알리려면 [메일 보내기]를 누르세요.'}`);
      if (res === null && thenMail) go({ page: 'mail', id: insp.id, kind: v.verdict === 'fail' ? 'fail' : 'pass', from: here });
      else if (res && res !== 'handled') setError(res);
    } catch (e) {
      if (e instanceof WorkflowError) setError({ code: 'UNKNOWN', message: e.message });
      else throw e;
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
      <div className="card-h" style={{ margin: 0 }}>{r ? '검수 결과 수정' : '검수 결과 입력'}
        {r && <span className="help" style={{ fontWeight: 400 }}>마지막 입력: {r.by} · {fmtStamp(r.at)}</span>}
      </div>
      {resultMissing(insp, today) && (
        <Alert kind="warn" icon="clock">검수일({fmtYMD(insp.schedule.date)})이 {daysBetween(insp.schedule.date!, today)}일 지났습니다. 결과를 입력하고 업체에 통보하세요.</Alert>
      )}
      <fieldset style={{ border: 0, margin: 0, padding: 0, display: 'flex', gap: 24, alignItems: 'center', flexWrap: 'wrap' }}>
        <legend className="lbl" style={{ padding: '0 0 6px' }}>판정 <span className="req">*</span></legend>
        <label className="radio"><input type="radio" name="verdict" checked={v.verdict === 'pass'} onChange={() => setV({ ...v, verdict: 'pass' })} />합격</label>
        <label className="radio"><input type="radio" name="verdict" checked={v.verdict === 'fail'} onChange={() => setV({ ...v, verdict: 'fail' })} />불합격</label>
        {v.verdict && <span className={`badge ${v.verdict === 'pass' ? 'st-pass' : 'st-fail'}`}>{verdictLabel}</span>}
        <FieldErr msg={show('verdict')} />
      </fieldset>
      <div className="form-row">
        <div className="field">
          <label htmlFor="rd">실제 검수일 <span className="req">*</span></label>
          <input id="rd" type="date" className={`input${show('date') ? ' invalid' : ''}`} value={v.date ?? ''} onChange={(e) => setV({ ...v, date: e.target.value || null })} />
          <FieldErr msg={show('date')} />
        </div>
        <div className="field">
          <label htmlFor="ri">검수자 <span className="req">*</span></label>
          <select id="ri" className={`sel${show('inspectorId') ? ' invalid' : ''}`} value={v.inspectorId ?? ''} onChange={(e) => setV({ ...v, inspectorId: e.target.value || null })}>
            <option value="">선택하세요</option>
            {pickable.map((p) => <option key={p.id} value={p.id}>{p.name}{p.dept ? ` (${p.dept})` : ''}{p.active ? '' : ' · 사용 안 함'}</option>)}
          </select>
          <FieldErr msg={show('inspectorId')} />
        </div>
      </div>
      {v.verdict === 'fail' && (
        <div className="field">
          <label htmlFor="rf">불합격 사유 <span className="req">*</span></label>
          <textarea id="rf" className={`input${show('failReason') ? ' invalid' : ''}`} style={{ minHeight: 64 }} value={v.failReason}
            onChange={(e) => setV({ ...v, failReason: e.target.value })} placeholder="예: 외경 치수 공차 초과 (규격 Ø52.0±0.1, 측정 52.18) — 40개 중 6개" />
          {show('failReason') ? <FieldErr msg={show('failReason')} /> : <span className="help">업체 통보 메일의 {'{불합격사유}'}에 그대로 들어갑니다.</span>}
        </div>
      )}
      <div className="field">
        <label htmlFor="rm">메모 <span className="muted">(내부용, 메일에 들어가지 않음)</span></label>
        <textarea id="rm" className="input" style={{ minHeight: 56 }} value={v.memo} onChange={(e) => setV({ ...v, memo: e.target.value })} />
      </div>
      {warnings.map((w) => <Alert key={w} kind="warn">{w}</Alert>)}
      {error && <ErrorAlert error={error} />}
      <div className="btn-row">
        <span className="help" style={{ marginRight: 'auto' }}>
          {v.verdict ? `저장하면 상태가 '${verdictLabel}'(으)로 바뀌고 변경 이력에 기록됩니다.` : '판정을 고르세요.'}
        </span>
        <button className="btn" type="button" disabled={busy || (!!r && !dirty)} onClick={() => save(false)}>저장</button>
        <button className="btn btn-primary" type="button" disabled={busy || (!!r && !dirty)} onClick={() => save(true)}>
          <Icon name="mail" />저장 후 결과 메일 미리보기
        </button>
      </div>
    </div>
  );
}

// ───────── 첨부파일 ─────────

const fmtSize = (n: number) => (n < 1024 ? `${n} B` : n < 1024 * 1024 ? `${(n / 1024).toFixed(0)} KB` : `${(n / 1024 / 1024).toFixed(1)} MB`);

export function FilesPanel({ insp, compact }: { insp: Inspection; compact?: boolean }) {
  const { reload } = useApp();
  const [kind, setKind] = useState<AttachmentKind>('검수자료');
  const [drag, setDrag] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<AppErrorInfo | null>(null);
  const [notice, setNotice] = useState('');
  const [del, setDel] = useState<Attachment | null>(null);
  const list = insp.attachments ?? [];

  const add = async (paths: string[]) => {
    if (!paths.length) return;
    setBusy(true);
    setError(null);
    setNotice('');
    const r = await api.attachAdd(insp.id, paths, kind);
    setBusy(false);
    if (r.ok) {
      setNotice(`${r.value.length}개 파일을 '${kind}'(으)로 첨부했습니다.`);
      await reload();
    } else setError(r.error);
  };
  const onDrop = (e: DragEvent) => {
    e.preventDefault();
    setDrag(false);
    add([...e.dataTransfer.files].map((f) => api.pathForFile(f)).filter(Boolean));
  };
  const remove = async () => {
    if (!del) return;
    const r = await api.attachRemove(insp.id, del.id);
    setDel(null);
    if (r.ok) {
      setNotice(`${del.name}을(를) 삭제했습니다. 변경 이력에 남았고, 파일은 첨부/${insp.id}/_삭제됨 폴더에 보관했습니다.`);
      await reload();
    } else setError(r.error);
  };
  const open = async (file: string) => {
    const r = await api.openDataFile(file);
    if (!r.ok) setError(r.error);
  };

  return (
    <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 12, alignSelf: 'start' }}>
      <div className="card-h" style={{ margin: 0 }}><Icon name="file" />첨부파일{list.length ? ` (${list.length})` : ''}</div>
      <div className="dropzone" onDragOver={(e) => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)} onDrop={onDrop}
        style={drag ? { borderColor: 'var(--primary)', background: 'var(--primary-soft)' } : undefined}>
        <Icon name="upload" lg style={{ color: 'var(--ink-muted)' }} />
        <span style={{ fontSize: 13, lineHeight: '20px', flex: 1 }}>
          파일을 끌어다 놓거나 선택하세요.<br /><span className="muted">첨부/{insp.id}/ 에 복사됩니다.</span>
        </span>
        <button className="btn btn-sm" type="button" disabled={busy} onClick={async () => { const r = await api.attachPick(); if (r.ok) add(r.value); }}>파일 선택</button>
      </div>
      <div className="field" style={{ maxWidth: 240 }}>
        <label htmlFor={`kind-${compact ? 'c' : 'f'}`}>등록할 파일 종류</label>
        <select id={`kind-${compact ? 'c' : 'f'}`} className="sel" value={kind} onChange={(e) => setKind(e.target.value as AttachmentKind)}>
          {ATTACHMENT_KINDS.map((k) => <option key={k}>{k}</option>)}
        </select>
      </div>
      {busy && <Alert kind="info" icon="clock">복사하는 중입니다…</Alert>}
      {notice && <Alert kind="ok">{notice}</Alert>}
      {error && <ErrorAlert error={error} />}
      {list.length > 0 && (
        <div className="tbl-wrap" style={{ overflow: 'auto' }}>
          <table className="tbl">
            <thead><tr><th>종류</th><th>파일명</th>{!compact && <th className="num">크기</th>}<th>등록</th><th /></tr></thead>
            <tbody>
              {list.map((a) => (
                <tr key={a.id}>
                  <td><span className="tag">{a.kind}</span></td>
                  <td style={{ maxWidth: compact ? 170 : 360 }} title={a.file.split('/').pop()}>{a.file.split('/').pop()}</td>
                  {!compact && <td className="num">{fmtSize(a.size)}</td>}
                  <td className="muted">{a.addedBy} · {fmtStamp(a.addedAt).slice(0, 5)}</td>
                  <td>
                    <div className="btn-row" style={{ flexWrap: 'nowrap' }}>
                      <button className="btn btn-sm" type="button" onClick={() => open(a.file)}>열기</button>
                      <button className="btn btn-sm btn-danger" type="button" onClick={() => setDel(a)}>삭제</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      <span className="help">삭제해도 변경 이력에 파일명·삭제한 사람·시각이 남습니다.</span>
      {!compact && insp.sources.length > 0 && (
        <>
          <div className="lbl" style={{ marginTop: 4 }}>업체 원본 엑셀</div>
          {insp.sources.map((s, k) => (
            <div key={k} className="list-i">
              <Icon name="file" />
              <a className="link" href="#" onClick={(e) => { e.preventDefault(); open(s.file); }}>{s.file.split('/').pop()}</a>
              <span className="muted">{s.row}행 · {s.importedBy} · {fmtStamp(s.importedAt)}</span>
            </div>
          ))}
        </>
      )}
      {del && (
        <Dialog title="첨부 삭제" footer={<>
          <button className="btn" type="button" onClick={() => setDel(null)}>닫기</button>
          <button className="btn btn-danger" type="button" onClick={remove}>삭제</button>
        </>}>
          <span><b>{del.name}</b> ({del.kind})을(를) 첨부 목록에서 삭제합니다.</span>
          <span className="help">파일은 첨부/{insp.id}/_삭제됨 폴더로 옮겨 보관하고, 변경 이력에 기록합니다.</span>
        </Dialog>
      )}
    </div>
  );
}
