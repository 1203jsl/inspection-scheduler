import { useMemo, useState, type DragEvent } from 'react';
import type { AppErrorInfo, CellIssue, FieldKey, ImportPreview, ImportSummary, PreviewRow } from '../shared/types';
import { api, type LockStatus } from './api';
import { Alert, Dialog, ErrorAlert, Icon } from './components';

const COLS: { key: FieldKey; label: string; num?: boolean }[] = [
  { key: 'id', label: '신청번호' },
  { key: 'vendorName', label: '업체명' },
  { key: 'category', label: '검수 품목' },
  { key: 'itemName', label: '품명' },
  { key: 'spec', label: '품번/규격' },
  { key: 'quantity', label: '수량', num: true },
  { key: 'requestedDate', label: '희망일' },
  { key: 'requestedTime', label: '시간' },
  { key: 'location', label: '장소' },
  { key: 'contactName', label: '담당자' },
  { key: 'contactPhone', label: '연락처' },
  { key: 'contactEmail', label: '이메일' },
  { key: 'note', label: '비고' },
];

const KIND = {
  new: <span className="tag tag-new">신규</span>,
  update: <span className="tag tag-mod">수정</span>,
  error: <span className="tag tag-err"><Icon name="error" />오류</span>,
};

function Flagged({ issues, children, tipUp }: { issues: CellIssue[]; children: React.ReactNode; tipUp: boolean }) {
  const err = issues.some((i) => i.level === 'error');
  const msg = issues.map((i) => i.message).join(' / ');
  return (
    <span className="cell-flag">
      <Icon name={err ? 'error' : 'warn'} />
      {children}
      <span className={`tip${tipUp ? ' up' : ''}`} role="tooltip">
        {issues.map((i, n) => (
          <span key={n} style={{ display: 'flex', gap: 4 }}><Icon name={i.level === 'error' ? 'error' : 'warn'} style={{ marginTop: 2 }} />{i.message}</span>
        ))}
      </span>
      <span className="sr-only">{msg}</span>
    </span>
  );
}

export function ImportPage({ onImported, onSaveError, onLockChange }: {
  onImported: () => void;
  onSaveError: (e: AppErrorInfo) => AppErrorInfo | null;
  onLockChange: (l: LockStatus) => void;
}) {
  const [preview, setPreview] = useState<ImportPreview | null>(null);
  const [include, setInclude] = useState<Set<string>>(new Set());
  const [busy, setBusy] = useState(false);
  const [drag, setDrag] = useState(false);
  const [error, setError] = useState<AppErrorInfo | null>(null);
  const [notice, setNotice] = useState('');
  const [summary, setSummary] = useState<ImportSummary | null>(null);
  const [confirmErrors, setConfirmErrors] = useState(false);

  const load = async (paths: string[]) => {
    if (!paths.length) return;
    setBusy(true);
    setError(null);
    setSummary(null);
    setNotice('');
    const r = await api.importPreview(paths);
    setBusy(false);
    if (!r.ok) return setError(r.error);
    setPreview(r.value);
    setInclude(new Set(r.value.rows.filter((x) => x.defaultInclude).map((x) => x.key)));
    if (r.value.rows.length) api.lockAcquire().then((l) => l.ok && onLockChange(l.value));
  };

  const pick = async () => {
    const r = await api.pickImportFiles();
    if (r.ok) load(r.value);
  };

  const onDrop = (e: DragEvent) => {
    e.preventDefault();
    setDrag(false);
    const paths = [...e.dataTransfer.files].map((f) => api.pathForFile(f)).filter(Boolean);
    load(paths);
  };

  const downloadTemplate = async () => {
    const r = await api.saveTemplate();
    if (!r.ok) setError(r.error);
    else if (r.value) setNotice(`빈 양식을 저장했습니다: ${r.value}`);
  };

  const rows = preview?.rows ?? [];
  const counts = useMemo(() => {
    const c = { new: 0, update: 0, error: 0, incNew: 0, incUpd: 0, incErr: 0 };
    for (const r of rows) {
      c[r.kind]++;
      if (include.has(r.key)) {
        if (r.kind === 'error') c.incErr++;
        if (r.target === 'new') c.incNew++;
        else c.incUpd++;
      }
    }
    return c;
  }, [rows, include]);
  const incCount = include.size;
  const allChecked = rows.length > 0 && incCount === rows.length;

  const toggle = (key: string) => {
    const s = new Set(include);
    if (s.has(key)) s.delete(key);
    else s.add(key);
    setInclude(s);
  };

  const cancel = () => {
    setPreview(null);
    setInclude(new Set());
    api.lockRelease();
  };

  const commit = async () => {
    setConfirmErrors(false);
    setBusy(true);
    setError(null);
    const r = await api.importCommit([...include]);
    setBusy(false);
    if (r.ok) {
      setSummary(r.value);
      setPreview(null);
      await api.lockRelease();
      onImported();
    } else {
      const rest = onSaveError(r.error);
      if (rest) setError(rest);
      else setPreview(null); // 충돌: 새로 불러온 뒤 파일을 다시 넣어야 함
    }
  };

  const fileNames = preview?.files.map((f) => f.name).join(', ');

  return (
    <>
      <div
        className="dropzone"
        onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={onDrop}
        style={drag ? { borderColor: 'var(--primary)', background: 'var(--primary-soft)' } : undefined}
      >
        <Icon name="upload" lg />
        <div style={{ minWidth: 0 }}>
          {preview ? (
            <>
              <b style={{ wordBreak: 'break-all' }}>{fileNames}</b>
              <div className="help">{rows.length}행을 읽었습니다 · 다른 파일을 끌어다 놓으면 다시 읽습니다</div>
            </>
          ) : (
            <>
              <b>업체가 보낸 엑셀(.xlsx)을 여기에 끌어다 놓으세요</b>
              <div className="help">여러 파일을 한 번에 넣을 수 있습니다. 가져오기 전에 행별로 미리 확인합니다.</div>
            </>
          )}
        </div>
        <div className="sp" />
        <div className="btn-row" style={{ flexWrap: 'nowrap' }}>
          <button className="btn" type="button" onClick={downloadTemplate}><Icon name="download" />빈 양식 내려받기</button>
          <button className="btn" type="button" onClick={pick} disabled={busy}><Icon name="file" />파일 선택</button>
        </div>
      </div>

      {busy && <Alert kind="info" icon="clock">처리 중입니다…</Alert>}
      {notice && <Alert kind="ok">{notice}</Alert>}
      {error && <ErrorAlert error={error} />}

      {summary && (
        <>
          <Alert kind="ok">
            <b>가져오기 완료</b> · 신규 {summary.created.length}건, 수정 {summary.updated.length}건, 제외 {summary.excluded}건을 가져왔습니다.
            원본은 {summary.originals.map((o, i) => (
              <span key={o}>{i > 0 && ', '}<a className="link" onClick={() => api.openDataFile(o)}>{o}</a></span>
            ))}에 보관했습니다.
          </Alert>
          {summary.created.length > 0 && (
            <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
              <div className="card-h" style={{ margin: 0 }}>새로 부여한 신청번호 <span className="help" style={{ fontWeight: 400 }}>업체에 회신할 번호입니다 (메일 발송은 5단계)</span></div>
              <div className="tbl-wrap">
                <table className="tbl">
                  <thead><tr><th>신청번호</th><th>업체명</th><th>품명</th><th>받는 사람</th></tr></thead>
                  <tbody>
                    {summary.created.map((c) => (
                      <tr key={c.id}><td><b>{c.id}</b></td><td>{c.vendorName}</td><td>{c.itemName}</td><td>{c.contactEmail}</td></tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
          {summary.updated.length > 0 && <span className="help">수정한 건: {summary.updated.join(', ')} (변경 이력에 기존값 → 새 값으로 기록)</span>}
        </>
      )}

      {preview && (
        <>
          {preview.files.filter((f) => f.error).map((f) => <ErrorAlert key={f.path} error={f.error!} />)}
          {preview.files.filter((f) => f.warnings?.length).map((f) => (
            <Alert key={f.path} kind="warn">{f.name}: {f.warnings!.join(' ')}</Alert>
          ))}

          {rows.length === 0 ? (
            <Alert kind="info">가져올 행이 없습니다. 업체가 3행부터 작성했는지 확인하세요.</Alert>
          ) : (
            <>
              <div className="toolbar">
                <div className="legend">
                  <span>{KIND.new} {counts.new}건</span>
                  <span>{KIND.update} {counts.update}건</span>
                  <span>{KIND.error} {counts.error}건</span>
                  <span className="cell-warn" style={{ padding: '0 8px', borderRadius: 4 }}><Icon name="warn" /> 경고 칸</span>
                  <span className="cell-err" style={{ padding: '0 8px', borderRadius: 4 }}><Icon name="error" /> 오류 칸</span>
                </div>
                <div className="sp" />
                <span className="help">칸에 마우스를 올리면 사유가 표시됩니다</span>
              </div>

              <div className="tbl-wrap" style={{ overflow: 'auto', flex: '1 1 auto', minHeight: 200 }}>
                <table className="tbl">
                  <thead>
                    <tr>
                      <th style={{ width: 32 }}>
                        <input type="checkbox" className="chk" aria-label="전체 포함" checked={allChecked}
                          onChange={() => setInclude(allChecked ? new Set() : new Set(rows.map((r) => r.key)))} />
                      </th>
                      <th>판정</th>
                      <th>파일·행</th>
                      {COLS.map((c) => <th key={c.key} className={c.num ? 'num' : undefined}>{c.label}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((r, idx) => (
                      <PreviewTr key={r.key} row={r} checked={include.has(r.key)} onToggle={() => toggle(r.key)}
                        tipUp={idx >= rows.length - 3 && rows.length > 4} />
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="toolbar" style={{ justifyContent: 'space-between' }}>
                <span>
                  <b>포함 {incCount}건</b> (신규 {counts.incNew} · 수정 {counts.incUpd}) · 제외 {rows.length - incCount}건
                  {counts.error > 0 && <span className="help"> · 오류 행은 기본 제외입니다. 확인 후 포함할 수 있습니다.</span>}
                </span>
                <div className="btn-row">
                  <button className="btn" type="button" onClick={cancel}>취소</button>
                  <button className="btn btn-primary" type="button" disabled={incCount === 0 || busy}
                    onClick={() => (counts.incErr > 0 ? setConfirmErrors(true) : commit())}>
                    {incCount}건 가져오기
                  </button>
                </div>
              </div>
            </>
          )}
        </>
      )}

      {confirmErrors && (
        <Dialog
          title="오류가 있는 행 포함"
          icon={<Icon name="warn" lg style={{ color: 'var(--warn-fg)' }} />}
          footer={<>
            <button className="btn" type="button" onClick={() => setConfirmErrors(false)}>돌아가기</button>
            <button className="btn btn-primary" type="button" onClick={commit}>그대로 가져오기</button>
          </>}
        >
          <Alert kind="warn">오류가 있는 {counts.incErr}행을 포함했습니다.</Alert>
          <span>형식이 잘못된 칸(날짜·시간·수량)은 빈 값으로 저장되고, 없는 신청번호는 새 번호로 등록됩니다. 가져온 뒤 상세 화면에서 고칠 수 있습니다.</span>
        </Dialog>
      )}
    </>
  );
}

function PreviewTr({ row, checked, onToggle, tipUp }: { row: PreviewRow; checked: boolean; onToggle: () => void; tipUp: boolean }) {
  const rowIssues = row.issues.row ?? [];
  return (
    <tr className={checked ? undefined : 'muted'}>
      <td><input type="checkbox" className="chk" aria-label={`${row.rowNumber}행 가져오기 포함`} checked={checked} onChange={onToggle} /></td>
      <td className={rowIssues.length ? 'cell-warn has-tip' : undefined}> 
        {rowIssues.length ? <Flagged issues={rowIssues} tipUp={tipUp}>{KIND[row.kind]}</Flagged> : KIND[row.kind]}
      </td>
      <td className="muted" title={row.fileName}>{row.fileName.length > 14 ? row.fileName.slice(0, 12) + '…' : row.fileName} · {row.rowNumber}</td>
      {COLS.map((c) => {
        const issues = row.issues[c.key] ?? [];
        const d = c.key !== 'id' ? row.diff?.[c.key as Exclude<FieldKey, 'id'>] : undefined;
        const value = c.key === 'id' ? (row.parsed.id || <span className="muted">(신규 부여)</span>)
          : d ? <span style={{ whiteSpace: 'nowrap' }}><span className="diff-old">{d.from}</span> → <span className="diff-new">{d.to}</span></span>
          : row.display[c.key] || (issues.length ? '(비어 있음)' : '');
        const cls = [
          c.num ? 'num' : '',
          d ? 'diff' : '',
          issues.length ? (issues.some((i) => i.level === 'error') ? 'cell-err' : 'cell-warn') + ' has-tip' : '',
        ].filter(Boolean).join(' ');
        return (
          <td key={c.key} className={cls || undefined}>
            {issues.length ? <Flagged issues={issues} tipUp={tipUp}>{value}</Flagged> : value}
          </td>
        );
      })}
    </tr>
  );
}
