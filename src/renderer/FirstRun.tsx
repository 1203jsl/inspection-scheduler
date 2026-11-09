import { useEffect, useState } from 'react';
import type { AppErrorInfo, AppState, CurrentUser, FolderCheck } from '../shared/types';
import { api, fmtDateTime } from './api';
import { Alert, ErrorAlert, Icon } from './components';

/** 첫 실행(또는 데이터 폴더 변경) — 디자인 A */
export function FirstRun({ user, initialPath, notice, onDone, onCancel }: {
  user: CurrentUser;
  initialPath?: string | null;
  notice?: AppErrorInfo | null;
  onDone: (s: AppState) => void;
  onCancel?: () => void;
}) {
  const [mode, setMode] = useState<'existing' | 'new'>('existing');
  const [dir, setDir] = useState(initialPath ?? '');
  const [check, setCheck] = useState<FolderCheck | null>(null);
  const [name, setName] = useState(user.displayName === user.account ? '' : user.displayName);
  const [error, setError] = useState<AppErrorInfo | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    if (!dir) return setCheck(null);
    api.checkFolder(dir).then((r) => setCheck(r.ok ? r.value : null));
  }, [dir]);

  const pick = async () => {
    const r = await api.pickFolder();
    if (r.ok && r.value) {
      setDir(r.value);
      setError(null);
    }
  };

  const problem = (() => {
    if (!dir) return null;
    if (check && !check.exists) return '폴더를 찾을 수 없습니다. 다시 선택하세요.';
    if (check && mode === 'existing' && !check.hasData) return '이 폴더에는 검수일정 데이터가 없습니다. 다른 폴더를 고르거나 "새 데이터 폴더 만들기"를 선택하세요.';
    if (check && mode === 'new' && check.hasData) return '이 폴더에는 이미 데이터가 있습니다. "이미 쓰고 있는 데이터 폴더에 연결"을 선택하세요.';
    return null;
  })();

  const submit = async () => {
    setBusy(true);
    setError(null);
    const r = await api.setupFolder({ path: dir, mode, displayName: name });
    setBusy(false);
    if (r.ok) onDone(r.value);
    else setError(r.error);
  };

  return (
    <div style={{ minHeight: '100vh', boxSizing: 'border-box', background: 'var(--bg)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 32 }}>
      <div className="card" style={{ width: 640, padding: 0, display: 'flex', flexDirection: 'column' }}>
        <div style={{ padding: '24px 28px 16px', display: 'flex', flexDirection: 'column', gap: 4, borderBottom: '1px solid var(--border)' }}>
          <div style={{ fontSize: 13, lineHeight: '18px', color: 'var(--ink-muted)', fontWeight: 500 }}>
            검수일정 관리 · {onCancel ? '데이터 폴더 변경' : '처음 실행'}
          </div>
          <h1 style={{ margin: 0, fontSize: 22, lineHeight: '30px', fontWeight: 700 }}>데이터 폴더를 선택하세요</h1>
          <p className="help" style={{ margin: 0, fontSize: 14, lineHeight: '22px' }}>
            부서원이 함께 쓰는 클라우드독 동기화 폴더를 고릅니다. 이 설정은 이 PC에만 저장되며, 나중에 설정 화면에서 바꿀 수 있습니다.
          </p>
        </div>
        <div style={{ padding: '20px 28px', display: 'flex', flexDirection: 'column', gap: 16 }}>
          {notice && <ErrorAlert error={notice} />}
          <fieldset style={{ border: 0, margin: 0, padding: 0, display: 'flex', flexDirection: 'column', gap: 8 }}>
            <legend className="lbl" style={{ padding: '0 0 6px' }}>1. 시작 방법</legend>
            <label className="radio">
              <input type="radio" name="mode" checked={mode === 'existing'} onChange={() => setMode('existing')} />
              <span>이미 쓰고 있는 데이터 폴더에 연결 <span className="muted">(다른 부서원이 먼저 만든 경우)</span></span>
            </label>
            <label className="radio">
              <input type="radio" name="mode" checked={mode === 'new'} onChange={() => setMode('new')} />
              <span>새 데이터 폴더 만들기 <span className="muted">(부서에서 처음 사용하는 경우)</span></span>
            </label>
          </fieldset>
          <div className="field">
            <label htmlFor="dir">2. 폴더 위치 <span className="req">*</span></label>
            <div style={{ display: 'flex', gap: 8 }}>
              <input id="dir" className="input" value={dir} readOnly placeholder="[폴더 선택]을 눌러 고르세요" style={{ flex: 1 }} />
              <button className="btn" type="button" onClick={pick}><Icon name="file" />폴더 선택</button>
            </div>
            <span className="help">동기화 폴더 안의 폴더를 고르세요. 바탕화면이나 내 문서는 다른 사람과 공유되지 않습니다.</span>
          </div>
          {problem && <Alert kind="warn">{problem}</Alert>}
          {!problem && check?.hasData && mode === 'existing' && check.summary && (
            <Alert kind="ok">
              <b>데이터를 찾았습니다.</b> 검수 {check.summary.inspections}건 · 업체 {check.summary.vendors}곳 · 검수자 {check.summary.inspectors}명
              {check.summary.lastSavedAt && <> · 마지막 저장: {check.summary.lastSavedBy}, {fmtDateTime(check.summary.lastSavedAt)}</>}
            </Alert>
          )}
          {!problem && dir && mode === 'new' && (
            <Alert kind="info">이 위치에 <b>data · 원본 · 첨부 · 백업</b> 폴더를 새로 만듭니다.</Alert>
          )}
          <div className="field">
            <span className="lbl">3. 표시 이름</span>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: 12 }}>
              <div className="field">
                <label className="help" htmlFor="acc">Windows 로그인 계정</label>
                <input id="acc" className="input" value={user.account} disabled />
              </div>
              <div className="field">
                <label className="help" htmlFor="nm">화면·변경 이력에 보일 이름</label>
                <input id="nm" className="input" value={name} onChange={(e) => setName(e.target.value)} placeholder="예: 홍길동" />
              </div>
            </div>
            <span className="help">비워 두면 Windows 계정 이름을 씁니다. 설정 화면에서 바꿀 수 있습니다.</span>
          </div>
          {error && <ErrorAlert error={error} />}
        </div>
        <div className="dlg-f" style={{ borderTop: '1px solid var(--border)', padding: '14px 28px' }}>
          {onCancel && <button className="btn" type="button" onClick={onCancel}>취소</button>}
          <button className="btn btn-primary" type="button" disabled={!dir || !!problem || busy} onClick={submit}>
            {mode === 'existing' ? '연결하고 시작' : '폴더 만들고 시작'}
          </button>
        </div>
      </div>
    </div>
  );
}
