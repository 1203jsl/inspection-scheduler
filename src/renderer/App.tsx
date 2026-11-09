import { useCallback, useEffect, useState } from 'react';
import type { AppErrorInfo, AppState, DataFiles } from '../shared/types';
import { api, fmtTime, type LockStatus } from './api';
import { Alert, ConflictDialog, ErrorAlert, Icon, type IconName } from './components';
import { FirstRun } from './FirstRun';
import { SettingsPage } from './SettingsPage';
import { StatusPage } from './StatusPage';

type PageId = 'dashboard' | 'calendar' | 'list' | 'import' | 'vendors' | 'inspectors' | 'settings';
const NAV: { id: PageId; label: string; icon: IconName; step?: string }[] = [
  { id: 'dashboard', label: '대시보드', icon: 'home', step: '7단계' },
  { id: 'calendar', label: '달력', icon: 'calendar', step: '3단계' },
  { id: 'list', label: '목록', icon: 'list', step: '3단계' },
  { id: 'import', label: '엑셀 가져오기', icon: 'upload', step: '2단계' },
  { id: 'vendors', label: '업체 관리', icon: 'building', step: '3단계' },
  { id: 'inspectors', label: '검수자 관리', icon: 'user', step: '3단계' },
  { id: 'settings', label: '설정', icon: 'gear' },
];

export function App() {
  const [state, setState] = useState<AppState | null>(null);
  const [changingFolder, setChangingFolder] = useState(false);
  const [data, setData] = useState<DataFiles | null>(null);
  const [loadedAt, setLoadedAt] = useState<string>('');
  const [loadError, setLoadError] = useState<AppErrorInfo | null>(null);
  const [conflict, setConflict] = useState<AppErrorInfo | null>(null);
  const [lock, setLock] = useState<LockStatus | null>(null);
  const [page, setPage] = useState<PageId>('settings');
  const [flash, setFlash] = useState('');

  const reload = useCallback(async () => {
    const r = await api.loadAll();
    if (r.ok) {
      setData(r.value);
      setLoadedAt(new Date().toISOString());
      setLoadError(null);
    } else {
      setLoadError(r.error);
    }
    const s = await api.getState();
    if (s.ok) setState(s.value);
  }, []);

  useEffect(() => {
    api.getState().then((r) => {
      if (r.ok) {
        setState(r.value);
        if (r.value.dataDir) reload();
      }
    });
  }, [reload]);

  // 다른 PC의 "편집 중" 표시를 30초마다 확인
  useEffect(() => {
    if (!state?.dataDir) return;
    const tick = () => api.lockStatus().then((r) => r.ok && setLock(r.value));
    tick();
    const t = setInterval(tick, 30_000);
    return () => clearInterval(t);
  }, [state?.dataDir]);

  /** 저장 결과 공통 처리: 충돌이면 안내창, 그 밖의 오류는 돌려줌 */
  const handleSaveError = (e: AppErrorInfo): AppErrorInfo | null => {
    setFlash('');
    if (e.code === 'CONFLICT') {
      setConflict(e);
      return null;
    }
    return e;
  };

  if (!state) return null;

  if (!state.dataDir || changingFolder) {
    return (
      <FirstRun
        user={state.user}
        initialPath={state.dataDir}
        onCancel={state.dataDir ? () => setChangingFolder(false) : undefined}
        onDone={(s) => {
          setState(s);
          setChangingFolder(false);
          reload();
        }}
      />
    );
  }

  const nav = NAV.find((n) => n.id === page)!;
  const otherEditing = !loadError && lock?.holder && !lock.mine ? lock.holder : null;

  return (
    <div className="app">
      <aside className="side">
        <div className="brand">검수일정 관리<small>외주업체 검수</small></div>
        <nav>
          {NAV.map((n) => (
            <a key={n.id} className="nav-i" href="#" aria-current={page === n.id ? 'page' : undefined}
              onClick={(e) => { e.preventDefault(); setPage(n.id); setFlash(''); }}>
              <Icon name={n.icon} />{n.label}
            </a>
          ))}
        </nav>
        <div className="sp" />
        <div className="help" style={{ color: 'var(--on-sidebar-muted)' }}>PC 사용자: {state.user.displayName}</div>
      </aside>
      <div className="main">
        <header className="top">
          <h1>{nav.label}</h1>
          <div className="sp" />
          {loadError ? (
            <span className="help" style={{ display: 'inline-flex', gap: 6, alignItems: 'center', color: 'var(--danger-fg)' }}>
              <Icon name="error" />데이터 폴더 연결 안 됨
            </span>
          ) : (
            <span className="help" style={{ display: 'inline-flex', gap: 6, alignItems: 'center' }}>
              <Icon name="ok" style={{ color: 'var(--ok-border)' }} />동기화 폴더 연결됨 · 불러온 시각 {fmtTime(loadedAt)}
            </span>
          )}
          <button className="btn" type="button" onClick={() => { setFlash(''); reload(); }}>새로 불러오기</button>
        </header>
        <div className="page">
          {otherEditing && (
            <Alert kind="info" icon="user">
              <b>{otherEditing.displayName} 편집 중</b> · {fmtTime(otherEditing.startedAt)}부터 (다른 PC: {otherEditing.host}).
              같은 내용을 동시에 고치면 나중에 저장하는 쪽이 저장되지 않습니다.
            </Alert>
          )}
          {flash && !loadError && <Alert kind="ok">{flash}</Alert>}
          {loadError && (
            <>
              <ErrorAlert error={loadError} onRetry={reload} />
              <div className="btn-row">
                <button className="btn" type="button" onClick={() => setChangingFolder(true)}>데이터 폴더 다시 선택</button>
              </div>
            </>
          )}
          {!loadError && data && page === 'settings' && (
            <SettingsPage
              key={loadedAt}
              data={data}
              user={state.user}
              dataDir={state.dataDir}
              onChangeFolder={() => setChangingFolder(true)}
              onSaved={(msg) => { setFlash(msg ?? ''); reload(); }}
              onSaveError={handleSaveError}
              onLockChange={setLock}
            />
          )}
          {!loadError && data && page !== 'settings' && <StatusPage data={data} step={nav.step!} label={nav.label} />}
        </div>
      </div>
      {conflict && (
        <ConflictDialog
          error={conflict}
          onReload={() => {
            setConflict(null);
            api.lockRelease();
            reload();
          }}
        />
      )}
    </div>
  );
}
