import { useCallback, useEffect, useMemo, useState } from 'react';
import type { AppErrorInfo, AppState, DataFiles, Inspection } from '../shared/types';
import { ymd } from '../shared/dates';
import { api, fmtTime, type LockStatus } from './api';
import { CalendarPage } from './CalendarPage';
import { Alert, ConflictDialog, ErrorAlert, Icon, type IconName } from './components';
import { Ctx, type AppCtx, type Route } from './ctx';
import { DetailPage } from './DetailPage';
import { FirstRun } from './FirstRun';
import { ImportPage } from './ImportPage';
import { ListPage } from './ListPage';
import { MailPage } from './MailPage';
import { InspectorsPage, VendorsPage } from './PeoplePages';
import { SettingsPage } from './SettingsPage';
import { StatusPage } from './StatusPage';

type NavId = 'dashboard' | 'calendar' | 'list' | 'import' | 'vendors' | 'inspectors' | 'settings';
const NAV: { id: NavId; label: string; icon: IconName }[] = [
  { id: 'dashboard', label: '대시보드', icon: 'home' },
  { id: 'calendar', label: '달력', icon: 'calendar' },
  { id: 'list', label: '목록', icon: 'list' },
  { id: 'import', label: '엑셀 가져오기', icon: 'upload' },
  { id: 'vendors', label: '업체 관리', icon: 'building' },
  { id: 'inspectors', label: '검수자 관리', icon: 'user' },
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
  const [route, setRoute] = useState<Route>({ page: 'list' });
  const [flash, setFlash] = useState('');
  const [search, setSearch] = useState('');
  const [today, setToday] = useState(ymd(new Date()));

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
    setToday(ymd(new Date()));
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

  /** 저장 결과 공통 처리: 충돌이면 안내창, 그 밖의 오류는 돌려줌 (가져오기·설정 화면용) */
  const handleSaveError = (e: AppErrorInfo): AppErrorInfo | null => {
    setFlash('');
    if (e.code === 'CONFLICT') {
      setConflict(e);
      return null;
    }
    return e;
  };

  const go = useCallback((r: Route) => {
    setRoute(r);
    setFlash('');
  }, []);

  const ctx = useMemo((): AppCtx | null => {
    if (!data || !state) return null;
    const names = new Map(data.inspectors.items.map((p) => [p.id, p.name]));
    const save: AppCtx['save'] = async (changes, okMessage) => {
      const r = await api.save(changes);
      if (r.ok) {
        setFlash(okMessage ?? '');
        await reload();
        return null;
      }
      setFlash('');
      if (r.error.code === 'CONFLICT') {
        setConflict(r.error);
        return 'handled';
      }
      return r.error;
    };
    return {
      data, user: state.user, today, go, save,
      saveInspections: (items: Inspection[], msg?: string) => save({ inspections: { ...data.inspections, items } }, msg),
      inspectorName: (id) => (id ? names.get(id) ?? '(삭제된 검수자)' : ''),
      setLock,
      reload,
    };
  }, [data, state, today, go, reload]);

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

  const origin = (r: Route): NavId => {
    if (r.page === 'detail' || r.page === 'mail') return r.from ? origin(r.from) : 'list';
    return r.page;
  };
  const navId: NavId = origin(route);
  const title = route.page === 'detail' ? `${route.id} · 검수 건 상세`
    : route.page === 'mail' ? `${route.id} · 메일 미리보기`
    : NAV.find((n) => n.id === route.page)!.label;
  const otherEditing = !loadError && lock?.holder && !lock.mine ? lock.holder : null;

  return (
    <div className="app">
      <aside className="side">
        <div className="brand">검수일정 관리<small>외주업체 검수</small></div>
        <nav>
          {NAV.map((n) => (
            <a key={n.id} className="nav-i" href="#" aria-current={navId === n.id ? 'page' : undefined}
              onClick={(e) => { e.preventDefault(); go({ page: n.id } as Route); }}>
              <Icon name={n.icon} />{n.label}
            </a>
          ))}
        </nav>
        <div className="sp" />
        <div className="help" style={{ color: 'var(--on-sidebar-muted)' }}>PC 사용자: {state.user.displayName}</div>
      </aside>
      <div className="main">
        <header className="top">
          <h1>{title}</h1>
          <div className="sp" />
          {loadError ? (
            <span className="help" style={{ display: 'inline-flex', gap: 6, alignItems: 'center', color: 'var(--danger-fg)' }}>
              <Icon name="error" />데이터 폴더 연결 안 됨
            </span>
          ) : (
            <span className="help" style={{ display: 'inline-flex', gap: 6, alignItems: 'center' }}>
              <Icon name="ok" style={{ color: 'var(--ok-border)' }} />불러온 시각 {fmtTime(loadedAt)}
            </span>
          )}
          <button className="btn" type="button" onClick={() => { setFlash(''); reload(); }}>새로 불러오기</button>
          <form className="search" role="search" onSubmit={(e) => {
            e.preventDefault();
            go({ page: 'list', preset: { q: search.trim() } });
          }}>
            <Icon name="search" />
            <input className="input" placeholder="신청번호·업체·품명 검색" aria-label="신청번호·업체·품명 검색" value={search} onChange={(e) => setSearch(e.target.value)} />
          </form>
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
          {!loadError && ctx && (
            <Ctx.Provider value={ctx}>
              {route.page === 'list' && <ListPage key={JSON.stringify(route.preset ?? {})} preset={route.preset} />}
              {route.page === 'calendar' && <CalendarPage initialDate={route.date} />}
              {route.page === 'detail' && <DetailPage key={route.id} id={route.id} from={route.from} />}
              {route.page === 'mail' && <MailPage key={`${route.id}:${route.kind ?? ''}`} id={route.id} kind={route.kind} from={route.from} />}
              {route.page === 'vendors' && <VendorsPage initial={route.vendor} />}
              {route.page === 'inspectors' && <InspectorsPage />}
              {route.page === 'dashboard' && <StatusPage data={ctx.data} step="7단계" label="대시보드" />}
              {route.page === 'import' && (
                <ImportPage onImported={() => reload()} onSaveError={handleSaveError} onLockChange={setLock} />
              )}
              {route.page === 'settings' && (
                <SettingsPage
                  key={loadedAt}
                  data={ctx.data}
                  user={state.user}
                  dataDir={state.dataDir}
                  onChangeFolder={() => setChangingFolder(true)}
                  onSaved={(msg) => { setFlash(msg ?? ''); reload(); }}
                  onSaveError={handleSaveError}
                  onLockChange={setLock}
                />
              )}
            </Ctx.Provider>
          )}
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
