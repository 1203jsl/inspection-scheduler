import { useEffect, useState } from 'react';
import type { AppErrorInfo, BackupEntry, CurrentUser, DataFiles } from '../shared/types';
import { api, fmtDateTime, type LockStatus } from './api';
import { Alert, Dialog, ErrorAlert, Icon } from './components';

/**
 * 설정 (1단계 범위): 데이터 폴더, 표시 이름 매핑, 잠금 유효 시간, 백업 보관·복원.
 * 충돌 조건·메일·알림 설정은 해당 단계에서 이 화면에 추가합니다.
 */
export function SettingsPage({ data, user, dataDir, onChangeFolder, onSaved, onSaveError, onLockChange }: {
  data: DataFiles;
  user: CurrentUser;
  dataDir: string;
  onChangeFolder: () => void;
  onSaved: (message?: string) => void;
  onSaveError: (e: AppErrorInfo) => AppErrorInfo | null;
  onLockChange: (l: LockStatus) => void;
}) {
  const s = data.settings;
  const [userMap, setUserMap] = useState<[string, string][]>(Object.entries(s.userMap ?? {}));
  const [lockMin, setLockMin] = useState(String(s.lockStaleMinutes));
  const [keep, setKeep] = useState(String(s.backupKeep));
  const [items, setItems] = useState<string[]>(s.inspectionItems ?? []);
  const [newItem, setNewItem] = useState('');
  const [dirty, setDirty] = useState(false);
  const [error, setError] = useState<AppErrorInfo | null>(null);
  const [backups, setBackups] = useState<BackupEntry[]>([]);
  const [restoreTarget, setRestoreTarget] = useState<BackupEntry | null>(null);

  useEffect(() => {
    api.listBackups().then((r) => r.ok && setBackups(r.value));
  }, [data]);

  /** 처음 고칠 때 "편집 중" 표시 */
  const touch = () => {
    if (!dirty) {
      setDirty(true);
      api.lockAcquire().then((r) => r.ok && onLockChange(r.value));
    }
  };

  const lockN = Number(lockMin);
  const keepN = Number(keep);
  const invalid = {
    lock: !Number.isInteger(lockN) || lockN < 5 || lockN > 1440 ? '5~1440 사이의 정수(분)로 입력하세요.' : '',
    keep: !Number.isInteger(keepN) || keepN < 1 || keepN > 500 ? '1~500 사이의 정수로 입력하세요.' : '',
    map: userMap.some(([a, n]) => !a.trim() || !n.trim()) ? '계정과 표시 이름을 모두 입력하세요. 필요 없는 줄은 삭제하세요.' : '',
    items: items.length === 0 ? '검수 품목을 하나 이상 등록하세요. 빈 양식의 드롭다운에 쓰입니다.' : '',
  };
  const trimmed = newItem.trim();
  const dupItem = trimmed && items.includes(trimmed) ? '이미 있는 품목입니다.' : '';
  const addItem = () => {
    if (!trimmed || dupItem) return;
    touch();
    setItems([...items, trimmed]);
    setNewItem('');
  };
  const hasInvalid = Object.values(invalid).some(Boolean);

  const save = async () => {
    setError(null);
    const map: Record<string, string> = {};
    for (const [a, n] of userMap) map[a.trim()] = n.trim();
    const r = await api.save({ settings: { ...s, userMap: map, lockStaleMinutes: lockN, backupKeep: keepN, inspectionItems: items } });
    if (r.ok) {
      await api.lockRelease();
      setDirty(false);
      onSaved('저장했습니다.');
    } else {
      setError(onSaveError(r.error));
    }
  };

  const cancel = async () => {
    await api.lockRelease();
    onSaved(); // 다시 불러와 원래 값으로
  };

  const doRestore = async () => {
    if (!restoreTarget) return;
    const r = await api.restoreBackup(restoreTarget.id);
    setRestoreTarget(null);
    if (r.ok) {
      onSaved(`${fmtDateTime(restoreTarget.createdAt)} 시점으로 복원했습니다. 복원 전 상태도 백업해 두었습니다.`);
    } else setError(r.error);
  };

  return (
    <>
      {error && <ErrorAlert error={error} />}

      <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <div className="card-h" style={{ margin: 0 }}>데이터 폴더 <span className="help" style={{ fontWeight: 400 }}>이 PC에만 저장</span></div>
        <div style={{ display: 'flex', gap: 8 }}>
          <input className="input" value={dataDir} readOnly aria-label="데이터 폴더" style={{ flex: 1 }} />
          <button className="btn" type="button" onClick={onChangeFolder}>변경</button>
        </div>
      </div>

      <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <div className="card-h" style={{ margin: 0 }}>사용자 표시 이름 <span className="help" style={{ fontWeight: 400 }}>Windows 계정 → 화면·변경 이력에 보일 이름 (모든 PC 공통)</span></div>
        <div className="tbl-wrap">
          <table className="tbl">
            <thead><tr><th style={{ width: '40%' }}>Windows 계정</th><th>표시 이름</th><th style={{ width: 80 }} /></tr></thead>
            <tbody>
              {userMap.map(([a, n], i) => (
                <tr key={i} className={a === user.account ? 'sel' : undefined}>
                  <td style={{ maxWidth: 'none' }}>
                    <input className="input" aria-label="Windows 계정" value={a}
                      onChange={(e) => { touch(); setUserMap(userMap.map((r, j) => (j === i ? [e.target.value, r[1]] : r))); }} />
                  </td>
                  <td style={{ maxWidth: 'none' }}>
                    <input className="input" aria-label="표시 이름" value={n}
                      onChange={(e) => { touch(); setUserMap(userMap.map((r, j) => (j === i ? [r[0], e.target.value] : r))); }} />
                  </td>
                  <td><button className="btn btn-sm btn-danger" type="button" onClick={() => { touch(); setUserMap(userMap.filter((_, j) => j !== i)); }}>삭제</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {invalid.map && <span className="help help-err"><Icon name="error" />{invalid.map}</span>}
        <div className="btn-row">
          <button className="btn" type="button" onClick={() => { touch(); setUserMap([...userMap, ['', '']]); }}>줄 추가</button>
          <span className="help">이 PC 계정: {user.account} (선택된 줄)</span>
        </div>
      </div>

      <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <div className="card-h" style={{ margin: 0 }}>검수 품목 <span className="help" style={{ fontWeight: 400 }}>빈 양식의 드롭다운 목록, 가져오기 때 '알 수 없는 품목' 점검에 쓰입니다</span></div>
        <div className="chips">
          {items.map((it) => (
            <span key={it} className="tagchip">{it}
              <button type="button" aria-label={`${it} 삭제`} onClick={() => { touch(); setItems(items.filter((x) => x !== it)); }}>×</button>
            </span>
          ))}
        </div>
        <div style={{ display: 'flex', gap: 8, maxWidth: 420 }}>
          <input className={`input${dupItem ? ' invalid' : ''}`} aria-label="추가할 검수 품목" placeholder="예: 가공품" value={newItem}
            onChange={(e) => setNewItem(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && addItem()} />
          <button className="btn" type="button" onClick={addItem} disabled={!trimmed || !!dupItem}>추가</button>
        </div>
        {(dupItem || invalid.items) && <span className="help help-err"><Icon name="error" />{dupItem || invalid.items}</span>}
        <span className="help">목록을 바꾼 뒤에는 [빈 양식 내려받기]로 새 양식을 만들어 업체에 보내세요.</span>
      </div>

      <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <div className="card-h" style={{ margin: 0 }}>잠금·백업</div>
        <div className="form-row">
          <div className="field">
            <label htmlFor="lk">'편집 중' 표시 유효 시간 (분)</label>
            <input id="lk" className={`input${invalid.lock ? ' invalid' : ''}`} value={lockMin} inputMode="numeric"
              onChange={(e) => { touch(); setLockMin(e.target.value); }} />
            {invalid.lock ? <span className="help help-err"><Icon name="error" />{invalid.lock}</span>
              : <span className="help">이 시간 넘게 갱신이 없는 '편집 중' 표시는 프로그램이 꺼진 것으로 보고 무시합니다.</span>}
          </div>
          <div className="field">
            <label htmlFor="bk">백업 보관 개수</label>
            <input id="bk" className={`input${invalid.keep ? ' invalid' : ''}`} value={keep} inputMode="numeric"
              onChange={(e) => { touch(); setKeep(e.target.value); }} />
            {invalid.keep ? <span className="help help-err"><Icon name="error" />{invalid.keep}</span>
              : <span className="help">저장할 때마다 이전 상태를 백업/ 에 남기고, 개수를 넘으면 오래된 것부터 지웁니다.</span>}
          </div>
        </div>
      </div>

      <div className="btn-row">
        <span className="sp" />
        <button className="btn" type="button" disabled={!dirty} onClick={cancel}>되돌리기</button>
        <button className="btn btn-primary" type="button" disabled={!dirty || hasInvalid} onClick={save}>저장</button>
      </div>

      <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <div className="card-h" style={{ margin: 0 }}>백업에서 복원 <span className="help" style={{ fontWeight: 400 }}>최근 {backups.length}개</span></div>
        {backups.length === 0 ? (
          <span className="help">아직 백업이 없습니다. 처음 저장할 때부터 생깁니다.</span>
        ) : (
          <div className="tbl-wrap" style={{ maxHeight: 320, overflow: 'auto' }}>
            <table className="tbl">
              <thead><tr><th>백업 시각</th><th>저장한 사람</th><th>폴더 이름</th><th /></tr></thead>
              <tbody>
                {backups.map((b) => (
                  <tr key={b.id}>
                    <td>{fmtDateTime(b.createdAt)}</td>
                    <td>{b.createdBy}</td>
                    <td className="muted">{b.id}</td>
                    <td><button className="btn btn-sm" type="button" onClick={() => setRestoreTarget(b)}>이 시점으로 복원</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {restoreTarget && (
        <Dialog
          title="백업에서 복원"
          footer={<>
            <button className="btn" type="button" onClick={() => setRestoreTarget(null)}>닫기</button>
            <button className="btn btn-primary" type="button" onClick={doRestore}>복원</button>
          </>}
        >
          <Alert kind="warn">
            모든 데이터(검수 건·업체·검수자·설정·메일 기록)를 <b>{fmtDateTime(restoreTarget.createdAt)}</b> 시점으로 되돌립니다.
            그 뒤에 저장된 내용은 사라집니다.
          </Alert>
          <span className="help">복원 직전 상태도 백업해 두므로, 잘못 복원했으면 다시 되돌릴 수 있습니다. 다른 PC는 새로 불러와야 바뀐 내용이 보입니다.</span>
        </Dialog>
      )}
    </>
  );
}
