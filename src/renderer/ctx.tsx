import { createContext, useContext, useEffect } from 'react';
import type { AppErrorInfo, CurrentUser, DataFiles, Inspection, InspectionStatus, MailKind } from '../shared/types';
import { api, type LockStatus } from './api';

export type Route =
  | { page: 'dashboard' | 'import' | 'settings' }
  | { page: 'list'; preset?: ListPreset }
  | { page: 'calendar'; date?: string }
  | { page: 'detail'; id: string; from?: Route }
  | { page: 'mail'; id: string; kind?: MailKind; from?: Route }
  | { page: 'vendors'; vendor?: string }
  | { page: 'inspectors' };

export interface ListPreset {
  q?: string;
  status?: InspectionStatus | '';
  vendor?: string;
  inspectorId?: string;
  /** 특정 날짜 (달력에서 빈 날짜를 눌렀을 때) */
  day?: string;
  unassigned?: boolean;
  /** 결과: pass | fail | missing */
  result?: string;
  /** 기간 */
  period?: 'today' | 'week' | 'month' | 'nextMonth';
}

export interface AppCtx {
  data: DataFiles;
  user: CurrentUser;
  today: string;
  go: (r: Route) => void;
  /**
   * 저장 공통 처리. 성공하면 null. 충돌이면 안내창을 띄우고 null 이 아닌 'handled'.
   * 그 밖의 오류는 오류 정보를 돌려줍니다.
   */
  save: (changes: Partial<DataFiles>, okMessage?: string) => Promise<AppErrorInfo | 'handled' | null>;
  saveInspections: (items: Inspection[], okMessage?: string) => Promise<AppErrorInfo | 'handled' | null>;
  inspectorName: (id: string | null | undefined) => string;
  setLock: (l: LockStatus) => void;
  /** 데이터 다시 읽기 */
  reload: () => Promise<void>;
}

export const Ctx = createContext<AppCtx | null>(null);
export const useApp = () => useContext(Ctx)!;

/** 팝업·편집 화면이 열려 있는 동안 '편집 중' 표시 */
export function useEditLock(active: boolean) {
  const { setLock } = useApp();
  useEffect(() => {
    if (!active) return;
    api.lockAcquire().then((r) => r.ok && setLock(r.value));
    return () => {
      api.lockRelease().then(() => api.lockStatus()).then((r) => r.ok && setLock(r.value));
    };
  }, [active]); // eslint-disable-line react-hooks/exhaustive-deps
}
