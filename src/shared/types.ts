// 프로세스(메인/화면) 공통 타입

/** data/ 폴더 안의 JSON 파일 종류 */
export type DataFileKey = 'inspections' | 'vendors' | 'inspectors' | 'settings' | 'mail_log';

export const DATA_FILE_KEYS: DataFileKey[] = ['inspections', 'vendors', 'inspectors', 'settings', 'mail_log'];

/** 모든 데이터 파일이 공통으로 갖는 머리 정보 */
export interface FileMeta {
  /** 구조 버전. 구조가 바뀌면 올리고 migrate.ts에 변환을 추가합니다. */
  schemaVersion: number;
  /** 저장할 때마다 1씩 오르는 번호. 덮어쓰기 방지에 씁니다. */
  revision: number;
  /** 마지막으로 저장한 Windows 계정 */
  updatedBy: string;
  /** 마지막으로 저장한 사람의 표시 이름 */
  updatedByName: string;
  /** 마지막 저장 시각 (ISO 8601) */
  updatedAt: string;
}

export interface ListFile<T> extends FileMeta {
  items: T[];
}

export interface Settings extends FileMeta {
  /** Windows 계정 → 표시 이름 */
  userMap: Record<string, string>;
  /** 이 시간(분)이 지나도록 갱신되지 않은 .lock 은 무시 */
  lockStaleMinutes: number;
  /** 백업/ 에 남길 최대 개수 */
  backupKeep: number;
}

export type DataFiles = {
  inspections: ListFile<unknown>;
  vendors: ListFile<unknown>;
  inspectors: ListFile<unknown>;
  settings: Settings;
  mail_log: ListFile<unknown>;
};

export interface LockInfo {
  user: string;
  displayName: string;
  host: string;
  startedAt: string;
  heartbeatAt: string;
}

export interface BackupEntry {
  /** 백업 폴더 이름 */
  id: string;
  createdAt: string;
  createdBy: string;
  files: DataFileKey[];
}

export interface CurrentUser {
  account: string;
  displayName: string;
  host: string;
}

/** 화면에 보여줄 오류. code로 종류를 구분합니다. */
export type AppErrorCode =
  | 'CONFLICT'        // 다른 사람이 먼저 저장함
  | 'FOLDER_MISSING'  // 데이터 폴더에 접근 불가
  | 'FILE_BUSY'       // 파일이 잠겨 교체 실패
  | 'FILE_CORRUPT'    // JSON이 깨짐
  | 'NEWER_SCHEMA'    // 더 새 버전 프로그램이 저장한 데이터
  | 'NOT_INITIALIZED' // 데이터 폴더 구조가 없음
  | 'UNKNOWN';

export interface AppErrorInfo {
  code: AppErrorCode;
  message: string;
  detail?: {
    file?: DataFileKey;
    updatedBy?: string;
    updatedByName?: string;
    updatedAt?: string;
    path?: string;
  };
}

export type Result<T> = { ok: true; value: T } | { ok: false; error: AppErrorInfo };

export interface FolderCheck {
  path: string;
  exists: boolean;
  /** data/settings.json 이 있는 기존 데이터 폴더인지 */
  hasData: boolean;
  summary?: {
    inspections: number;
    vendors: number;
    inspectors: number;
    lastSavedBy?: string;
    lastSavedAt?: string;
  };
}

export interface AppState {
  dataDir: string | null;
  user: CurrentUser;
}
