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

// ───────── 엑셀 양식 ─────────

/** 업체 신청 항목 (엑셀 열과 1:1) */
export type FieldKey =
  | 'id' | 'vendorName' | 'category' | 'itemName' | 'spec' | 'quantity'
  | 'requestedDate' | 'requestedTime' | 'location' | 'contactName' | 'contactPhone' | 'contactEmail' | 'note';

export interface ColumnDef {
  key: FieldKey;
  /** 양식 머리글 (이 이름으로 열을 찾음) */
  header: string;
  /** 업체가 머리글을 조금 바꿔 써도 찾을 수 있게 하는 다른 이름 */
  aliases?: string[];
  required?: boolean;
  /** 엑셀 열 너비 */
  width?: number;
}

/** 양식 버전별 열 정의. 양식이 바뀌면 settings.json 에 새 버전을 추가합니다. */
export interface ExcelTemplateDef {
  version: string;
  sheetName: string;
  columns: ColumnDef[];
}

// ───────── 검수 건 ─────────

export type InspectionStatus =
  | 'received'     // 신청접수
  | 'registered'   // 등록(신청번호 부여)
  | 'negotiating'  // 일정협의중
  | 'confirmed'    // 확정
  | 'changed'      // 변경(재확정 필요)
  | 'done'         // 검수완료
  | 'canceled';    // 취소

export const STATUS_LABEL: Record<InspectionStatus, string> = {
  received: '신청접수',
  registered: '등록',
  negotiating: '일정협의중',
  confirmed: '확정',
  changed: '변경됨',
  done: '검수완료',
  canceled: '취소',
};

/** 업체가 신청한 값 */
export interface RequestFields {
  vendorName: string;
  category: string;
  itemName: string;
  spec: string;
  quantity: number | null;
  /** YYYY-MM-DD */
  requestedDate: string | null;
  /** HH:MM */
  requestedTime: string | null;
  location: string;
  contactName: string;
  contactPhone: string;
  contactEmail: string;
  note: string;
}

export const REQUEST_FIELDS: (keyof RequestFields)[] = [
  'vendorName', 'category', 'itemName', 'spec', 'quantity', 'requestedDate', 'requestedTime',
  'location', 'contactName', 'contactPhone', 'contactEmail', 'note',
];

export interface HistoryEntry {
  at: string;
  by: string;
  byName: string;
  /** created | status | field | import | ... */
  action: string;
  field?: string;
  from?: unknown;
  to?: unknown;
  reason?: string;
}

export interface SourceRef {
  /** 데이터 폴더 기준 상대 경로 (원본/2026-11/...) */
  file: string;
  row: number;
  importedAt: string;
  importedBy: string;
}

export interface Inspection extends RequestFields {
  /** 신청번호 IN-2026-0001 */
  id: string;
  status: InspectionStatus;
  /** 처음 신청한 값 (상세 화면의 '업체 신청 원본' 비교용) */
  original: RequestFields;
  sources: SourceRef[];
  negotiations: unknown[];
  result: unknown | null;
  attachments: unknown[];
  history: HistoryEntry[];
  createdAt: string;
  createdBy: string;
}

export interface Vendor {
  name: string;
  contactName: string;
  contactPhone: string;
  contactEmail: string;
  updatedAt: string;
}

// ───────── 가져오기 미리보기 ─────────

export interface CellIssue {
  level: 'warn' | 'error';
  message: string;
}

export interface PreviewRow {
  /** 파일번호:행번호 */
  key: string;
  fileIndex: number;
  fileName: string;
  rowNumber: number;
  /** 판정: 신규 / 수정 / 오류 */
  kind: 'new' | 'update' | 'error';
  /** 오류 행을 포함했을 때 처리 방식 */
  target: 'new' | 'update';
  /** 엑셀에 적힌 그대로(표시용) */
  display: Record<FieldKey, string>;
  parsed: RequestFields & { id: string };
  /** 칸별 문제 ('row' 는 행 전체) */
  issues: Partial<Record<FieldKey | 'row', CellIssue[]>>;
  /** 수정 행: 바뀐 칸만 기존값 → 새값 */
  diff?: Partial<Record<keyof RequestFields, { from: string; to: string }>>;
  defaultInclude: boolean;
}

export interface PreviewFile {
  path: string;
  name: string;
  templateVersion?: string;
  rowCount: number;
  error?: AppErrorInfo;
  warnings?: string[];
}

export interface ImportPreview {
  files: PreviewFile[];
  rows: PreviewRow[];
}

export interface ImportSummary {
  created: { id: string; vendorName: string; itemName: string; contactEmail: string }[];
  updated: string[];
  excluded: number;
  originals: string[];
}

export interface Settings extends FileMeta {
  /** Windows 계정 → 표시 이름 */
  userMap: Record<string, string>;
  /** 이 시간(분)이 지나도록 갱신되지 않은 .lock 은 무시 */
  lockStaleMinutes: number;
  /** 백업/ 에 남길 최대 개수 */
  backupKeep: number;
  /** 검수 품목 목록 (양식 드롭다운, 가져오기 점검) */
  inspectionItems: string[];
  /** 양식 버전별 열 정의 */
  excelTemplates: ExcelTemplateDef[];
  /** 빈 양식을 만들 때 쓰는 버전 */
  currentTemplate: string;
}

export type DataFiles = {
  inspections: ListFile<Inspection>;
  vendors: ListFile<Vendor>;
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
