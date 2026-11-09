import type {
  AppState, BackupEntry, DataFiles, FolderCheck, ImportPreview, ImportSummary, MailDraft, MailLogEntry, Result, SmtpAccountInfo,
} from '../shared/types';
import type { LockStatus } from '../main/core/lock';

export interface Api {
  getState(): Promise<Result<AppState>>;
  pickFolder(): Promise<Result<string | null>>;
  checkFolder(path: string): Promise<Result<FolderCheck>>;
  setupFolder(o: { path: string; mode: 'existing' | 'new'; displayName: string }): Promise<Result<AppState>>;
  loadAll(): Promise<Result<DataFiles>>;
  save(changes: Partial<DataFiles>): Promise<Result<Partial<DataFiles>>>;
  lockStatus(): Promise<Result<LockStatus>>;
  lockAcquire(): Promise<Result<LockStatus>>;
  lockRelease(): Promise<Result<void>>;
  listBackups(): Promise<Result<BackupEntry[]>>;
  restoreBackup(id: string): Promise<Result<boolean>>;
  saveTemplate(): Promise<Result<string | null>>;
  pickImportFiles(): Promise<Result<string[]>>;
  importPreview(paths: string[]): Promise<Result<ImportPreview>>;
  importCommit(keys: string[]): Promise<Result<ImportSummary>>;
  openDataFile(rel: string): Promise<Result<boolean>>;
  mailAccount(): Promise<Result<SmtpAccountInfo>>;
  /** password: 비우면(undefined) 기존 유지, '' 이면 삭제 */
  mailSaveAccount(o: { account: string; password?: string }): Promise<Result<SmtpAccountInfo>>;
  mailSend(d: MailDraft): Promise<Result<MailLogEntry>>;
  mailRecordManual(d: MailDraft): Promise<Result<MailLogEntry>>;
  mailTest(): Promise<Result<MailLogEntry>>;
  mailPickAttachments(): Promise<Result<string[]>>;
  mailOpenWebmail(url: string): Promise<Result<boolean>>;
  copyText(text: string): Promise<Result<boolean>>;
  pathForFile(f: File): string;
}

declare global {
  interface Window {
    api: Api;
  }
}

export const api = window.api;
export type { LockStatus };

const p = (n: number) => String(n).padStart(2, '0');
export const fmtTime = (iso?: string) => {
  if (!iso) return '';
  const d = new Date(iso);
  return `${p(d.getHours())}:${p(d.getMinutes())}`;
};
const WD = ['일', '월', '화', '수', '목', '금', '토'];
export const fmtDateTime = (iso?: string) => {
  if (!iso) return '';
  const d = new Date(iso);
  return `${p(d.getMonth() + 1)}-${p(d.getDate())}(${WD[d.getDay()]}) ${p(d.getHours())}:${p(d.getMinutes())}`;
};
