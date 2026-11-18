import type { AppErrorCode, AppErrorInfo, DataFileKey } from '../../shared/types';

export class AppError extends Error {
  constructor(
    public code: AppErrorCode,
    message: string,
    public detail: AppErrorInfo['detail'] = {},
  ) {
    super(message);
  }
  toInfo(): AppErrorInfo {
    return { code: this.code, message: this.message, detail: this.detail };
  }
}

/** 파일을 만들거나 고칠 권한이 없을 때 (회사 보안·문서중앙화 정책으로 막힌 경우 포함) */
export function noPermission(file: string): AppError {
  return new AppError(
    'NO_PERMISSION',
    `이 폴더에 파일을 저장할 권한이 없습니다: ${file}\n` +
      '회사 보안 프로그램(문서중앙화·DRM)이 허용 목록에 없는 프로그램의 저장을 막았을 수 있습니다. ' +
      '메모장 등 다른 프로그램으로는 이 폴더에 저장되는데 이 프로그램만 안 되면, ' +
      'IT 담당자에게 \'검수일정 관리(InspectionScheduler.exe)\'의 저장 허용을 요청하세요. ' +
      'Windows 보안의 \'제어된 폴더 액세스\' 알림이 떴다면 그곳에서 허용해도 됩니다.',
    { path: file },
  );
}

const PERM_CODES = ['EPERM', 'EACCES'];
export const isPermError = (e: unknown) => PERM_CODES.includes((e as NodeJS.ErrnoException)?.code ?? '');

export function toErrorInfo(e: unknown): AppErrorInfo {
  if (e instanceof AppError) return e.toInfo();
  if (isPermError(e) && (e as NodeJS.ErrnoException).path) return noPermission((e as NodeJS.ErrnoException).path!).toInfo();
  const msg = e instanceof Error ? e.message : String(e);
  return { code: 'UNKNOWN', message: `예상하지 못한 오류가 발생했습니다: ${msg}` };
}

/** "11-09 10:32" 처럼 사람이 읽기 쉬운 시각 */
export function hhmm(iso: string): string {
  const d = new Date(iso);
  const p = (n: number) => String(n).padStart(2, '0');
  return `${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

export function conflictError(file: DataFileKey, by: string, byName: string, at: string): AppError {
  const who = byName || by || '다른 사용자';
  return new AppError(
    'CONFLICT',
    `${who}이(가) ${hhmm(at)}에 수정했습니다. 새로 불러온 뒤 다시 작업하세요.`,
    { file, updatedBy: by, updatedByName: byName, updatedAt: at },
  );
}

export function folderMissing(path: string): AppError {
  return new AppError(
    'FOLDER_MISSING',
    `데이터 폴더에 접근할 수 없습니다 (${path}). 클라우드독이 실행 중이고 동기화 폴더가 연결되어 있는지 확인하세요. 데이터는 바뀌지 않았습니다.`,
    { path },
  );
}
