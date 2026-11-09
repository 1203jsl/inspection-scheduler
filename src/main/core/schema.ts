import type { DataFileKey, DataFiles, FileMeta, Settings } from '../../shared/types';
import { AppError } from './errors';

/** 이 프로그램이 이해하는 최신 구조 버전 (파일별) */
export const CURRENT_SCHEMA: Record<DataFileKey, number> = {
  inspections: 1,
  vendors: 1,
  inspectors: 1,
  settings: 1,
  mail_log: 1,
};

/**
 * 구조 변환표. 예: 2단계에서 inspections 를 2로 올리면
 * MIGRATIONS.inspections[1] = (d) => ({ ...d, schemaVersion: 2, ... }) 를 추가합니다.
 */
type Migration = (data: any) => any;
export const MIGRATIONS: Partial<Record<DataFileKey, Record<number, Migration>>> = {};

function baseMeta(user: string, name: string, now: string, key: DataFileKey): FileMeta {
  return { schemaVersion: CURRENT_SCHEMA[key], revision: 0, updatedBy: user, updatedByName: name, updatedAt: now };
}

export function defaultSettings(user: string, name: string, now: string): Settings {
  return {
    ...baseMeta(user, name, now, 'settings'),
    userMap: name ? { [user]: name } : {},
    lockStaleMinutes: 30,
    backupKeep: 50,
  };
}

export function defaultFile<K extends DataFileKey>(key: K, user: string, name: string, now: string): DataFiles[K] {
  if (key === 'settings') return defaultSettings(user, name, now) as DataFiles[K];
  return { ...baseMeta(user, name, now, key), items: [] } as unknown as DataFiles[K];
}

/** 읽은 파일을 최신 구조로 바꿉니다. 더 새 구조면 오류. */
export function migrate<K extends DataFileKey>(key: K, raw: any): DataFiles[K] {
  if (typeof raw !== 'object' || raw === null || typeof raw.schemaVersion !== 'number') {
    throw new AppError('FILE_CORRUPT', `${key}.json 에 schemaVersion 이 없습니다. 설정 > 백업에서 이전 시점으로 복원하세요.`, { file: key });
  }
  let data = raw;
  const target = CURRENT_SCHEMA[key];
  if (data.schemaVersion > target) {
    throw new AppError(
      'NEWER_SCHEMA',
      `${key}.json 이 더 새 버전의 프로그램으로 저장되었습니다. 프로그램을 최신 버전으로 업데이트한 뒤 사용하세요.`,
      { file: key },
    );
  }
  while (data.schemaVersion < target) {
    const step = MIGRATIONS[key]?.[data.schemaVersion];
    if (!step) throw new AppError('FILE_CORRUPT', `${key}.json 구조 ${data.schemaVersion} 을 변환할 수 없습니다.`, { file: key });
    data = step(data);
  }
  if (typeof data.revision !== 'number') data.revision = 0;
  return data as DataFiles[K];
}
