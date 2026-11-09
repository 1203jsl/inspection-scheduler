import type { DataFileKey, DataFiles, ExcelTemplateDef, FileMeta, Settings } from '../../shared/types';
import { DEFAULT_CONFLICT_RULES } from '../../shared/conflicts';
import { DEFAULT_MAIL_SETTINGS, DEFAULT_TEMPLATES } from '../../shared/mail';
import { AppError } from './errors';

/** 이 프로그램이 이해하는 최신 구조 버전 (파일별) */
export const CURRENT_SCHEMA: Record<DataFileKey, number> = {
  inspections: 2,
  vendors: 1,
  inspectors: 1,
  settings: 4,
  mail_log: 1,
};

/** 엑셀 양식 v1 (요청서 2-5의 열 순서) */
export const TEMPLATE_V1: ExcelTemplateDef = {
  version: 'v1',
  sheetName: '검수신청',
  columns: [
    { key: 'id', header: '신청번호', width: 15 },
    { key: 'vendorName', header: '업체명', aliases: ['업체', '회사명'], required: true, width: 14 },
    { key: 'category', header: '검수 품목', aliases: ['검수품목', '품목'], required: true, width: 12 },
    { key: 'itemName', header: '품명', required: true, width: 20 },
    { key: 'spec', header: '품번/규격', aliases: ['품번', '규격', '품번규격'], width: 20 },
    { key: 'quantity', header: '수량', required: true, width: 8 },
    { key: 'requestedDate', header: '검수 희망일', aliases: ['희망일', '검수희망일'], required: true, width: 13 },
    { key: 'requestedTime', header: '희망 시간', aliases: ['희망시간', '시간'], width: 10 },
    { key: 'location', header: '검수 장소', aliases: ['검수장소', '장소'], width: 14 },
    { key: 'contactName', header: '업체 담당자', aliases: ['담당자', '업체담당자'], required: true, width: 11 },
    { key: 'contactPhone', header: '연락처', aliases: ['전화번호', '휴대폰'], width: 15 },
    { key: 'contactEmail', header: '이메일', aliases: ['메일', 'e-mail', 'email'], required: true, width: 24 },
    { key: 'note', header: '비고', width: 24 },
  ],
};

export const DEFAULT_INSPECTION_ITEMS = ['TANK', 'BODY', 'VMB_TURN-KEY', 'UNITTURN-KEY', "Part's"];

/**
 * 구조 변환표. MIGRATIONS[파일][옛 버전] = 옛 버전 → 다음 버전 변환.
 */
type Migration = (data: any) => any;
export const MIGRATIONS: Partial<Record<DataFileKey, Record<number, Migration>>> = {
  inspections: {
    // v1 → v2: 현재 일정(schedule)·확정 정보 추가. 처음에는 업체 희망값으로 채움
    1: (d) => ({
      ...d,
      schemaVersion: 2,
      items: (d.items ?? []).map((i: any) => ({
        ...i,
        schedule: i.schedule ?? { date: i.requestedDate ?? null, time: i.requestedTime ?? null, location: i.location ?? '', inspectorId: null },
        confirmedAt: i.confirmedAt ?? null,
        confirmedBy: i.confirmedBy ?? null,
        negotiations: i.negotiations ?? [],
        result: i.result ?? null,
      })),
    }),
  },
  settings: {
    // v1 → v2: 엑셀 양식 정의와 검수 품목 목록 추가
    1: (d) => ({
      ...d,
      schemaVersion: 2,
      inspectionItems: d.inspectionItems ?? [...DEFAULT_INSPECTION_ITEMS],
      excelTemplates: d.excelTemplates ?? [TEMPLATE_V1],
      currentTemplate: d.currentTemplate ?? 'v1',
    }),
    // v2 → v3: 충돌 조건 추가
    2: (d) => ({ ...d, schemaVersion: 3, conflictRules: d.conflictRules ?? structuredClone(DEFAULT_CONFLICT_RULES) }),
    // v3 → v4: 메일 템플릿·발송 설정 추가
    3: (d) => ({
      ...d,
      schemaVersion: 4,
      mailTemplates: { ...structuredClone(DEFAULT_TEMPLATES), ...(d.mailTemplates ?? {}) },
      mail: d.mail ?? structuredClone(DEFAULT_MAIL_SETTINGS),
    }),
  },
};

function baseMeta(user: string, name: string, now: string, key: DataFileKey): FileMeta {
  return { schemaVersion: CURRENT_SCHEMA[key], revision: 0, updatedBy: user, updatedByName: name, updatedAt: now };
}

export function defaultSettings(user: string, name: string, now: string): Settings {
  return {
    ...baseMeta(user, name, now, 'settings'),
    userMap: name ? { [user]: name } : {},
    lockStaleMinutes: 30,
    backupKeep: 50,
    inspectionItems: [...DEFAULT_INSPECTION_ITEMS],
    excelTemplates: [TEMPLATE_V1],
    currentTemplate: 'v1',
    conflictRules: structuredClone(DEFAULT_CONFLICT_RULES),
    mailTemplates: structuredClone(DEFAULT_TEMPLATES),
    mail: structuredClone(DEFAULT_MAIL_SETTINGS),
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

/** 양식 버전 이름으로 열 정의를 찾습니다 (없으면 현재 양식). */
export function templateDef(settings: Settings, version?: string): ExcelTemplateDef {
  const list = settings.excelTemplates?.length ? settings.excelTemplates : [TEMPLATE_V1];
  return list.find((t) => t.version === version)
    ?? list.find((t) => t.version === settings.currentTemplate)
    ?? list[0];
}
