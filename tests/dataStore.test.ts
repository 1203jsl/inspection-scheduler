import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { DataStore } from '../src/main/core/dataStore';
import { AppError } from '../src/main/core/errors';

const KIM = { account: 'kim.mj', displayName: '김민준' };
const LEE = { account: 'lee.sy', displayName: '이서연' };

let dir: string;
beforeEach(async () => {
  dir = await fs.mkdtemp(path.join(os.tmpdir(), 'insp-'));
});
afterEach(async () => {
  await fs.rm(dir, { recursive: true, force: true });
});

async function code(p: Promise<unknown>) {
  try {
    await p;
    return 'OK';
  } catch (e) {
    return e instanceof AppError ? e.code : String(e);
  }
}

describe('데이터 폴더 만들기·확인', () => {
  it('새 폴더에 data/원본/첨부/백업 과 기본 JSON 을 만든다', async () => {
    expect((await DataStore.checkFolder(dir)).hasData).toBe(false);
    await DataStore.initFolder(dir, KIM);
    for (const d of ['data', '원본', '첨부', '백업']) {
      expect((await fs.stat(path.join(dir, d))).isDirectory()).toBe(true);
    }
    const s = JSON.parse(await fs.readFile(path.join(dir, 'data/settings.json'), 'utf8'));
    expect(s.schemaVersion).toBe(2);
    expect(s.revision).toBe(0);
    expect(s.userMap).toEqual({ 'kim.mj': '김민준' });
    const chk = await DataStore.checkFolder(dir);
    expect(chk.hasData).toBe(true);
    expect(chk.summary?.inspections).toBe(0);
  });

  it('없는 폴더는 FOLDER_MISSING', async () => {
    const s = new DataStore(path.join(dir, 'nope'), KIM);
    expect(await code(s.load('settings'))).toBe('FOLDER_MISSING');
  });

  it('데이터가 없는 폴더는 NOT_INITIALIZED', async () => {
    const s = new DataStore(dir, KIM);
    expect(await code(s.load('settings'))).toBe('NOT_INITIALIZED');
  });
});

describe('덮어쓰기 방지', () => {
  beforeEach(async () => {
    await DataStore.initFolder(dir, KIM);
  });

  it('저장할 때마다 revision 이 1씩 오르고 저장자가 기록된다', async () => {
    const s = new DataStore(dir, KIM);
    const v = await s.load('vendors');
    const saved = await s.save({ vendors: { ...v, items: [{ name: '한솔정밀' }] as any } });
    expect(saved.vendors!.revision).toBe(1);
    expect(saved.vendors!.updatedByName).toBe('김민준');
    const again = await new DataStore(dir, LEE).load('vendors');
    expect(again.revision).toBe(1);
    expect(again.items).toHaveLength(1);
  });

  it('다른 사람이 먼저 저장했으면 저장하지 않고 누가 언제 바꿨는지 알려준다', async () => {
    const kim = new DataStore(dir, KIM);
    const lee = new DataStore(dir, LEE, () => new Date(2026, 10, 9, 10, 32));
    const kv = await kim.load('vendors');
    const lv = await lee.load('vendors');

    await lee.save({ vendors: { ...lv, items: [{ name: '대양산업' }] as any } });

    let err: AppError | null = null;
    try {
      await kim.save({ vendors: { ...kv, items: [{ name: '한솔정밀' }] as any } });
    } catch (e) {
      err = e as AppError;
    }
    expect(err?.code).toBe('CONFLICT');
    expect(err?.message).toBe('이서연이(가) 11-09 10:32에 수정했습니다. 새로 불러온 뒤 다시 작업하세요.');
    // 이서연의 저장 내용이 그대로 남아 있어야 함
    const disk = JSON.parse(await fs.readFile(path.join(dir, 'data/vendors.json'), 'utf8'));
    expect(disk.items).toEqual([{ name: '대양산업' }]);
  });

  it('새로 불러온 뒤에는 저장된다', async () => {
    const kim = new DataStore(dir, KIM);
    const lee = new DataStore(dir, LEE);
    await kim.load('vendors');
    const lv = await lee.load('vendors');
    await lee.save({ vendors: { ...lv, items: [{ name: 'A' }] as any } });
    const fresh = await kim.load('vendors');
    const saved = await kim.save({ vendors: { ...fresh, items: [...fresh.items, { name: 'B' }] as any } });
    expect(saved.vendors!.revision).toBe(2);
  });

  it('여러 파일 저장 중 하나라도 충돌이면 아무것도 저장하지 않는다', async () => {
    const kim = new DataStore(dir, KIM);
    const lee = new DataStore(dir, LEE);
    const ki = await kim.load('inspections');
    const kv = await kim.load('vendors');
    const lv = await lee.load('vendors');
    await lee.save({ vendors: { ...lv, items: [{ name: 'X' }] as any } });
    expect(await code(kim.save({ inspections: { ...ki, items: [{ id: 1 }] as any }, vendors: kv }))).toBe('CONFLICT');
    const disk = JSON.parse(await fs.readFile(path.join(dir, 'data/inspections.json'), 'utf8'));
    expect(disk.items).toEqual([]);
    expect(disk.revision).toBe(0);
  });

  it('불러오지 않은 파일은 저장할 수 없다', async () => {
    const kim = new DataStore(dir, KIM);
    expect(await code(kim.save({ vendors: { items: [] as any } as any }))).toBe('UNKNOWN');
  });

  it('저장 중 임시 파일이 남지 않는다', async () => {
    const kim = new DataStore(dir, KIM);
    const v = await kim.load('vendors');
    await kim.save({ vendors: v });
    const files = await fs.readdir(path.join(dir, 'data'));
    expect(files.filter((f) => f.endsWith('.tmp'))).toEqual([]);
  });
});

describe('손상·버전', () => {
  beforeEach(async () => {
    await DataStore.initFolder(dir, KIM);
  });

  it('깨진 JSON 은 FILE_CORRUPT, 파일은 그대로 둔다', async () => {
    const f = path.join(dir, 'data/vendors.json');
    await fs.writeFile(f, '{"schemaVersion":1, "revi');
    const s = new DataStore(dir, KIM);
    expect(await code(s.load('vendors'))).toBe('FILE_CORRUPT');
    expect(await fs.readFile(f, 'utf8')).toBe('{"schemaVersion":1, "revi');
  });

  it('더 새 구조 버전이면 NEWER_SCHEMA', async () => {
    const f = path.join(dir, 'data/vendors.json');
    await fs.writeFile(f, JSON.stringify({ schemaVersion: 99, revision: 3, items: [] as any }));
    expect(await code(new DataStore(dir, KIM).load('vendors'))).toBe('NEWER_SCHEMA');
  });
});

describe('백업·복원', () => {
  beforeEach(async () => {
    await DataStore.initFolder(dir, KIM);
  });

  it('저장할 때마다 이전 버전을 백업하고, 개수를 넘으면 오래된 것부터 지운다', async () => {
    let t = new Date(2026, 10, 9, 9, 0, 0).getTime();
    const s = new DataStore(dir, KIM, () => new Date((t += 1000)));
    const st = await s.load('settings');
    await s.save({ settings: { ...st, backupKeep: 3 } });
    for (let i = 0; i < 5; i++) {
      const v = await s.load('vendors');
      await s.save({ vendors: { ...v, items: [{ n: i }] as any } });
    }
    const list = await s.listBackups();
    expect(list).toHaveLength(3);
    expect(list[0].createdBy).toBe('김민준');
    // 가장 최근 백업에는 저장 직전 상태(n:3)가 있어야 함
    const last = JSON.parse(await fs.readFile(path.join(dir, '백업', list[0].id, 'vendors.json'), 'utf8'));
    expect(last.items).toEqual([{ n: 3 }]);
  });

  it('복원하면 백업 내용으로 돌아가고 revision 은 올라가 다른 PC가 변경을 알아챈다', async () => {
    let t = new Date(2026, 10, 9, 9, 0, 0).getTime();
    const kim = new DataStore(dir, KIM, () => new Date((t += 1000)));
    let v = await kim.load('vendors');
    await kim.save({ vendors: { ...v, items: [{ name: '첫번째' }] as any } });
    v = await kim.load('vendors');
    await kim.save({ vendors: { ...v, items: [{ name: '잘못 입력' }] as any } });

    const lee = new DataStore(dir, LEE);
    const lv = await lee.load('vendors'); // revision 2

    const target = (await kim.listBackups())[0]; // 두 번째 저장 직전 = '첫번째'
    await kim.restoreBackup(target.id);
    const now = await kim.load('vendors');
    expect(now.items).toEqual([{ name: '첫번째' }]);
    expect(now.revision).toBe(3);

    expect(await code(lee.save({ vendors: lv }))).toBe('CONFLICT');
  });

  it('손상된 파일도 복원으로 되살린다', async () => {
    const kim = new DataStore(dir, KIM);
    const v = await kim.load('vendors');
    await kim.save({ vendors: { ...v, items: [{ ok: true }] as any } });
    const id = (await kim.listBackups())[0].id;
    await fs.writeFile(path.join(dir, 'data/vendors.json'), '{broken');
    await kim.restoreBackup(id);
    expect((await kim.load('vendors')).items).toEqual([]);
  });

  it('이상한 백업 이름은 거부한다', async () => {
    const kim = new DataStore(dir, KIM);
    expect(await code(kim.restoreBackup('../data'))).toBe('UNKNOWN');
  });
});
