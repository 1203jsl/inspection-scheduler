import { promises as fs } from 'node:fs';
import path from 'node:path';
import type { BackupEntry, DataFileKey, DataFiles, FileMeta, FolderCheck } from '../../shared/types';
import { DATA_FILE_KEYS } from '../../shared/types';
import { AppError, conflictError, folderMissing } from './errors';
import { exists, isDir, readText, writeFileAtomic } from './fsSafe';
import { defaultFile, migrate } from './schema';

/** 데이터 폴더 안의 하위 폴더 이름 */
export const DIRS = { data: 'data', original: '원본', attach: '첨부', backup: '백업' } as const;
export const LOCK_FILE = '.lock';
/** 저장 잠금 파일 (data/ 안, 저장하는 동안만 존재) */
export const WRITE_LOCK = '.writing';
const WRITE_LOCK_WAIT_MS = 5_000;
const WRITE_LOCK_STALE_MS = 15_000;

export interface StoreUser {
  account: string;
  displayName: string;
}

const pad = (n: number, w = 2) => String(n).padStart(w, '0');
function stamp(d: Date): string {
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}-${pad(d.getMilliseconds(), 3)}`;
}
/** 폴더 이름에 쓸 수 없는 글자 제거 */
const safeName = (s: string) => s.replace(/[\\/:*?"<>|\s]+/g, '_').slice(0, 40) || 'user';

/**
 * 동기화 폴더 안의 JSON 데이터를 읽고 씁니다.
 * - 읽을 때 revision 을 기억했다가, 저장 직전 디스크의 revision 과 비교합니다.
 * - 다르면 저장하지 않고 CONFLICT 오류를 냅니다 (덮어쓰기 방지).
 * - 저장 전에는 현재 data/ 를 백업/ 에 복사합니다.
 */
export class DataStore {
  private loadedRevision = new Map<DataFileKey, number>();

  constructor(
    public readonly root: string,
    private user: StoreUser,
    private now: () => Date = () => new Date(),
  ) {}

  setUser(user: StoreUser) {
    this.user = user;
  }

  fileOf(key: DataFileKey) {
    return path.join(this.root, DIRS.data, `${key}.json`);
  }

  // ───────── 폴더 확인·생성 ─────────

  static async checkFolder(dir: string): Promise<FolderCheck> {
    const res: FolderCheck = { path: dir, exists: await isDir(dir), hasData: false };
    if (!res.exists) return res;
    const settingsFile = path.join(dir, DIRS.data, 'settings.json');
    res.hasData = await exists(settingsFile);
    if (res.hasData) {
      const count = async (k: DataFileKey) => {
        try {
          const j = JSON.parse(await readText(path.join(dir, DIRS.data, `${k}.json`)));
          return Array.isArray(j.items) ? j.items.length : 0;
        } catch {
          return 0;
        }
      };
      let lastSavedBy: string | undefined;
      let lastSavedAt: string | undefined;
      for (const k of DATA_FILE_KEYS) {
        try {
          const j = JSON.parse(await readText(path.join(dir, DIRS.data, `${k}.json`))) as FileMeta;
          if (!lastSavedAt || j.updatedAt > lastSavedAt) {
            lastSavedAt = j.updatedAt;
            lastSavedBy = j.updatedByName || j.updatedBy;
          }
        } catch {
          /* 요약용이므로 무시 */
        }
      }
      res.summary = {
        inspections: await count('inspections'),
        vendors: await count('vendors'),
        inspectors: await count('inspectors'),
        lastSavedBy,
        lastSavedAt,
      };
    }
    return res;
  }

  /** 새 데이터 폴더 구조를 만듭니다. 이미 있는 파일은 건드리지 않습니다. */
  static async initFolder(dir: string, user: StoreUser, now = new Date()): Promise<void> {
    if (!(await isDir(dir))) throw folderMissing(dir);
    for (const d of Object.values(DIRS)) await fs.mkdir(path.join(dir, d), { recursive: true });
    const iso = now.toISOString();
    for (const k of DATA_FILE_KEYS) {
      const f = path.join(dir, DIRS.data, `${k}.json`);
      if (await exists(f)) continue;
      await writeFileAtomic(f, JSON.stringify(defaultFile(k, user.account, user.displayName, iso), null, 2));
    }
  }

  private async ensureRoot() {
    if (!(await isDir(this.root))) throw folderMissing(this.root);
    if (!(await exists(path.join(this.root, DIRS.data, 'settings.json')))) {
      throw new AppError('NOT_INITIALIZED', `이 폴더에는 검수일정 데이터가 없습니다 (${this.root}). 설정에서 데이터 폴더를 다시 선택하세요.`, { path: this.root });
    }
  }

  // ───────── 읽기 ─────────

  /** 디스크에서 그대로 읽습니다 (기억한 revision 은 바꾸지 않음). */
  private async readRaw<K extends DataFileKey>(key: K): Promise<DataFiles[K]> {
    const file = this.fileOf(key);
    let text: string;
    try {
      text = await readText(file);
    } catch (e) {
      if ((e as NodeJS.ErrnoException).code === 'ENOENT') {
        if (!(await isDir(this.root))) throw folderMissing(this.root);
        // 파일만 없으면 빈 파일로 간주 (revision 0)
        const iso = this.now().toISOString();
        return defaultFile(key, this.user.account, this.user.displayName, iso);
      }
      throw e;
    }
    let raw: unknown;
    try {
      raw = JSON.parse(text);
    } catch {
      throw new AppError(
        'FILE_CORRUPT',
        `data/${key}.json 파일이 손상되어 읽을 수 없습니다. 동기화가 끝났는지 확인한 뒤 다시 불러오고, 계속되면 설정 > 백업에서 이전 시점으로 복원하세요.`,
        { file: key },
      );
    }
    return migrate(key, raw);
  }

  async load<K extends DataFileKey>(key: K): Promise<DataFiles[K]> {
    await this.ensureRoot();
    const data = await this.readRaw(key);
    this.loadedRevision.set(key, data.revision);
    return data;
  }

  async loadAll(): Promise<DataFiles> {
    await this.ensureRoot();
    const out = {} as DataFiles;
    for (const k of DATA_FILE_KEYS) (out as any)[k] = await this.load(k);
    return out;
  }

  getLoadedRevision(key: DataFileKey) {
    return this.loadedRevision.get(key);
  }

  // ───────── 저장 ─────────

  /**
   * 여러 파일을 한 번에 저장합니다.
   * 하나라도 다른 사람이 먼저 바꿨으면 아무것도 저장하지 않습니다.
   */
  async save(changes: Partial<DataFiles>): Promise<Partial<DataFiles>> {
    await this.ensureRoot();
    return this.withWriteLock(() => this.saveLocked(changes));
  }

  /**
   * 저장 잠금: '확인 → 쓰기' 사이에 다른 PC가 끼어들지 못하게 data/.writing 파일을 독점 생성합니다.
   * 다른 PC가 쓰는 중이면 잠깐 기다리고, 오래된 잠금(프로그램이 멈춘 PC)은 지웁니다.
   */
  private async withWriteLock<T>(fn: () => Promise<T>): Promise<T> {
    const lock = path.join(this.root, DIRS.data, WRITE_LOCK);
    const deadline = Date.now() + WRITE_LOCK_WAIT_MS;
    for (;;) {
      try {
        const fh = await fs.open(lock, 'wx');
        await fh.writeFile(JSON.stringify({ by: this.user.displayName || this.user.account, at: new Date().toISOString() }));
        await fh.close();
        break;
      } catch (e) {
        const code = (e as NodeJS.ErrnoException).code;
        if (code === 'ENOENT') throw folderMissing(this.root);
        if (code !== 'EEXIST') throw e;
        const st = await fs.stat(lock).catch(() => null);
        if (st && Date.now() - st.mtimeMs > WRITE_LOCK_STALE_MS) {
          await fs.rm(lock, { force: true });
          continue;
        }
        if (Date.now() > deadline) {
          throw new AppError('FILE_BUSY', '다른 PC가 저장하는 중이라 기다렸지만 끝나지 않았습니다. 잠시 후 다시 저장하세요.');
        }
        await new Promise((r) => setTimeout(r, 40 + Math.random() * 60));
      }
    }
    try {
      return await fn();
    } finally {
      await fs.rm(lock, { force: true }).catch(() => {});
    }
  }

  private async saveLocked(changes: Partial<DataFiles>): Promise<Partial<DataFiles>> {
    const keys = Object.keys(changes) as DataFileKey[];
    if (keys.length === 0) return {};

    // 1) 저장 직전 디스크 상태를 다시 확인.
    //    기준 revision 은 넘겨받은 데이터에 적힌 값(그 데이터를 읽었을 때의 번호),
    //    없으면 마지막으로 불러온 번호를 씁니다.
    const expected = new Map<DataFileKey, number>();
    for (const k of keys) {
      const own = (changes[k] as Partial<FileMeta> | undefined)?.revision;
      const known = typeof own === 'number' ? own : this.loadedRevision.get(k);
      if (known === undefined) {
        throw new AppError('UNKNOWN', `${k} 데이터를 불러오지 않은 상태에서 저장하려고 했습니다. 새로 불러온 뒤 다시 저장하세요.`, { file: k });
      }
      const disk = await this.readRaw(k);
      if (disk.revision !== known) {
        throw conflictError(k, disk.updatedBy, disk.updatedByName, disk.updatedAt);
      }
      expected.set(k, known);
    }

    // 2) 지금 상태 백업
    await this.backup();

    // 3) 새 revision 으로 기록
    const iso = this.now().toISOString();
    const saved: Partial<DataFiles> = {};
    for (const k of keys) {
      const next = {
        ...(changes[k] as object),
        revision: expected.get(k)! + 1,
        updatedBy: this.user.account,
        updatedByName: this.user.displayName,
        updatedAt: iso,
      } as DataFiles[typeof k];
      await writeFileAtomic(this.fileOf(k), JSON.stringify(next, null, 2));
      this.loadedRevision.set(k, next.revision);
      (saved as any)[k] = next;
    }
    await this.pruneBackups();
    return saved;
  }

  // ───────── 백업 ─────────

  private backupDir() {
    return path.join(this.root, DIRS.backup);
  }

  /** 현재 data/ 의 파일들을 백업/<시각>_<계정>/ 에 복사합니다. */
  async backup(): Promise<string | null> {
    const dest = path.join(this.backupDir(), `${stamp(this.now())}_${safeName(this.user.account)}`);
    const copied: DataFileKey[] = [];
    await fs.mkdir(dest, { recursive: true });
    for (const k of DATA_FILE_KEYS) {
      const src = this.fileOf(k);
      if (!(await exists(src))) continue;
      await fs.copyFile(src, path.join(dest, `${k}.json`));
      copied.push(k);
    }
    if (copied.length === 0) {
      await fs.rm(dest, { recursive: true, force: true });
      return null;
    }
    const meta: BackupEntry = {
      id: path.basename(dest),
      createdAt: this.now().toISOString(),
      createdBy: this.user.displayName || this.user.account,
      files: copied,
    };
    await writeFileAtomic(path.join(dest, 'backup.json'), JSON.stringify(meta, null, 2));
    return dest;
  }

  async listBackups(): Promise<BackupEntry[]> {
    const dir = this.backupDir();
    if (!(await isDir(dir))) return [];
    const names = (await fs.readdir(dir)).filter((n) => /^\d{8}-\d{6}-\d{3}_/.test(n)).sort().reverse();
    const out: BackupEntry[] = [];
    for (const n of names) {
      try {
        out.push(JSON.parse(await readText(path.join(dir, n, 'backup.json'))));
      } catch {
        out.push({ id: n, createdAt: '', createdBy: '', files: [] });
      }
    }
    return out;
  }

  private async keepCount(): Promise<number> {
    try {
      const s = await this.readRaw('settings');
      return Math.max(1, Number(s.backupKeep) || 50);
    } catch {
      return 50;
    }
  }

  /** 설정 개수를 넘는 오래된 백업부터 지웁니다. */
  async pruneBackups(): Promise<number> {
    const keep = await this.keepCount();
    const dir = this.backupDir();
    if (!(await isDir(dir))) return 0;
    const names = (await fs.readdir(dir)).filter((n) => /^\d{8}-\d{6}-\d{3}_/.test(n)).sort();
    const remove = names.slice(0, Math.max(0, names.length - keep));
    for (const n of remove) await fs.rm(path.join(dir, n), { recursive: true, force: true });
    return remove.length;
  }

  /**
   * 백업 시점으로 되돌립니다. 되돌리기 전 현재 상태도 백업합니다.
   * 다른 PC가 변경을 알아챌 수 있도록 revision 은 현재보다 크게 올립니다.
   */
  async restoreBackup(id: string): Promise<void> {
    if (!/^\d{8}-\d{6}-\d{3}_[^\\/]+$/.test(id)) throw new AppError('UNKNOWN', '백업 이름이 올바르지 않습니다.');
    const src = path.join(this.backupDir(), id);
    if (!(await isDir(src))) throw new AppError('UNKNOWN', `백업(${id})을 찾을 수 없습니다. 목록을 새로 불러오세요.`);
    await this.ensureRoot();
    await this.withWriteLock(() => this.restoreLocked(src));
  }

  private async restoreLocked(src: string): Promise<void> {
    await this.backup();
    const iso = this.now().toISOString();
    for (const k of DATA_FILE_KEYS) {
      const f = path.join(src, `${k}.json`);
      if (!(await exists(f))) continue;
      const fromBackup = migrate(k, JSON.parse(await readText(f)));
      let currentRev = -1;
      try {
        currentRev = (await this.readRaw(k)).revision;
      } catch {
        /* 손상된 파일도 복원으로 덮어씁니다 */
      }
      const next = {
        ...fromBackup,
        revision: Math.max(currentRev, fromBackup.revision) + 1,
        updatedBy: this.user.account,
        updatedByName: this.user.displayName,
        updatedAt: iso,
      };
      await writeFileAtomic(this.fileOf(k), JSON.stringify(next, null, 2));
      this.loadedRevision.set(k, next.revision);
    }
    await this.pruneBackups();
  }
}
