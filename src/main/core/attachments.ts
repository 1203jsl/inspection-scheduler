import { promises as fs } from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import type { Attachment, AttachmentKind, HistoryEntry, Inspection } from '../../shared/types';
import { ATTACHMENT_KINDS } from '../../shared/types';
import { DataStore, DIRS } from './dataStore';
import { AppError } from './errors';

type User = { account: string; displayName: string };

/** 파일 이름에 쓸 수 없는 글자 정리 */
const safeFileName = (s: string) => s.replace(/[\\/:*?"<>|]+/g, '_').replace(/^\.+/, '_').trim() || '파일';

const pad = (n: number) => String(n).padStart(2, '0');
const stamp = (d: Date) => `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;

async function exists(p: string) {
  return fs.access(p).then(() => true, () => false);
}

/** 같은 이름이 있으면 '이름 (2).확장자' */
async function freeName(dir: string, name: string): Promise<string> {
  const ext = path.extname(name);
  const base = name.slice(0, name.length - ext.length);
  let n = name;
  for (let k = 2; await exists(path.join(dir, n)); k++) n = `${base} (${k})${ext}`;
  return n;
}

/**
 * 한 건만 바꿔 저장. 첨부 추가·삭제는 '그 건에 덧붙이기'라
 * 다른 사람이 먼저 저장했으면 새로 읽어서 다시 적용합니다 (남의 변경을 덮어쓰지 않음).
 */
export async function updateInspection(store: DataStore, id: string, fn: (i: Inspection) => Inspection, tries = 4): Promise<Inspection> {
  for (let n = 1; ; n++) {
    const file = await store.load('inspections');
    const cur = file.items.find((i) => i.id === id);
    if (!cur) throw new AppError('UNKNOWN', `신청번호 ${id}를 찾을 수 없습니다. 새로 불러온 뒤 다시 하세요.`);
    const next = fn(cur);
    try {
      await store.save({ inspections: { ...file, items: file.items.map((i) => (i.id === id ? next : i)) } });
      return next;
    } catch (e) {
      if (e instanceof AppError && e.code === 'CONFLICT' && n < tries) continue;
      throw e;
    }
  }
}

export class AttachmentService {
  constructor(
    private store: DataStore,
    private user: () => User,
    private now: () => Date = () => new Date(),
  ) {}

  private dirOf(id: string) {
    if (!/^[A-Za-z0-9-]+$/.test(id)) throw new AppError('UNKNOWN', '신청번호가 올바르지 않습니다.');
    return path.join(this.store.root, DIRS.attach, id);
  }

  /** PC의 파일들을 첨부/{신청번호}/ 에 복사하고 건에 등록 */
  async add(id: string, paths: string[], kind: AttachmentKind): Promise<Attachment[]> {
    if (!ATTACHMENT_KINDS.includes(kind)) throw new AppError('UNKNOWN', '첨부 종류를 고르세요.');
    if (paths.length === 0) return [];
    const dir = this.dirOf(id);
    const u = this.user();
    const at = this.now().toISOString();
    const added: Attachment[] = [];
    const copied: string[] = [];
    try {
      await fs.mkdir(dir, { recursive: true });
      for (const src of paths) {
        let st;
        try {
          st = await fs.stat(src);
        } catch {
          throw new AppError('UNKNOWN', `파일을 찾을 수 없습니다: ${path.basename(src)}`);
        }
        if (!st.isFile()) throw new AppError('UNKNOWN', `폴더는 첨부할 수 없습니다: ${path.basename(src)}. 파일을 골라 주세요.`);
        const name = await freeName(dir, safeFileName(path.basename(src)));
        try {
          await fs.copyFile(src, path.join(dir, name), fs.constants.COPYFILE_EXCL);
        } catch (e) {
          const code = (e as NodeJS.ErrnoException).code;
          throw new AppError('UNKNOWN', code === 'EBUSY' || code === 'EPERM'
            ? `${path.basename(src)} 파일이 다른 프로그램에서 열려 있어 복사하지 못했습니다. 닫고 다시 넣으세요.`
            : `${path.basename(src)} 파일을 첨부 폴더로 복사하지 못했습니다. 동기화 폴더 연결을 확인하세요.`);
        }
        copied.push(path.join(dir, name));
        added.push({
          id: `A-${Date.now().toString(36)}-${crypto.randomBytes(2).toString('hex')}`,
          kind,
          file: `${DIRS.attach}/${id}/${name}`,
          name: path.basename(src),
          size: st.size,
          addedAt: at,
          addedBy: u.displayName,
        });
      }
      await updateInspection(this.store, id, (i) => ({
        ...i,
        attachments: [...(i.attachments ?? []), ...added],
        history: [
          ...i.history,
          ...added.map((a): HistoryEntry => ({ at, by: u.account, byName: u.displayName, action: 'attach', field: 'attachments', from: null, to: `${a.file.split('/').pop()} (${a.kind})` })),
        ],
      }));
      return added;
    } catch (e) {
      for (const c of copied) await fs.rm(c, { force: true }).catch(() => {});
      throw e;
    }
  }

  /**
   * 첨부 삭제. 실수에 대비해 파일은 지우지 않고 첨부/{신청번호}/_삭제됨/ 으로 옮깁니다.
   */
  async remove(id: string, attachmentId: string): Promise<void> {
    const dir = this.dirOf(id);
    const file = await this.store.load('inspections');
    const att = file.items.find((i) => i.id === id)?.attachments?.find((a) => a.id === attachmentId);
    if (!att) throw new AppError('UNKNOWN', '이미 삭제된 첨부파일입니다. 새로 불러오세요.');
    const u = this.user();
    const now = this.now();
    const src = path.join(this.store.root, att.file);
    const trash = path.join(dir, '_삭제됨');
    let moved: string | null = null;
    if (await exists(src)) {
      await fs.mkdir(trash, { recursive: true });
      moved = path.join(trash, `${stamp(now)}_${path.basename(src)}`);
      try {
        await fs.rename(src, moved);
      } catch {
        throw new AppError('UNKNOWN', `${att.name} 파일이 다른 프로그램에서 열려 있어 삭제하지 못했습니다. 닫고 다시 하세요.`);
      }
    }
    try {
      await updateInspection(this.store, id, (i) => ({
        ...i,
        attachments: (i.attachments ?? []).filter((a) => a.id !== attachmentId),
        history: [...i.history, {
          at: now.toISOString(), by: u.account, byName: u.displayName, action: 'detach', field: 'attachments',
          from: `${path.basename(att.file)} (${att.kind})`, to: null, reason: '삭제 (파일은 _삭제됨 폴더에 보관)',
        }],
      }));
    } catch (e) {
      if (moved) await fs.rename(moved, src).catch(() => {});
      throw e;
    }
  }
}
