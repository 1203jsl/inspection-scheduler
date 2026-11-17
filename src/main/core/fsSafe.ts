import { promises as fs } from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { AppError, isPermError, noPermission } from './errors';

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** 잠깐 잠긴 파일(동기화 프로그램·백신)을 고려해 몇 번 다시 시도합니다. */
async function retry<T>(fn: () => Promise<T>, tries = 5, delayMs = 150): Promise<T> {
  let last: unknown;
  for (let i = 0; i < tries; i++) {
    try {
      return await fn();
    } catch (e) {
      last = e;
      const code = (e as NodeJS.ErrnoException).code;
      if (!['EPERM', 'EBUSY', 'EACCES', 'EAGAIN'].includes(code ?? '')) throw e;
      await sleep(delayMs * (i + 1));
    }
  }
  throw last;
}

/** 임시 파일 이름. 점(.)으로 시작하거나 .tmp 로 끝나는 파일을 막는 보안·동기화 프로그램이 있어 원래 확장자를 씁니다. */
export function tempName(file: string): string {
  const ext = path.extname(file) || '.json';
  const stem = path.basename(file, path.extname(file));
  return path.join(path.dirname(file), `${stem}.saving-${process.pid}-${crypto.randomBytes(3).toString('hex')}${ext}`);
}

/** 원래 파일에 바로 씁니다 (임시 파일을 쓸 수 없는 폴더용). */
async function writeDirect(file: string, content: string): Promise<void> {
  try {
    await retry(() => fs.writeFile(file, content, 'utf8'));
  } catch (e) {
    if (isPermError(e)) throw noPermission(file);
    throw e;
  }
}

/**
 * 임시 파일에 다 쓴 뒤 원래 파일과 교체합니다.
 * 중간에 끊겨도 원래 파일은 그대로 남습니다.
 * 회사 보안 정책 등으로 임시 파일을 만들거나 교체할 수 없으면 원래 파일에 바로 씁니다
 * (저장 전에 백업을 남기므로 그래도 되돌릴 수 있습니다).
 */
export async function writeFileAtomic(file: string, content: string): Promise<void> {
  const tmp = tempName(file);
  try {
    const fh = await fs.open(tmp, 'w');
    try {
      await fh.writeFile(content, 'utf8');
      await fh.sync();
    } finally {
      await fh.close();
    }
  } catch (e) {
    await fs.rm(tmp, { force: true }).catch(() => {});
    if (isPermError(e)) return writeDirect(file, content);
    throw e;
  }
  try {
    await retry(() => fs.rename(tmp, file));
  } catch (e) {
    const code = (e as NodeJS.ErrnoException).code;
    if (['EPERM', 'EBUSY', 'EACCES'].includes(code ?? '')) {
      // 바꿔치기(이름 변경)가 막힌 경우: 원래 파일에 바로 써 봄
      try {
        await writeDirect(file, content);
        return;
      } catch (e2) {
        if (e2 instanceof AppError && e2.code === 'NO_PERMISSION') throw e2;
        throw new AppError(
          'FILE_BUSY',
          `${path.basename(file)} 파일이 다른 프로그램에서 사용 중이라 저장하지 못했습니다. 잠시 후 다시 저장하세요. 계속되면 클라우드독 동기화 상태를 확인하세요.`,
          { path: file },
        );
      } finally {
        await fs.rm(tmp, { force: true }).catch(() => {});
      }
    }
    await fs.rm(tmp, { force: true }).catch(() => {});
    throw e;
  }
}

export async function readText(file: string): Promise<string> {
  return retry(() => fs.readFile(file, 'utf8'));
}

export async function exists(p: string): Promise<boolean> {
  try {
    await fs.access(p);
    return true;
  } catch {
    return false;
  }
}

export async function isDir(p: string): Promise<boolean> {
  try {
    return (await fs.stat(p)).isDirectory();
  } catch {
    return false;
  }
}
