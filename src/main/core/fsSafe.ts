import { promises as fs } from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { AppError } from './errors';

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

/**
 * 임시 파일에 다 쓴 뒤 원래 파일과 교체합니다.
 * 중간에 끊겨도 원래 파일은 그대로 남습니다.
 */
export async function writeFileAtomic(file: string, content: string): Promise<void> {
  const dir = path.dirname(file);
  const tmp = path.join(dir, `.${path.basename(file)}.${process.pid}.${crypto.randomBytes(4).toString('hex')}.tmp`);
  const fh = await fs.open(tmp, 'w');
  try {
    await fh.writeFile(content, 'utf8');
    await fh.sync();
  } finally {
    await fh.close();
  }
  try {
    await retry(() => fs.rename(tmp, file));
  } catch (e) {
    await fs.rm(tmp, { force: true }).catch(() => {});
    const code = (e as NodeJS.ErrnoException).code;
    if (['EPERM', 'EBUSY', 'EACCES'].includes(code ?? '')) {
      throw new AppError(
        'FILE_BUSY',
        `${path.basename(file)} 파일이 다른 프로그램에서 사용 중이라 저장하지 못했습니다. 잠시 후 다시 저장하세요. 계속되면 클라우드독 동기화 상태를 확인하세요.`,
        { path: file },
      );
    }
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
