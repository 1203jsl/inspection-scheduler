import { promises as fs } from 'node:fs';
import path from 'node:path';
import type { LockInfo } from '../../shared/types';
import { readText, writeFileAtomic } from './fsSafe';
import { LOCK_FILE } from './dataStore';

export interface LockStatus {
  /** 지금 유효한 잠금 (오래된 잠금은 null) */
  holder: LockInfo | null;
  /** 내 잠금인지 */
  mine: boolean;
}

/**
 * .lock 파일로 "○○○ 편집 중"을 알립니다.
 * 잠금은 표시용이며 저장을 막지는 않습니다(실제 보호는 revision 비교).
 * heartbeatAt 이 staleMinutes 보다 오래되면 무시합니다.
 */
export class LockManager {
  constructor(
    private root: string,
    private me: { account: string; displayName: string; host: string },
    private staleMinutes: () => number,
    private now: () => Date = () => new Date(),
  ) {}

  private file() {
    return path.join(this.root, LOCK_FILE);
  }

  private isMine(l: LockInfo) {
    return l.user === this.me.account && l.host === this.me.host;
  }

  private isStale(l: LockInfo) {
    const t = Date.parse(l.heartbeatAt);
    if (Number.isNaN(t)) return true;
    return this.now().getTime() - t > this.staleMinutes() * 60_000;
  }

  async read(): Promise<LockStatus> {
    let l: LockInfo;
    try {
      l = JSON.parse(await readText(this.file()));
    } catch {
      return { holder: null, mine: false }; // 없거나 깨진 잠금은 무시
    }
    if (!l || typeof l.user !== 'string' || this.isStale(l)) return { holder: null, mine: false };
    return { holder: l, mine: this.isMine(l) };
  }

  /** 편집 시작. 다른 사람의 유효한 잠금이 있으면 덮어쓰지 않고 그 정보를 돌려줍니다. */
  async acquire(): Promise<LockStatus> {
    const cur = await this.read();
    if (cur.holder && !cur.mine) return cur;
    const iso = this.now().toISOString();
    const info: LockInfo = {
      user: this.me.account,
      displayName: this.me.displayName,
      host: this.me.host,
      startedAt: cur.mine && cur.holder ? cur.holder.startedAt : iso,
      heartbeatAt: iso,
    };
    await writeFileAtomic(this.file(), JSON.stringify(info, null, 2));
    return { holder: info, mine: true };
  }

  /** 편집 중임을 갱신 (내 잠금일 때만) */
  async heartbeat(): Promise<void> {
    const cur = await this.read();
    if (cur.mine) await this.acquire();
  }

  /** 편집 끝. 내 잠금일 때만 지웁니다. */
  async release(): Promise<void> {
    const cur = await this.read();
    if (cur.mine) await fs.rm(this.file(), { force: true });
  }
}
