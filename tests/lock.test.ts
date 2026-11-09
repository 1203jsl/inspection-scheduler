import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { LockManager } from '../src/main/core/lock';

let dir: string;
let clock: number;
const now = () => new Date(clock);
const KIM = { account: 'kim.mj', displayName: '김민준', host: 'PC-01' };
const LEE = { account: 'lee.sy', displayName: '이서연', host: 'PC-02' };

beforeEach(async () => {
  dir = await fs.mkdtemp(path.join(os.tmpdir(), 'lock-'));
  clock = new Date(2026, 10, 9, 10, 0).getTime();
});
afterEach(async () => {
  await fs.rm(dir, { recursive: true, force: true });
});

describe('.lock 편집 중 표시', () => {
  it('잠금이 없으면 내가 잡는다', async () => {
    const kim = new LockManager(dir, KIM, () => 30, now);
    const st = await kim.acquire();
    expect(st.mine).toBe(true);
    const lee = new LockManager(dir, LEE, () => 30, now);
    const seen = await lee.read();
    expect(seen.holder?.displayName).toBe('김민준');
    expect(seen.mine).toBe(false);
  });

  it('다른 사람의 유효한 잠금은 덮어쓰지 않는다', async () => {
    await new LockManager(dir, KIM, () => 30, now).acquire();
    const st = await new LockManager(dir, LEE, () => 30, now).acquire();
    expect(st.mine).toBe(false);
    expect(st.holder?.user).toBe('kim.mj');
  });

  it('설정 시간이 지난 잠금은 무시한다', async () => {
    await new LockManager(dir, KIM, () => 30, now).acquire();
    clock += 31 * 60_000;
    const lee = new LockManager(dir, LEE, () => 30, now);
    expect((await lee.read()).holder).toBeNull();
    expect((await lee.acquire()).mine).toBe(true);
  });

  it('heartbeat 로 갱신하면 계속 유효하다', async () => {
    const kim = new LockManager(dir, KIM, () => 30, now);
    await kim.acquire();
    clock += 20 * 60_000;
    await kim.heartbeat();
    clock += 20 * 60_000;
    expect((await new LockManager(dir, LEE, () => 30, now).read()).holder?.user).toBe('kim.mj');
  });

  it('내 잠금만 해제한다', async () => {
    const kim = new LockManager(dir, KIM, () => 30, now);
    await kim.acquire();
    await new LockManager(dir, LEE, () => 30, now).release();
    expect((await kim.read()).mine).toBe(true);
    await kim.release();
    expect((await kim.read()).holder).toBeNull();
  });

  it('깨진 잠금 파일은 무시한다', async () => {
    await fs.writeFile(path.join(dir, '.lock'), 'xx');
    expect((await new LockManager(dir, LEE, () => 30, now).read()).holder).toBeNull();
  });
});
