import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { tempName, writeFileAtomic } from '../src/main/core/fsSafe';
import { AppError, toErrorInfo } from '../src/main/core/errors';

const eperm = (p: string) => Object.assign(new Error(`EPERM: operation not permitted, open '${p}'`), { code: 'EPERM', path: p });

describe('보안 정책으로 파일이 막힌 폴더 (회사 PC 클라우드독)', () => {
  let dir: string;
  beforeEach(async () => { dir = await fs.mkdtemp(path.join(os.tmpdir(), 'fss-')); });
  afterEach(async () => { vi.restoreAllMocks(); await fs.rm(dir, { recursive: true, force: true }); });

  it('임시 파일 이름은 점으로 시작하지 않고 .tmp 가 아니라 원래 확장자', () => {
    const t = path.basename(tempName(path.join(dir, 'inspections.json')));
    expect(t.startsWith('.')).toBe(false);
    expect(t).toMatch(/^inspections\.saving-\d+-[0-9a-f]{6}\.json$/);
  });

  it('임시 파일을 만들 수 없으면 원래 파일에 바로 쓴다', async () => {
    const real = fs.open.bind(fs);
    vi.spyOn(fs, 'open').mockImplementation(((p: string, f: string) => (String(p).includes('.saving-') ? Promise.reject(eperm(p)) : real(p, f))) as any);
    const f = path.join(dir, 'settings.json');
    await writeFileAtomic(f, '{"a":1}');
    expect(await fs.readFile(f, 'utf8')).toBe('{"a":1}');
    expect(await fs.readdir(dir)).toEqual(['settings.json']);
  });

  it('이름 바꾸기가 막혀도 원래 파일에 바로 쓰고 임시 파일은 지운다', async () => {
    vi.spyOn(fs, 'rename').mockImplementation(((a: string) => Promise.reject(eperm(a))) as any);
    const f = path.join(dir, 'settings.json');
    await fs.writeFile(f, 'old');
    await writeFileAtomic(f, 'new');
    expect(await fs.readFile(f, 'utf8')).toBe('new');
    expect(await fs.readdir(dir)).toEqual(['settings.json']);
  }, 10_000);

  it('원래 파일도 쓸 수 없으면 권한 안내 메시지', async () => {
    vi.spyOn(fs, 'open').mockImplementation(((p: string) => Promise.reject(eperm(p))) as any);
    vi.spyOn(fs, 'writeFile').mockImplementation(((p: string) => Promise.reject(eperm(p))) as any);
    const f = path.join(dir, 'inspections.json');
    const err = await writeFileAtomic(f, 'x').catch((e) => e);
    expect(err).toBeInstanceOf(AppError);
    expect(err.code).toBe('NO_PERMISSION');
    expect(err.message).toContain('권한이 없습니다');
  }, 10_000);

  it('예상 못한 권한 오류도 알아듣기 쉬운 문장으로', () => {
    const info = toErrorInfo(eperm('Y:\\검수\\data\\x.json'));
    expect(info.code).toBe('NO_PERMISSION');
  });
});
