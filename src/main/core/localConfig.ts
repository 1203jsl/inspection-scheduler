import { promises as fs } from 'node:fs';
import path from 'node:path';
import { readText, writeFileAtomic } from './fsSafe';

/** 이 PC에만 저장되는 설정 (동기화 폴더가 아니라 사용자 AppData) */
export interface LocalConfig {
  dataDir: string | null;
}

const FILE = 'local-config.json';

export async function readLocalConfig(dir: string): Promise<LocalConfig> {
  try {
    const j = JSON.parse(await readText(path.join(dir, FILE)));
    return { dataDir: typeof j.dataDir === 'string' && j.dataDir ? j.dataDir : null };
  } catch {
    return { dataDir: null };
  }
}

export async function writeLocalConfig(dir: string, cfg: LocalConfig): Promise<void> {
  await fs.mkdir(dir, { recursive: true });
  await writeFileAtomic(path.join(dir, FILE), JSON.stringify(cfg, null, 2));
}
