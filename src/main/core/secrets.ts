import { promises as fs } from 'node:fs';
import path from 'node:path';
import type { SmtpAccountInfo } from '../../shared/types';
import { AppError } from './errors';
import { readText, writeFileAtomic } from './fsSafe';

/** Electron safeStorage 와 같은 모양 (테스트에서는 가짜를 넣음) */
export interface Crypto {
  available(): boolean;
  encrypt(plain: string): Buffer;
  decrypt(enc: Buffer): string;
}

const FILE = 'smtp-account.json';

/**
 * SMTP 계정과 비밀번호를 '이 PC의 이 Windows 사용자'에게만 저장합니다.
 * - 위치: 사용자 AppData (동기화 폴더가 아님)
 * - 비밀번호는 safeStorage(Windows DPAPI)로 암호화한 값만 파일에 씀
 * - 화면(renderer)에는 비밀번호를 돌려주지 않음
 */
export class SmtpCredentialStore {
  constructor(private dir: string, private crypto: Crypto) {}

  private file() {
    return path.join(this.dir, FILE);
  }

  private async read(): Promise<{ account: string; password?: string }> {
    try {
      const j = JSON.parse(await readText(this.file()));
      return { account: typeof j.account === 'string' ? j.account : '', password: typeof j.password === 'string' ? j.password : undefined };
    } catch {
      return { account: '' };
    }
  }

  async info(): Promise<SmtpAccountInfo> {
    const r = await this.read();
    return { account: r.account, hasPassword: !!r.password, encryptionAvailable: this.crypto.available() };
  }

  /**
   * password: undefined 면 기존 비밀번호 유지, '' 면 지움
   */
  async save(account: string, password?: string): Promise<SmtpAccountInfo> {
    const cur = await this.read();
    let enc = cur.password;
    let refused = false;
    if (password === '') enc = undefined;
    else if (password !== undefined) {
      if (this.crypto.available()) enc = this.crypto.encrypt(password).toString('base64');
      else refused = true; // 평문으로는 절대 저장하지 않음
    }
    await fs.mkdir(this.dir, { recursive: true });
    await writeFileAtomic(this.file(), JSON.stringify({ account: account.trim(), password: enc }, null, 2));
    if (refused) {
      throw new AppError('UNKNOWN', '계정은 저장했지만, 이 PC에서는 비밀번호 암호화를 사용할 수 없어 비밀번호는 저장하지 않았습니다. Windows에 로그인한 상태인지 확인하세요. 수동 발송 방식은 비밀번호 없이 쓸 수 있습니다.');
    }
    return this.info();
  }

  /** 발송할 때만 메인 프로세스 안에서 꺼내 씀 */
  async credentials(): Promise<{ account: string; password: string | null }> {
    const r = await this.read();
    if (!r.password) return { account: r.account, password: null };
    try {
      return { account: r.account, password: this.crypto.decrypt(Buffer.from(r.password, 'base64')) };
    } catch {
      throw new AppError('UNKNOWN', '저장된 메일 비밀번호를 풀 수 없습니다. 다른 Windows 계정에서 저장했을 수 있습니다. 설정 > 메일 발송 설정에서 비밀번호를 다시 입력하세요.');
    }
  }
}
