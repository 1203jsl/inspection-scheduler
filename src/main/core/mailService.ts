import type { MailDraft, MailLogEntry } from '../../shared/types';
import { DataStore } from './dataStore';
import { AppError } from './errors';
import { appendMailLog, checkDraft, logEntry, resolveAttachments, sendViaSmtp } from './mailer';
import { SmtpCredentialStore } from './secrets';

type User = { account: string; displayName: string };

/**
 * 메일 발송·기록.
 * - 자동(SMTP): 보내고 성공/실패를 모두 기록
 * - 수동: 사용자가 웹메일로 보낸 뒤 [발송 완료로 기록]
 * 업체 메일은 화면의 미리보기를 거친 draft 만 받습니다 (자동 즉시 발송 없음).
 */
export class MailService {
  constructor(
    private store: DataStore,
    private creds: SmtpCredentialStore,
    private user: () => User,
    private now: () => Date = () => new Date(),
  ) {}

  async send(draft: MailDraft): Promise<MailLogEntry> {
    checkDraft(draft);
    const settings = await this.store.load('settings');
    if (settings.mail.method !== 'smtp') {
      throw new AppError('UNKNOWN', '발송 방식이 수동입니다. 웹메일로 보낸 뒤 [발송 완료로 기록]을 누르세요.');
    }
    const attachments = await resolveAttachments(this.store.root, draft.attachments);
    const { account, password } = await this.creds.credentials();
    const u = this.user();
    try {
      await sendViaSmtp({ mail: settings.mail, account, password, senderName: u.displayName, draft, attachments });
    } catch (e) {
      const msg = e instanceof AppError ? e.message : String(e);
      // 실패도 기록 (기록 실패는 원래 오류를 가리지 않게 무시)
      await appendMailLog(this.store, logEntry(draft, u, 'smtp', 'failed', this.now(), msg)).catch(() => {});
      throw e;
    }
    const entry = logEntry(draft, u, 'smtp', 'sent', this.now());
    try {
      await appendMailLog(this.store, entry);
    } catch (e) {
      const why = e instanceof AppError ? e.message : String(e);
      throw new AppError('UNKNOWN', `메일은 보냈지만 발송 기록을 저장하지 못했습니다 (${why}). 같은 메일을 다시 보내지 말고, 새로 불러온 뒤 메일 이력을 확인하세요.`);
    }
    return entry;
  }

  async recordManual(draft: MailDraft): Promise<MailLogEntry> {
    checkDraft(draft);
    const entry = logEntry(draft, this.user(), 'manual', 'manual', this.now());
    await appendMailLog(this.store, entry);
    return entry;
  }

  /** 설정한 계정으로 본인에게 시험 메일 */
  async test(): Promise<MailLogEntry> {
    const { account } = await this.creds.credentials();
    if (!account) throw new AppError('UNKNOWN', '메일 계정을 먼저 입력하고 설정을 저장하세요.');
    const u = this.user();
    return this.send({
      inspectionId: null,
      kind: 'test',
      to: account,
      cc: '',
      subject: '[검수일정 관리] 테스트 메일',
      body: `${u.displayName} 님,\n\n검수일정 관리 프로그램의 메일 발송 설정이 정상입니다.\n이 메일은 본인에게만 보낸 시험 메일입니다.`,
      attachments: [],
    });
  }
}
