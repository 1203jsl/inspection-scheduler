import type { MailDraft, MailLogEntry } from '../../shared/types';
import { fillTemplate } from '../../shared/mail';
import { DEFAULT_NOTIFY, DEFAULT_REMINDER_TEMPLATE, reminderHandled, reminderTargets, reminderVars } from '../../shared/reminders';
import { ymd } from '../../shared/dates';
import { DataStore } from './dataStore';
import { AppError } from './errors';
import { appendMailLog, logEntry, sendViaSmtp } from './mailer';
import { SmtpCredentialStore } from './secrets';

type User = { account: string; displayName: string };

export type ReminderRun =
  | { status: 'off' }               // 설정에서 알림 메일 꺼짐
  | { status: 'manual' }            // 발송 방식이 수동이라 자동으로 못 보냄
  | { status: 'notConfigured'; reason: string } // 이 PC에 SMTP 계정이 없음 (다른 PC가 보내도록 아무것도 안 함)
  | { status: 'none' }              // 오늘 보낼 대상 없음 (또는 다른 PC가 이미 보냄)
  | { status: 'preview'; drafts: MailDraft[] }
  | { status: 'done'; sent: string[]; failed: { to: string; error: string }[] };

/**
 * 검수자 알림 메일. 그날 처음 프로그램을 연 PC가 보냅니다.
 * 보내기 전에 mail_log 에 '보내는 중'을 먼저 기록(선점)해서 다른 PC가 같은 메일을 또 보내지 않게 합니다.
 */
export class ReminderService {
  constructor(
    private store: DataStore,
    private creds: SmtpCredentialStore,
    private user: () => User,
    private now: () => Date = () => new Date(),
  ) {}

  /** 오늘 보낼 알림 메일 초안 (이미 보낸 검수자는 뺌) */
  async drafts(): Promise<MailDraft[]> {
    const settings = await this.store.load('settings');
    const [insp, people, log] = [await this.store.load('inspections'), await this.store.load('inspectors'), await this.store.load('mail_log')];
    const notify = settings.notify ?? DEFAULT_NOTIFY;
    const tpl = settings.reminderTemplate ?? DEFAULT_REMINDER_TEMPLATE;
    const today = ymd(this.now());
    return reminderTargets(insp.items, people.items, today, notify.daysBefore)
      .filter((t) => !reminderHandled(log.items, today, t.inspector.id, this.now()))
      .map((t) => {
        const vars = reminderVars(t.inspector, t.items, today, notify.daysBefore);
        return {
          inspectionId: null,
          inspectionIds: t.items.map((i) => i.id),
          kind: 'reminder' as const,
          to: t.inspector.email,
          cc: '',
          subject: fillTemplate(tpl.subject, vars).text,
          body: fillTemplate(tpl.body, vars).text,
          attachments: [],
          reminderDate: today,
          inspectorId: t.inspector.id,
        };
      });
  }

  /** 프로그램을 열 때(그리고 날짜가 바뀔 때) 호출 */
  async auto(): Promise<ReminderRun> {
    const settings = await this.store.load('settings');
    const notify = settings.notify ?? DEFAULT_NOTIFY;
    if (!notify.mail) return { status: 'off' };
    if (settings.mail.method !== 'smtp') return { status: 'manual' };
    const ready = await this.ready(settings.mail.smtp.host);
    if (ready) return { status: 'notConfigured', reason: ready };
    const drafts = await this.drafts();
    if (drafts.length === 0) return { status: 'none' };
    if (notify.mailPreview) return { status: 'preview', drafts };
    return this.send(drafts);
  }

  private async ready(host: string): Promise<string | null> {
    if (!host.trim()) return 'SMTP 서버가 설정되지 않았습니다.';
    const c = await this.creds.credentials().catch(() => ({ account: '', password: null }));
    if (!c.account || !c.password) return '이 PC에 메일 계정·비밀번호가 저장되지 않았습니다.';
    return null;
  }

  /** 선점 → 발송 → 결과 기록 */
  async send(drafts: MailDraft[]): Promise<ReminderRun> {
    const settings = await this.store.load('settings');
    if (settings.mail.method !== 'smtp') throw new AppError('UNKNOWN', '알림 메일은 자동(SMTP) 발송 방식에서만 보낼 수 있습니다.');
    const notReady = await this.ready(settings.mail.smtp.host);
    if (notReady) throw new AppError('UNKNOWN', `알림 메일을 보낼 수 없습니다: ${notReady} 설정 > 메일 발송 설정을 확인하세요.`);
    const claimed = await this.claim(drafts);
    if (claimed.length === 0) return { status: 'none' };
    const { account, password } = await this.creds.credentials();
    const u = this.user();
    const sent: string[] = [];
    const failed: { to: string; error: string }[] = [];
    for (const d of claimed) {
      try {
        await sendViaSmtp({ mail: settings.mail, account, password, senderName: u.displayName, draft: d, attachments: [] });
        await appendMailLog(this.store, logEntry(d, u, 'smtp', 'sent', this.now()));
        sent.push(d.to);
      } catch (e) {
        const msg = e instanceof AppError ? e.message : String(e);
        await appendMailLog(this.store, logEntry(d, u, 'smtp', 'failed', this.now(), msg)).catch(() => {});
        failed.push({ to: d.to, error: msg });
      }
    }
    return { status: 'done', sent, failed };
  }

  /**
   * 아직 아무 PC도 처리하지 않은 검수자만 '보내는 중'으로 기록하고 그 초안을 돌려줍니다.
   * 다른 PC가 같은 순간에 기록했으면(CONFLICT) 새로 읽어 다시 판단합니다.
   */
  private async claim(drafts: MailDraft[], tries = 4): Promise<MailDraft[]> {
    const u = this.user();
    for (let n = 1; ; n++) {
      const log = await this.store.load('mail_log');
      const mine = drafts.filter((d) => d.reminderDate && d.inspectorId && !reminderHandled(log.items, d.reminderDate, d.inspectorId, this.now()));
      if (mine.length === 0) return [];
      const claims: MailLogEntry[] = mine.map((d) => logEntry(d, u, 'smtp', 'claimed', this.now()));
      try {
        await this.store.save({ mail_log: { ...log, items: [...log.items, ...claims] } });
        return mine;
      } catch (e) {
        if (e instanceof AppError && e.code === 'CONFLICT' && n < tries) continue;
        throw e;
      }
    }
  }
}
