import { promises as fs } from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import nodemailer from 'nodemailer';
import type { MailDraft, MailLogEntry, MailSettings } from '../../shared/types';
import { parseAddresses } from '../../shared/mail';
import { DataStore } from './dataStore';
import { AppError } from './errors';

// ───────── 주소·첨부 점검 ─────────

export function checkDraft(d: MailDraft): void {
  const to = parseAddresses(d.to);
  if (to.ok.length === 0) throw new AppError('UNKNOWN', '받는 사람 이메일을 입력하세요. 업체 관리 화면에서 업체 담당자 이메일을 확인할 수 있습니다.');
  const bad = [...to.bad, ...parseAddresses(d.cc).bad];
  if (bad.length) throw new AppError('UNKNOWN', `이메일 형식이 올바르지 않습니다: ${bad.join(', ')} (예: name@company.co.kr, 여러 명은 쉼표로 구분)`);
  if (!d.subject.trim()) throw new AppError('UNKNOWN', '제목을 입력하세요.');
}

/** 첨부 경로: 데이터 폴더 기준 상대 경로 또는 PC의 절대 경로 */
export async function resolveAttachments(root: string, list: string[]): Promise<{ filename: string; path: string }[]> {
  const out: { filename: string; path: string }[] = [];
  for (const a of list) {
    const full = path.isAbsolute(a) ? a : path.resolve(root, a);
    if (!path.isAbsolute(a) && !full.startsWith(path.resolve(root) + path.sep)) {
      throw new AppError('UNKNOWN', `데이터 폴더 밖의 첨부 경로는 쓸 수 없습니다: ${a}`);
    }
    try {
      await fs.access(full);
    } catch {
      throw new AppError('UNKNOWN', `첨부파일을 찾을 수 없습니다: ${path.basename(a)}. 파일이 옮겨졌거나 동기화가 끝나지 않았을 수 있습니다. 첨부를 빼고 보내거나 다시 고르세요.`);
    }
    out.push({ filename: path.basename(full), path: full });
  }
  return out;
}

// ───────── SMTP 발송 ─────────

/** nodemailer 오류를 무엇이 문제이고 어떻게 고치는지 알려주는 문구로 */
export function friendlyMailError(e: unknown, host: string, port: number): string {
  const err = e as { code?: string; responseCode?: number; message?: string };
  const code = err.code ?? '';
  const where = `${host}:${port}`;
  if (code === 'EAUTH' || err.responseCode === 535) {
    return '메일 계정 또는 비밀번호가 맞지 않습니다. 설정 > 메일 발송 설정에서 계정과 비밀번호를 다시 입력하세요. (웹메일에서 SMTP 사용이 허용돼 있는지도 사내 IT에 확인하세요)';
  }
  if (code === 'ECONNECTION' || code === 'ECONNREFUSED' || code === 'ETIMEDOUT' || code === 'ENOTFOUND' || code === 'EDNS' || code === 'ESOCKET' && /ECONNREFUSED|ENOTFOUND|ETIMEDOUT/.test(err.message ?? '')) {
    return `메일 서버(${where})에 연결할 수 없습니다. 서버 주소·포트가 맞는지 확인하고, 회사 네트워크에서 SMTP 사용이 허용되는지 사내 IT에 확인하세요. 급하면 수동 발송을 쓰세요.`;
  }
  if (code === 'ESOCKET' || code === 'ETLS' || /SSL|TLS|wrong version/i.test(err.message ?? '')) {
    return `메일 서버(${where})와 보안 연결을 맺지 못했습니다. 보안 방식(SSL/TLS·STARTTLS)과 포트가 서로 맞는지 확인하세요. (보통 SSL/TLS=465, STARTTLS=587)`;
  }
  if (code === 'EENVELOPE' || (err.responseCode && err.responseCode >= 550)) {
    return `메일 서버가 받는 사람 주소를 거부했습니다. 주소가 맞는지 확인하세요. (${err.message ?? ''})`;
  }
  return `메일을 보내지 못했습니다: ${err.message ?? String(e)}`;
}

export interface SendInput {
  mail: MailSettings;
  account: string;
  password: string | null;
  senderName: string;
  draft: MailDraft;
  attachments: { filename: string; path: string }[];
}

export async function sendViaSmtp(inp: SendInput): Promise<void> {
  const { host, port, security } = inp.mail.smtp;
  if (!host.trim()) throw new AppError('UNKNOWN', 'SMTP 서버 주소가 비어 있습니다. 설정 > 메일 발송 설정에서 입력하거나, 발송 방식을 수동으로 바꾸세요.');
  if (!inp.account) throw new AppError('UNKNOWN', '메일 계정이 비어 있습니다. 설정 > 메일 발송 설정에서 이 PC의 메일 계정을 입력하세요.');
  const transport = nodemailer.createTransport({
    host: host.trim(),
    port,
    secure: security === 'ssl',
    requireTLS: security === 'starttls',
    ignoreTLS: security === 'none',
    auth: inp.password ? { user: inp.account, pass: inp.password } : undefined,
    connectionTimeout: 15_000,
    greetingTimeout: 15_000,
    socketTimeout: 30_000,
  });
  try {
    await transport.sendMail({
      from: { name: inp.senderName, address: inp.account },
      to: inp.draft.to,
      cc: inp.draft.cc || undefined,
      subject: inp.draft.subject,
      text: inp.draft.body,
      attachments: inp.attachments,
    });
  } catch (e) {
    throw new AppError('UNKNOWN', friendlyMailError(e, host, port));
  } finally {
    transport.close();
  }
}

// ───────── 발송 기록 ─────────

export function logEntry(d: MailDraft, user: { account: string; displayName: string }, method: 'smtp' | 'manual', result: MailLogEntry['result'], now: Date, error?: string): MailLogEntry {
  return {
    id: `M-${now.getTime().toString(36)}-${crypto.randomBytes(2).toString('hex')}`,
    at: now.toISOString(),
    by: user.account,
    byName: user.displayName,
    inspectionId: d.inspectionId,
    ...(d.inspectionIds && d.inspectionIds.length > 1 ? { inspectionIds: [...d.inspectionIds] } : {}),
    ...(d.reminderDate ? { reminderDate: d.reminderDate, inspectorId: d.inspectorId } : {}),
    kind: d.kind,
    to: d.to.trim(),
    cc: d.cc.trim(),
    subject: d.subject,
    body: d.body,
    attachments: d.attachments.map((a) => path.basename(a)),
    method,
    result,
    error,
  };
}

/**
 * mail_log.json 에 한 줄 추가. 기록은 '덧붙이기'라 다른 사람이 먼저 저장했으면
 * 새로 읽어서 다시 덧붙입니다(덮어쓰지 않음).
 */
export async function appendMailLog(store: DataStore, entry: MailLogEntry, tries = 4): Promise<void> {
  for (let n = 1; ; n++) {
    const cur = await store.load('mail_log');
    try {
      await store.save({ mail_log: { ...cur, items: [...cur.items, entry] } });
      return;
    } catch (e) {
      if (e instanceof AppError && e.code === 'CONFLICT' && n < tries) continue;
      throw e;
    }
  }
}
