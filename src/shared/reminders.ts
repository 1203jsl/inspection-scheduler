/**
 * 검수자 알림 (우리 쪽 검수자에게만).
 * - 팝업: 이 PC 사용자와 연결된 검수자에게 배정된 다가오는 검수
 * - 메일: 그날 처음 프로그램을 연 PC가 대상 검수자 모두에게 한 번 (기록으로 중복 방지)
 */
import type { CurrentUser, Inspection, Inspector, MailLogEntry, MailTemplate, NotifySettings } from './types';
import { addDays, fmtMD, parseYmd } from './dates';
import { pendingMail } from './mail';

export const DEFAULT_NOTIFY: NotifySettings = { daysBefore: 1, popup: true, mail: true, mailPreview: false, resultAlert: true, resultDays: 7 };

export const DEFAULT_REMINDER_TEMPLATE: MailTemplate = {
  subject: '[검수 알림] {검수자} 님 배정 검수 {건수}건 ({기간})',
  body: '{검수자} 님,\n\n배정된 검수 일정을 알려드립니다.\n\n{목록}\n\n상세 내용은 검수일정 관리 프로그램에서 확인하세요.\n(이 메일은 검수일정 관리 프로그램이 자동으로 보냈습니다.)',
};

export const DEFAULT_RESULT_REMINDER_TEMPLATE: MailTemplate = {
  subject: '[결과 통보 요청] {검수자} 님 결과 통보 대기 {건수}건',
  body: '{검수자} 님,\n\n검수 후 업체에 결과를 아직 통보하지 않은 건이 있습니다.\n검수일 다음날부터 {기한}일 안에 결과를 입력하고 업체에 통보해 주세요.\n\n{목록}\n\n검수일정 관리 프로그램에서 [결과 입력] → [저장 후 결과 메일 미리보기]로 통보할 수 있습니다.\n(이 메일은 검수일정 관리 프로그램이 자동으로 보냈습니다.)',
};

/** 일정이 잡힌 상태 (알림 대상) */
const SCHEDULED: Inspection['status'][] = ['confirmed', 'changed'];

/** 오늘부터 days일 뒤까지의 일정 (시간순) */
export function upcoming(items: Inspection[], today: string, days: number, inspectorId?: string): Inspection[] {
  const until = addDays(today, Math.max(0, days));
  return items
    .filter((i) => SCHEDULED.includes(i.status) && i.schedule.date && i.schedule.date >= today && i.schedule.date <= until)
    .filter((i) => !inspectorId || i.schedule.inspectorId === inspectorId)
    .sort((a, b) => `${a.schedule.date} ${a.schedule.time ?? ''}`.localeCompare(`${b.schedule.date} ${b.schedule.time ?? ''}`));
}

/** 이 PC 사용자 = 검수자? (Windows 계정이 같거나, 계정이 비어 있으면 표시 이름이 같음) */
export function matchInspectors(user: Pick<CurrentUser, 'account' | 'displayName'>, inspectors: Inspector[]): Inspector[] {
  const acc = user.account.toLowerCase();
  return inspectors.filter((p) => p.active && (p.account ? p.account.toLowerCase() === acc : p.name === user.displayName));
}

export function reminderLine(i: Inspection): string {
  return `- ${fmtMD(i.schedule.date)} ${i.schedule.time ?? ''} · ${i.vendorName}${i.projectName ? ` · ${i.projectName}` : ''} · ${i.itemName} · ${i.schedule.location || '장소 미정'} (${i.id})${i.status === 'changed' ? ' [변경됨·재확정 필요]' : ''}`;
}

export function reminderVars(p: Inspector, items: Inspection[], today: string, days: number): Record<string, string> {
  const until = addDays(today, Math.max(0, days));
  return {
    검수자: p.name,
    건수: String(items.length),
    목록: items.map(reminderLine).join('\n'),
    기간: days > 0 ? `${fmtMD(today)}~${fmtMD(until)}` : fmtMD(today),
    오늘: fmtMD(today),
  };
}

/** 알림 메일 대상: 사용 중이고 이메일이 있는 검수자 중 배정된 일정이 있는 사람 */
export function reminderTargets(items: Inspection[], inspectors: Inspector[], today: string, days: number): { inspector: Inspector; items: Inspection[] }[] {
  return inspectors
    .filter((p) => p.active && p.email)
    .map((p) => ({ inspector: p, items: upcoming(items, today, days, p.id) }))
    .filter((t) => t.items.length > 0);
}

/** 보내는 중(claimed) 표시가 이 시간보다 오래되면 그 PC가 멈춘 것으로 보고 다시 보냄 */
export const CLAIM_STALE_MS = 30 * 60_000;

/**
 * 이 검수자의 이 날 알림이 이미 처리됐는지.
 * 마지막 기록이 '보냄'이거나 최근의 '보내는 중'이면 처리된 것, '실패'면 다시 보낼 수 있음.
 */
export function reminderHandled(
  log: MailLogEntry[], date: string, inspectorId: string, now: Date, kind: 'reminder' | 'resultReminder' = 'reminder',
): boolean {
  const mine = log.filter((m) => m.kind === kind && m.reminderDate === date && m.inspectorId === inspectorId);
  const last = mine.at(-1);
  if (!last) return false;
  if (last.result === 'sent' || last.result === 'manual') return true;
  if (last.result === 'claimed') return now.getTime() - Date.parse(last.at) < CLAIM_STALE_MS;
  return false;
}

// ───────── 결과 통보 기한 ─────────

const dayDiff = (from: string, to: string) => Math.round((parseYmd(to).getTime() - parseYmd(from).getTime()) / 86_400_000);

export interface ResultNotice {
  /** 실제 검수일(결과에 적힌 날, 없으면 일정) */
  inspected: string;
  /** 통보 기한 = 검수일 + N일 (검수일 다음날부터 N일) */
  deadline: string;
  /** 기한까지 남은 날 (0 = 오늘까지, 음수 = 초과) */
  daysLeft: number;
  /** 결과를 아직 입력하지 않음 / 결과는 입력했지만 업체 통보 메일 미발송 */
  stage: 'noResult' | 'notSent';
  /** 알림 받을 검수자 */
  inspectorId: string | null;
}

/**
 * 검수일이 지났는데 업체에 결과(합격·불합격) 통보가 안 된 건이면 기한 정보를 돌려줍니다.
 * 통보 완료 = 검수완료 상태이고 그 결과 메일을 보냄(수동 기록 포함).
 */
export function resultNotice(i: Inspection, log: MailLogEntry[], today: string, days: number): ResultNotice | null {
  if (!['confirmed', 'changed', 'done'].includes(i.status)) return null;
  const inspected = i.result?.date ?? i.schedule.date;
  if (!inspected || inspected >= today) return null;
  if (i.status === 'done' && pendingMail(i, log) === null) return null;
  const deadline = addDays(inspected, Math.max(1, days));
  return {
    inspected,
    deadline,
    daysLeft: dayDiff(today, deadline),
    stage: i.status === 'done' ? 'notSent' : 'noResult',
    inspectorId: i.result?.inspectorId ?? i.schedule.inspectorId,
  };
}

/** '기한 11-17(화) · 2일 남음' / '오늘까지' / '3일 초과' */
export function deadlineLabel(n: ResultNotice): string {
  const left = n.daysLeft > 0 ? `${n.daysLeft}일 남음` : n.daysLeft === 0 ? '오늘까지' : `${-n.daysLeft}일 초과`;
  return `기한 ${fmtMD(n.deadline)} · ${left}`;
}

export function resultLine(i: Inspection, n: ResultNotice): string {
  const what = n.stage === 'noResult' ? '결과 미입력' : `결과 입력됨(${i.result?.verdict === 'fail' ? '불합격' : '합격'}) · 통보 메일 미발송`;
  return `- ${i.vendorName}${i.projectName ? ` · ${i.projectName}` : ''} · ${i.itemName} (${i.id}) · 검수일 ${fmtMD(n.inspected)} · ${deadlineLabel(n)} · ${what}`;
}

/** 검수자별 결과 통보 대기 건 (기한 임박·초과 순) */
export function resultTargets(
  items: Inspection[], inspectors: Inspector[], log: MailLogEntry[], today: string, days: number,
): { inspector: Inspector; items: { i: Inspection; n: ResultNotice }[] }[] {
  const pending = items
    .map((i) => ({ i, n: resultNotice(i, log, today, days) }))
    .filter((x): x is { i: Inspection; n: ResultNotice } => !!x.n)
    .sort((a, b) => a.n.daysLeft - b.n.daysLeft || a.i.id.localeCompare(b.i.id));
  return inspectors
    .filter((p) => p.active && p.email)
    .map((p) => ({ inspector: p, items: pending.filter((x) => x.n.inspectorId === p.id) }))
    .filter((t) => t.items.length > 0);
}

export function resultVars(p: Inspector, list: { i: Inspection; n: ResultNotice }[], days: number): Record<string, string> {
  const over = list.filter((x) => x.n.daysLeft < 0).length;
  return {
    검수자: p.name,
    건수: String(list.length),
    초과건수: String(over),
    기한: String(days),
    목록: list.map((x) => resultLine(x.i, x.n)).join('\n'),
  };
}
