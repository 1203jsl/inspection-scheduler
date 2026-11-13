/**
 * 검수자 알림 (우리 쪽 검수자에게만).
 * - 팝업: 이 PC 사용자와 연결된 검수자에게 배정된 다가오는 검수
 * - 메일: 그날 처음 프로그램을 연 PC가 대상 검수자 모두에게 한 번 (기록으로 중복 방지)
 */
import type { CurrentUser, Inspection, Inspector, MailLogEntry, MailTemplate, NotifySettings } from './types';
import { addDays, fmtMD } from './dates';

export const DEFAULT_NOTIFY: NotifySettings = { daysBefore: 1, popup: true, mail: true, mailPreview: false };

export const DEFAULT_REMINDER_TEMPLATE: MailTemplate = {
  subject: '[검수 알림] {검수자} 님 배정 검수 {건수}건 ({기간})',
  body: '{검수자} 님,\n\n배정된 검수 일정을 알려드립니다.\n\n{목록}\n\n상세 내용은 검수일정 관리 프로그램에서 확인하세요.\n(이 메일은 검수일정 관리 프로그램이 자동으로 보냈습니다.)',
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
  return `- ${fmtMD(i.schedule.date)} ${i.schedule.time ?? ''} · ${i.vendorName} · ${i.itemName} · ${i.schedule.location || '장소 미정'} (${i.id})${i.status === 'changed' ? ' [변경됨·재확정 필요]' : ''}`;
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
export function reminderHandled(log: MailLogEntry[], date: string, inspectorId: string, now: Date): boolean {
  const mine = log.filter((m) => m.kind === 'reminder' && m.reminderDate === date && m.inspectorId === inspectorId);
  const last = mine.at(-1);
  if (!last) return false;
  if (last.result === 'sent' || last.result === 'manual') return true;
  if (last.result === 'claimed') return now.getTime() - Date.parse(last.at) < CLAIM_STALE_MS;
  return false;
}
