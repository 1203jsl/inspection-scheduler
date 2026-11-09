/**
 * 메일 템플릿·치환값·묶음·미발송 판단.
 * 업체에 가는 메일은 이 값으로 채운 뒤 미리보기 화면에서 사용자가 확인·수정하고 보냅니다.
 *
 * 업체 회신은 '일정이 확정되면' 보냅니다. 같은 업체에 같은 날 확정한 건은 한 통으로 묶습니다.
 */
import type { Inspection, MailKind, MailLogEntry, MailSettings, MailTemplate } from './types';
import { fmtYMD, ymd } from './dates';

export const DEFAULT_MAIL_SETTINGS: MailSettings = {
  method: 'manual',
  // 사내 IT 확인 전이라 비워 둡니다
  smtp: { host: '', port: 587, security: 'starttls' },
  webmailUrl: '',
  defaultCc: '',
};

const SIGN = '\n\n감사합니다.\n{보내는사람} 드림';

/** 5단계까지 쓰던 확정 템플릿 (사용자가 고치지 않았으면 새 기본값으로 바꾸기 위해 보관) */
export const OLD_CONFIRMED_TEMPLATE: MailTemplate = {
  subject: '[검수 일정 확정 안내] {신청번호} {품명} ({업체명})',
  body: '{담당자명} 님께,\n\n신청번호 {신청번호} ({품명}) 검수 일정이 아래와 같이 확정되었습니다.\n\n- 검수 일시: {검수일} {검수시간}\n- 검수 장소: {검수장소}\n- 검수 수량: {수량}\n- 담당 검수자: {검수자}\n\n일정에 맞춰 준비 부탁드립니다.' + SIGN,
};

export const DEFAULT_TEMPLATES: Record<MailKind, MailTemplate> = {
  registered: {
    subject: '[검수 신청 접수] {신청번호} {품명} ({업체명})',
    body: '{담당자명} 님께,\n\n요청하신 검수 신청이 접수되었습니다.\n\n{목록}\n\n일정은 협의 후 확정되면 다시 안내드리겠습니다.' + SIGN,
  },
  confirmed: {
    subject: '[검수 일정 확정 안내] {업체명} {건수}건 ({확정일})',
    body: '{담당자명} 님께,\n\n요청하신 검수 신청 {건수}건의 일정이 아래와 같이 확정되어 안내드립니다.\n\n{목록}\n\n일정에 맞춰 준비 부탁드립니다.\n내용을 수정하실 때는 엑셀의 [신청번호] 칸에 위 신청번호를 적어 회신해 주세요.' + SIGN,
  },
  changed: {
    subject: '[검수 일정 변경 안내] {신청번호} {품명} ({업체명})',
    body: '{담당자명} 님께,\n\n신청번호 {신청번호} ({품명}) 검수 일정이 아래와 같이 변경되었습니다.\n\n- 변경 일시: {검수일} {검수시간}\n- 검수 장소: {검수장소}\n- 변경 사유: {변경사유}\n\n확정되면 다시 안내드리겠습니다.' + SIGN,
  },
  canceled: {
    subject: '[검수 취소 안내] {신청번호} {품명} ({업체명})',
    body: '{담당자명} 님께,\n\n신청번호 {신청번호} ({품명}) 검수가 취소되었습니다.\n\n- 취소 사유: {취소사유}\n\n문의 사항이 있으시면 회신 부탁드립니다.' + SIGN,
  },
  pass: {
    subject: '[검수 결과 안내 - 합격] {신청번호} {품명} ({업체명})',
    body: '{담당자명} 님께,\n\n신청번호 {신청번호} ({품명}) 검수 결과를 알려드립니다.\n\n- 검수일: {검수일}\n- 결과: {결과}\n\n수고하셨습니다.' + SIGN,
  },
  fail: {
    subject: '[검수 결과 안내 - 불합격] {신청번호} {품명} ({업체명})',
    body: '{담당자명} 님께,\n\n신청번호 {신청번호} ({품명}) 검수 결과를 알려드립니다.\n\n- 검수일: {검수일}\n- 결과: {결과}\n- 불합격 사유: {불합격사유}\n\n시정 조치 후 재검수 일정을 협의해 주시기 바랍니다.' + SIGN,
  },
};

/** 설정 화면에 보여줄 치환값 (같은 뜻의 다른 이름도 받아 줌) */
export const PLACEHOLDERS = [
  '{업체명}', '{신청번호}', '{품명}', '{품번}', '{수량}', '{검수품목}', '{희망일}', '{희망시간}',
  '{검수일}', '{검수시간}', '{검수장소}', '{검수자}', '{담당자명}', '{목록}', '{건수}', '{확정일}',
  '{결과}', '{불합격사유}', '{변경사유}', '{취소사유}', '{보내는사람}',
];

function lastReason(i: Inspection, to: string): string {
  const h = [...i.history].reverse().find((x) => x.field === 'status' && x.to === to);
  return h?.reason ?? '';
}

const dash = (v: unknown) => (v === null || v === undefined || v === '' ? '-' : String(v));

/** {목록}의 한 건 (확정·변경은 일정까지) */
export function itemLine(i: Inspection, kind: MailKind, inspectorName: (id: string | null) => string): string {
  const head = `- ${i.id} · ${dash(i.itemName)}${i.spec ? ` (${i.spec})` : ''} · 수량 ${dash(i.quantity)}`;
  if (kind === 'confirmed' || kind === 'changed') {
    const who = inspectorName(i.schedule.inspectorId);
    return `${head}\n  검수 일시: ${dash(fmtYMD(i.schedule.date))} ${i.schedule.time ?? ''} · 장소: ${dash(i.schedule.location)}${who ? ` · 검수자: ${who}` : ''}`;
  }
  if (kind === 'registered') return `${head}\n  희망 일시: ${dash(fmtYMD(i.requestedDate))} ${i.requestedTime ?? ''}`;
  return head;
}

/**
 * 치환값 표. 여러 건을 묶으면 {목록}·{건수}에 모두 들어가고,
 * 나머지 한 건짜리 치환값({신청번호} 등)은 첫 건 기준입니다.
 */
export function mailVars(
  items: Inspection | Inspection[], inspectorName: (id: string | null) => string, senderName: string, kind: MailKind = 'confirmed',
): Record<string, string> {
  const list = Array.isArray(items) ? items : [items];
  const i = list[0];
  const result = i.result ? (i.result.verdict === 'pass' ? '합격' : '불합격') : '-';
  const v: Record<string, string> = {
    업체명: dash(i.vendorName),
    신청번호: list.map((x) => x.id).join(', '),
    품명: dash(i.itemName),
    품번: dash(i.spec),
    '품번/규격': dash(i.spec),
    수량: dash(i.quantity),
    검수품목: dash(i.category),
    희망일: dash(fmtYMD(i.requestedDate)),
    희망시간: dash(i.requestedTime),
    검수일: dash(fmtYMD(i.result?.date ?? i.schedule.date)),
    검수시간: dash(i.schedule.time),
    검수장소: dash(i.schedule.location),
    검수자: dash(inspectorName(i.result?.inspectorId ?? i.schedule.inspectorId)),
    담당자명: dash(i.contactName),
    목록: list.map((x) => itemLine(x, kind, inspectorName)).join('\n'),
    건수: String(list.length),
    확정일: dash(fmtYMD(i.confirmedAt ? ymd(new Date(i.confirmedAt)) : null)),
    결과: result,
    불합격사유: dash(i.result?.failReason),
    변경사유: dash(lastReason(i, 'changed')),
    취소사유: dash(lastReason(i, 'canceled')),
    보내는사람: dash(senderName),
  };
  v.시간 = v.검수시간;
  v.장소 = v.검수장소;
  return v;
}

/** {이름} 을 값으로 바꿉니다. 모르는 이름은 그대로 두고 따로 알려줍니다. */
export function fillTemplate(text: string, vars: Record<string, string>): { text: string; unknown: string[] } {
  const unknown = new Set<string>();
  const out = text.replace(/\{([^{}\n]{1,20})\}/g, (m, name: string) => {
    const k = name.trim();
    if (k in vars) return vars[k];
    unknown.add(m);
    return m;
  });
  return { text: out, unknown: [...unknown] };
}

/**
 * 지금 상태에서 업체에 보낼 메일 종류.
 * 등록(신청접수) 단계에서는 보내지 않고, 일정이 확정되면 회신합니다.
 */
export function kindForStatus(i: Pick<Inspection, 'status' | 'result'>): MailKind | null {
  switch (i.status) {
    case 'confirmed': return 'confirmed';
    case 'changed': return 'changed';
    case 'canceled': return 'canceled';
    case 'done': return i.result?.verdict === 'fail' ? 'fail' : 'pass';
    default: return null;
  }
}

/** 이 메일이 이 건을 다뤘는지 (묶음 메일 포함) */
export function mailCovers(m: Pick<MailLogEntry, 'inspectionId' | 'inspectionIds'>, id: string): boolean {
  return m.inspectionId === id || !!m.inspectionIds?.includes(id);
}

/**
 * 현재 상태가 된 뒤 그 안내 메일을 아직 안 보냈으면 그 종류를 돌려줍니다.
 * (실패한 발송은 보낸 것으로 치지 않음)
 */
export function pendingMail(i: Inspection, log: MailLogEntry[]): MailKind | null {
  const kind = kindForStatus(i);
  if (!kind) return null;
  const since = [...i.history].reverse().find((h) => h.field === 'status' && h.to === i.status)?.at ?? i.createdAt ?? '';
  const done = log.some((m) => mailCovers(m, i.id) && m.kind === kind && m.result !== 'failed' && m.at >= since);
  return done ? null : kind;
}

/**
 * 한 통으로 묶을 건.
 * 확정 회신: 같은 업체 + 같은 날 확정(현재 '확정' 상태)인 건. 그 밖의 메일은 그 건 하나.
 */
export function mailGroup(insp: Inspection, all: Inspection[], kind: MailKind): Inspection[] {
  if (kind !== 'confirmed' || insp.status !== 'confirmed' || !insp.confirmedAt) return [insp];
  const day = ymd(new Date(insp.confirmedAt));
  const same = all.filter((x) =>
    x.vendorName === insp.vendorName && x.status === 'confirmed' && x.confirmedAt && ymd(new Date(x.confirmedAt)) === day);
  return same.sort((a, b) => (a.schedule.date ?? '').localeCompare(b.schedule.date ?? '') || a.id.localeCompare(b.id));
}

/** 받는 사람: 묶인 건들의 업체 담당자 (중복 제거) */
export function recipientsFor(items: Inspection[], fallback?: { email: string; name: string }): string {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const i of items) {
    const email = (i.contactEmail || '').trim();
    if (!email || seen.has(email.toLowerCase())) continue;
    seen.add(email.toLowerCase());
    out.push(i.contactName ? `${i.contactName} <${email}>` : email);
  }
  if (!out.length && fallback?.email) out.push(fallback.name ? `${fallback.name} <${fallback.email}>` : fallback.email);
  return out.join(', ');
}

/** '이름 <a@b.c>, d@e.f' 형식에서 주소만 뽑아 점검 */
export function parseAddresses(s: string): { ok: string[]; bad: string[] } {
  const ok: string[] = [];
  const bad: string[] = [];
  for (const part of s.split(/[,;\n]/).map((x) => x.trim()).filter(Boolean)) {
    const m = part.match(/<([^<>]+)>\s*$/);
    const addr = (m ? m[1] : part).trim();
    (/^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]{2,}$/.test(addr) ? ok : bad).push(part);
  }
  return { ok, bad };
}
