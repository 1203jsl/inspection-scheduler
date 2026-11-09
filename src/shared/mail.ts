/**
 * 메일 템플릿·치환값·미발송 판단.
 * 업체에 가는 메일은 이 값으로 채운 뒤 미리보기 화면에서 사용자가 확인·수정하고 보냅니다.
 */
import type { Inspection, MailKind, MailLogEntry, MailSettings, MailTemplate } from './types';
import { fmtYMD } from './dates';

export const DEFAULT_MAIL_SETTINGS: MailSettings = {
  method: 'manual',
  // 사내 IT 확인 전이라 비워 둡니다
  smtp: { host: '', port: 587, security: 'starttls' },
  webmailUrl: '',
  defaultCc: '',
};

const SIGN = '\n\n감사합니다.\n{보내는사람} 드림';

export const DEFAULT_TEMPLATES: Record<MailKind, MailTemplate> = {
  registered: {
    subject: '[검수 신청 접수] {신청번호} {품명} ({업체명})',
    body: '{담당자명} 님께,\n\n요청하신 검수 신청이 접수되어 신청번호를 부여했습니다.\n\n- 신청번호: {신청번호}\n- 품명: {품명} ({품번})\n- 수량: {수량}\n- 희망 일시: {희망일} {희망시간}\n\n일정은 협의 후 확정되면 다시 안내드리겠습니다.\n내용을 수정하실 때는 엑셀의 [신청번호] 칸에 위 번호를 적어 회신해 주세요.' + SIGN,
  },
  confirmed: {
    subject: '[검수 일정 확정 안내] {신청번호} {품명} ({업체명})',
    body: '{담당자명} 님께,\n\n신청번호 {신청번호} ({품명}) 검수 일정이 아래와 같이 확정되었습니다.\n\n- 검수 일시: {검수일} {검수시간}\n- 검수 장소: {검수장소}\n- 검수 수량: {수량}\n- 담당 검수자: {검수자}\n\n일정에 맞춰 준비 부탁드립니다.' + SIGN,
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
  '{검수일}', '{검수시간}', '{검수장소}', '{검수자}', '{담당자명}', '{결과}', '{불합격사유}',
  '{변경사유}', '{취소사유}', '{보내는사람}',
];

function lastReason(i: Inspection, to: string): string {
  const h = [...i.history].reverse().find((x) => x.field === 'status' && x.to === to);
  return h?.reason ?? '';
}

/** 검수 건으로 치환값 표를 만듭니다. 값이 없으면 빈 칸 대신 '-' */
export function mailVars(i: Inspection, inspectorName: (id: string | null) => string, senderName: string): Record<string, string> {
  const d = (v: unknown) => (v === null || v === undefined || v === '' ? '-' : String(v));
  const result = i.result ? (i.result.verdict === 'pass' ? '합격' : '불합격') : '-';
  const v: Record<string, string> = {
    업체명: d(i.vendorName),
    신청번호: i.id,
    품명: d(i.itemName),
    품번: d(i.spec),
    '품번/규격': d(i.spec),
    수량: d(i.quantity),
    검수품목: d(i.category),
    희망일: d(fmtYMD(i.requestedDate)),
    희망시간: d(i.requestedTime),
    검수일: d(fmtYMD(i.result?.date ?? i.schedule.date)),
    검수시간: d(i.schedule.time),
    검수장소: d(i.schedule.location),
    검수자: d(inspectorName(i.result?.inspectorId ?? i.schedule.inspectorId)),
    담당자명: d(i.contactName),
    결과: result,
    불합격사유: d(i.result?.failReason),
    변경사유: d(lastReason(i, 'changed')),
    취소사유: d(lastReason(i, 'canceled')),
    보내는사람: d(senderName),
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

/** 지금 상태에서 업체에 보낼 메일 종류 */
export function kindForStatus(i: Pick<Inspection, 'status' | 'result'>): MailKind | null {
  switch (i.status) {
    case 'received':
    case 'registered': return 'registered';
    case 'confirmed': return 'confirmed';
    case 'changed': return 'changed';
    case 'canceled': return 'canceled';
    case 'done': return i.result?.verdict === 'fail' ? 'fail' : 'pass';
    default: return null;
  }
}

/**
 * 현재 상태가 된 뒤 그 안내 메일을 아직 안 보냈으면 그 종류를 돌려줍니다.
 * (실패한 발송은 보낸 것으로 치지 않음)
 */
export function pendingMail(i: Inspection, log: MailLogEntry[]): MailKind | null {
  const kind = kindForStatus(i);
  if (!kind) return null;
  const since = [...i.history].reverse().find((h) => h.field === 'status' && h.to === i.status)?.at ?? i.createdAt ?? '';
  const done = log.some((m) => m.inspectionId === i.id && m.kind === kind && m.result !== 'failed' && m.at >= since);
  return done ? null : kind;
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
