// 날짜 도우미 (YYYY-MM-DD 문자열 기준, 현지 시간)

const pad = (n: number) => String(n).padStart(2, '0');
export const WEEKDAYS = ['일', '월', '화', '수', '목', '금', '토'];

export const ymd = (d: Date) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

export function parseYmd(s: string): Date {
  const [y, m, d] = s.split('-').map(Number);
  return new Date(y, m - 1, d);
}

export const addDays = (s: string, n: number) => {
  const d = parseYmd(s);
  d.setDate(d.getDate() + n);
  return ymd(d);
};

/** 그 주의 월요일 */
export function weekStart(s: string): string {
  const d = parseYmd(s);
  const wd = (d.getDay() + 6) % 7; // 월=0
  return addDays(s, -wd);
}

export const monthStart = (s: string) => s.slice(0, 8) + '01';
export function monthEnd(s: string): string {
  const d = parseYmd(monthStart(s));
  d.setMonth(d.getMonth() + 1);
  d.setDate(0);
  return ymd(d);
}
export function addMonths(s: string, n: number): string {
  const d = parseYmd(monthStart(s));
  d.setMonth(d.getMonth() + n);
  return ymd(d);
}

/** 11-10(화) */
export function fmtMD(s: string | null | undefined): string {
  if (!s) return '';
  const d = parseYmd(s);
  return `${pad(d.getMonth() + 1)}-${pad(d.getDate())}(${WEEKDAYS[d.getDay()]})`;
}

/** 2026-11-10(화) */
export function fmtYMD(s: string | null | undefined): string {
  if (!s) return '';
  return `${s}(${WEEKDAYS[parseYmd(s).getDay()]})`;
}

/** ISO 시각 → 11-09 14:05 */
export function fmtStamp(iso: string | null | undefined): string {
  if (!iso) return '';
  const d = new Date(iso);
  return `${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/** '9:30', '0930', '13시 30분' → '09:30' (24시간). 알아볼 수 없으면 null */
export function normalizeTime(raw: string): string | null {
  const t = raw.trim().replace(/\s+/g, '');
  const m = t.match(/^(\d{1,2})[:시]?(\d{2})분?$/);
  if (!m) return null;
  const h = Number(m[1]), mi = Number(m[2]);
  return h < 24 && mi < 60 ? `${String(h).padStart(2, '0')}:${String(mi).padStart(2, '0')}` : null;
}
