/**
 * 엑셀 셀 값을 프로그램 값으로 바꿉니다.
 * 업체마다 입력 방식이 달라 날짜·시간·숫자를 여러 형태로 받아 줍니다.
 */

type Cell = unknown;

const pad = (n: number) => String(n).padStart(2, '0');

/** 셀 값을 화면에 보일 문자열로 (서식·수식·링크 처리) */
export function cellText(v: Cell): string {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    // 시간만 있는 셀(1899-12-30 기준)은 HH:MM
    if (v.getUTCFullYear() < 1901) return `${pad(v.getUTCHours())}:${pad(v.getUTCMinutes())}`;
    const base = `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())}`;
    return v.getUTCHours() || v.getUTCMinutes() ? `${base} ${pad(v.getUTCHours())}:${pad(v.getUTCMinutes())}` : base;
  }
  if (typeof v === 'object') {
    const o = v as any;
    if (Array.isArray(o.richText)) return o.richText.map((r: any) => r.text ?? '').join('').trim();
    if ('result' in o) return cellText(o.result);
    if ('text' in o) return cellText(o.text);
    if ('hyperlink' in o) return String(o.hyperlink).replace(/^mailto:/i, '').trim();
    if ('error' in o) return String(o.error);
    return '';
  }
  return String(v).trim();
}

/** 수식 셀이면 결과값 */
export function cellValue(v: Cell): Cell {
  if (v && typeof v === 'object' && !(v instanceof Date)) {
    const o = v as any;
    if ('result' in o) return o.result;
    if (Array.isArray(o.richText) || 'text' in o) return cellText(v);
  }
  return v;
}

function validYmd(y: number, m: number, d: number): string | null {
  if (y < 1900 || y > 2999 || m < 1 || m > 12 || d < 1) return null;
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (dt.getUTCMonth() !== m - 1) return null;
  return `${y}-${pad(m)}-${pad(d)}`;
}

/** 엑셀 날짜 일련번호 → YYYY-MM-DD */
function serialToYmd(n: number): string | null {
  if (!Number.isFinite(n) || n < 1 || n > 2958465) return null;
  const ms = Math.round((n - 25569) * 86400000);
  const d = new Date(ms);
  return validYmd(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
}

/** 날짜. 비었으면 '' / 알아볼 수 없으면 null */
export function parseDate(raw: Cell): string | '' | null {
  const v = cellValue(raw);
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) return validYmd(v.getUTCFullYear(), v.getUTCMonth() + 1, v.getUTCDate());
  if (typeof v === 'number') return serialToYmd(v);
  const s = String(v).trim().replace(/\s+/g, '');
  if (!s) return '';
  let m = s.match(/^(\d{4})[-./년](\d{1,2})[-./월](\d{1,2})[일.]?$/);
  if (m) return validYmd(+m[1], +m[2], +m[3]);
  m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m) return validYmd(+m[1], +m[2], +m[3]);
  return null;
}

/** 시간. 비었으면 '' / 알아볼 수 없으면 null */
export function parseTime(raw: Cell): string | '' | null {
  const v = cellValue(raw);
  if (v === null || v === undefined || v === '') return '';
  const hm = (h: number, mi: number) => (h >= 0 && h < 24 && mi >= 0 && mi < 60 ? `${pad(h)}:${pad(mi)}` : null);
  if (v instanceof Date) return hm(v.getUTCHours(), v.getUTCMinutes());
  if (typeof v === 'number') {
    if (v >= 0 && v < 1) {
      const total = Math.round(v * 24 * 60);
      return hm(Math.floor(total / 60), total % 60);
    }
    return null;
  }
  const s = String(v).trim().replace(/\s+/g, '');
  if (!s) return '';
  let m = s.match(/^(\d{1,2})[:시](\d{2})분?$/);
  if (m) return hm(+m[1], +m[2]);
  m = s.match(/^(오전|오후)(\d{1,2})(?:[:시](\d{2})분?|시)?$/);
  if (m) {
    let h = +m[2];
    if (h > 12) return null;
    if (m[1] === '오후' && h < 12) h += 12;
    if (m[1] === '오전' && h === 12) h = 0;
    return hm(h, m[3] ? +m[3] : 0);
  }
  m = s.match(/^(\d{1,2})시$/);
  if (m) return hm(+m[1], 0);
  return null;
}

/** 수량: 1 이상 정수. 비었으면 '' / 잘못되면 null */
export function parseQuantity(raw: Cell): number | '' | null {
  const v = cellValue(raw);
  if (v === null || v === undefined || v === '') return '';
  let n: number;
  if (typeof v === 'number') n = v;
  else {
    const s = String(v).trim().replace(/,/g, '').replace(/(개|ea|EA|pcs|PCS|set|SET)$/, '').trim();
    if (!s) return '';
    if (!/^\d+(\.0+)?$/.test(s)) return null;
    n = Number(s);
  }
  return Number.isInteger(n) && n >= 1 ? n : null;
}

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
export function isEmail(s: string): boolean {
  return EMAIL_RE.test(s);
}

/** 머리글 비교용: 공백·기호 제거, 소문자 */
export function normHeader(s: string): string {
  return s.replace(/[\s/·._\-()*]/g, '').toLowerCase();
}
