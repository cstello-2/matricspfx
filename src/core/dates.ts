import { DateBucket } from './types';

const ISO_DATE = /^\d{4}-\d{2}-\d{2}([T ][\d:.]+(Z|[+-]\d{2}:?\d{2})?)?$/;

export function toDate(v: unknown): Date | null {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v === 'number') return new Date(v);
  if (typeof v === 'string' && v) {
    if (/^\d{4}-\d{2}-\d{2}$/.test(v)) {
      const [y, m, d] = v.split('-').map(Number);
      return new Date(y, m - 1, d);
    }
    if (/^\d{4}-\d{2}$/.test(v)) {
      const [y, m] = v.split('-').map(Number);
      return new Date(y, m - 1, 1);
    }
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

export function looksLikeDate(v: unknown): boolean {
  return v instanceof Date || (typeof v === 'string' && ISO_DATE.test(v));
}

export function addDays(d: Date, n: number): Date {
  const r = new Date(d.getTime());
  r.setDate(r.getDate() + n);
  return r;
}

export function addMonths(d: Date, n: number): Date {
  const r = new Date(d.getTime());
  const day = r.getDate();
  r.setDate(1);
  r.setMonth(r.getMonth() + n);
  const last = new Date(r.getFullYear(), r.getMonth() + 1, 0).getDate();
  r.setDate(Math.min(day, last));
  return r;
}

export function startOf(d: Date, unit: DateBucket): Date {
  switch (unit) {
    case 'year': return new Date(d.getFullYear(), 0, 1);
    case 'quarter': return new Date(d.getFullYear(), Math.floor(d.getMonth() / 3) * 3, 1);
    case 'month': return new Date(d.getFullYear(), d.getMonth(), 1);
    case 'week': {
      const day = (d.getDay() + 6) % 7;
      const s = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      s.setDate(s.getDate() - day);
      return s;
    }
    default: return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
}

const pad = (n: number, w = 2): string => {
  let s = String(n);
  while (s.length < w) s = '0' + s;
  return s;
};

export function bucketKey(d: Date, unit: DateBucket): string {
  const s = startOf(d, unit);
  switch (unit) {
    case 'year': return String(s.getFullYear());
    case 'quarter': return `${s.getFullYear()}-Q${Math.floor(s.getMonth() / 3) + 1}`;
    case 'month': return `${s.getFullYear()}-${pad(s.getMonth() + 1)}`;
    default: return `${s.getFullYear()}-${pad(s.getMonth() + 1)}-${pad(s.getDate())}`;
  }
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

export function formatDatePattern(d: Date, pattern: string, locale?: string): string {
  return pattern.replace(/yyyy|yy|MMMM|MMM|MM|M|dd|d|HH|mm/g, t => {
    switch (t) {
      case 'yyyy': return String(d.getFullYear());
      case 'yy': return pad(d.getFullYear() % 100);
      case 'MMMM': return d.toLocaleDateString(locale, { month: 'long' });
      case 'MMM': return locale ? d.toLocaleDateString(locale, { month: 'short' }) : MONTHS[d.getMonth()];
      case 'MM': return pad(d.getMonth() + 1);
      case 'M': return String(d.getMonth() + 1);
      case 'dd': return pad(d.getDate());
      case 'd': return String(d.getDate());
      case 'HH': return pad(d.getHours());
      default: return pad(d.getMinutes());
    }
  });
}

export function bucketLabel(key: string, locale?: string): string {
  let m = /^(\d{4})-(\d{2})$/.exec(key);
  if (m) return new Date(+m[1], +m[2] - 1, 1).toLocaleDateString(locale, { month: 'short', year: 'numeric' });
  m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(key);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]).toLocaleDateString(locale, { month: 'short', day: 'numeric' });
  return key;
}
