import { Format, FormatName, FormatSpec } from './types';
import { toNumber, isNil } from './expr/evaluate';
import { toDate, looksLikeDate, bucketLabel } from './dates';

function spec(f?: Format): FormatSpec {
  if (!f) return {};
  return typeof f === 'string' ? { type: f } : f;
}

export function formatValue(v: unknown, f?: Format, locale?: string): string {
  if (isNil(v) || (typeof v === 'number' && isNaN(v))) return '–';
  const s = spec(f);
  const pre = s.prefix || '';
  const suf = s.suffix || '';
  const isBucket = typeof v === 'string' && /^\d{4}-(\d{2}|Q\d)$/.test(v);
  const type = s.type || (typeof v === 'number' ? 'number' : looksLikeDate(v) || isBucket ? 'date' : 'text');

  if (type === 'text') return pre + (Array.isArray(v) ? v.map(x => formatValue(x, undefined, locale)).join(', ') : objectText(v)) + suf;

  if (type === 'date' || type === 'datetime') {
    if (typeof v === 'string' && /^\d{4}(-Q\d|-\d{2})?$/.test(v)) return pre + bucketLabel(v, locale) + suf;
    const d = toDate(v);
    if (!d) return pre + String(v) + suf;
    const opts: Intl.DateTimeFormatOptions = type === 'date'
      ? { year: 'numeric', month: 'short', day: 'numeric' }
      : { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' };
    return pre + d.toLocaleString(locale, opts) + suf;
  }

  const n = toNumber(v);
  if (isNaN(n)) return pre + String(v) + suf;
  return pre + new Intl.NumberFormat(locale, numberOptions(type, s, n)).format(n) + suf;
}

function numberOptions(type: FormatName, s: FormatSpec, n: number): Intl.NumberFormatOptions {
  const dec = s.decimals;
  const digits = (fallback: number): Intl.NumberFormatOptions => {
    const max = dec !== undefined ? dec : fallback;
    return { minimumFractionDigits: dec !== undefined ? dec : 0, maximumFractionDigits: max };
  };
  switch (type) {
    case 'integer':
      return { maximumFractionDigits: 0 };
    case 'decimal':
      return { minimumFractionDigits: dec !== undefined ? dec : 2, maximumFractionDigits: dec !== undefined ? dec : 2 };
    case 'percent':
      return { style: 'percent', ...digits(1) };
    case 'currency': {
      const fraction = dec !== undefined ? dec : Math.abs(n) >= 1000 ? 0 : 2;
      return { style: 'currency', currency: s.currency || 'USD', minimumFractionDigits: fraction, maximumFractionDigits: fraction };
    }
    case 'compact':
      return { notation: 'compact', ...digits(1) };
    default:
      return digits(Math.abs(n) >= 100 ? 0 : 2);
  }
}

export function formatAxis(n: number, f?: Format, locale?: string): string {
  const s = spec(f);
  if (s.type === 'percent') return formatValue(n, { type: 'percent', decimals: 0 }, locale);
  const compact = Math.abs(n) >= 10000;
  const options: Intl.NumberFormatOptions = {
    notation: compact ? 'compact' : 'standard',
    maximumFractionDigits: compact ? 1 : 2,
  };
  if (s.type === 'currency') {
    options.style = 'currency';
    options.currency = s.currency || 'USD';
    options.minimumFractionDigits = 0;
    if (!compact) options.maximumFractionDigits = 0;
  }
  return new Intl.NumberFormat(locale, options).format(n);
}

function objectText(v: unknown): string {
  if (v && typeof v === 'object') {
    const o = v as Record<string, unknown>;
    if ('Title' in o) return String(o.Title);
    if ('title' in o) return String(o.title);
    if ('Url' in o) return String(o.Description || o.Url);
    return JSON.stringify(v);
  }
  return String(v);
}
