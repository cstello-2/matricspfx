import { Node, parse, ExpressionError } from './parser';
import { DateBucket, Row } from '../types';
import { toDate, bucketKey, addDays, addMonths, startOf, formatDatePattern } from '../dates';

export interface Scope {
  row: Row;
  vars?: Record<string, unknown>;
  rows?: Row[];
}

export function isNil(v: unknown): v is null | undefined {
  return v === null || v === undefined;
}

export function toNumber(v: unknown): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v instanceof Date) return v.getTime();
  if (typeof v === 'string' && v.trim() !== '') {
    const n = Number(v);
    return isNaN(n) ? NaN : n;
  }
  return NaN;
}

function truthy(v: unknown): boolean {
  if (Array.isArray(v)) return v.length > 0;
  return !!v;
}

export function getProp(obj: unknown, key: unknown): unknown {
  if (isNil(obj)) return undefined;
  if (Array.isArray(obj)) {
    if (typeof key === 'number') return obj[key < 0 ? obj.length + key : key];
    if (key === 'length') return obj.length;
    const out: unknown[] = [];
    for (const item of obj) {
      const v = getProp(item, key);
      if (Array.isArray(v)) out.push(...v);
      else if (!isNil(v)) out.push(v);
    }
    return out;
  }
  if (typeof obj !== 'object') return undefined;
  const o = obj as Record<string, unknown>;
  const k = String(key);
  if (k in o) return o[k];
  const lk = k.toLowerCase();
  for (const name of Object.keys(o)) {
    if (name.toLowerCase() === lk) return o[name];
  }
  return undefined;
}

export function getPath(row: unknown, path: string): unknown {
  if (isNil(row)) return undefined;
  const direct = (row as Row)[path];
  if (direct !== undefined || path.indexOf('.') < 0) return direct !== undefined ? direct : getProp(row, path);
  let cur: unknown = row;
  for (const part of path.split('.')) {
    cur = getProp(cur, part);
    if (isNil(cur)) return cur;
  }
  return cur;
}

function looseEq(a: unknown, b: unknown): boolean {
  if (isNil(a) || isNil(b)) return isNil(a) && isNil(b);
  if (a instanceof Date || b instanceof Date) {
    const da = toDate(a), db = toDate(b);
    return !!da && !!db && da.getTime() === db.getTime();
  }
  if (typeof a === typeof b) return a === b;
  if (typeof a === 'number' || typeof b === 'number') return toNumber(a) === toNumber(b);
  return String(a) === String(b);
}

export function compare(a: unknown, b: unknown): number {
  if (isNil(a) && isNil(b)) return 0;
  if (isNil(a)) return -1;
  if (isNil(b)) return 1;
  if (a instanceof Date || b instanceof Date) {
    const da = toDate(a), db = toDate(b);
    if (da && db) return da.getTime() - db.getTime();
  }
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  const na = toNumber(a), nb = toNumber(b);
  if (typeof a !== 'string' || typeof b !== 'string') {
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
  }
  return String(a).localeCompare(String(b), undefined, { numeric: true });
}

function flatValues(values: unknown[]): unknown[] {
  const out: unknown[] = [];
  for (const v of values) {
    if (Array.isArray(v)) out.push(...v);
    else out.push(v);
  }
  return out;
}

function nums(values: unknown[]): number[] {
  const out: number[] = [];
  for (const v of flatValues(values)) {
    if (isNil(v) || v === '') continue;
    const n = toNumber(v);
    if (!isNaN(n)) out.push(n);
  }
  return out;
}

type AggFn = (values: unknown[], extra: unknown[]) => unknown;

function percentile(values: number[], p: number): number | null {
  if (!values.length) return null;
  const s = values.slice().sort((x, y) => x - y);
  const idx = (s.length - 1) * Math.min(1, Math.max(0, p));
  const lo = Math.floor(idx), hi = Math.ceil(idx);
  return s[lo] + (s[hi] - s[lo]) * (idx - lo);
}

const AGGREGATES: Record<string, AggFn> = {
  sum: v => nums(v).reduce((a, b) => a + b, 0),
  avg: v => { const n = nums(v); return n.length ? n.reduce((a, b) => a + b, 0) / n.length : null; },
  min: v => { const f = flatValues(v).filter(x => !isNil(x) && x !== ''); return f.length ? f.reduce((a, b) => (compare(a, b) <= 0 ? a : b)) : null; },
  max: v => { const f = flatValues(v).filter(x => !isNil(x) && x !== ''); return f.length ? f.reduce((a, b) => (compare(a, b) >= 0 ? a : b)) : null; },
  count: v => flatValues(v).filter(x => !isNil(x)).length,
  countIf: v => v.filter(truthy).length,
  distinct: v => { const s = new Set<string>(); flatValues(v).forEach(x => { if (!isNil(x)) s.add(x instanceof Date ? x.toISOString() : String(x)); }); return s.size; },
  median: v => percentile(nums(v), 0.5),
  percentile: (v, extra) => percentile(nums(v), toNumber(extra[0]) > 1 ? toNumber(extra[0]) / 100 : toNumber(extra[0])),
  stdev: v => {
    const n = nums(v);
    if (n.length < 2) return null;
    const m = n.reduce((a, b) => a + b, 0) / n.length;
    return Math.sqrt(n.reduce((a, b) => a + (b - m) * (b - m), 0) / (n.length - 1));
  },
  first: v => (v.length ? v[0] : null),
  last: v => (v.length ? v[v.length - 1] : null),
  values: v => flatValues(v).filter(x => !isNil(x)),
};
AGGREGATES.mean = AGGREGATES.avg;
AGGREGATES.average = AGGREGATES.avg;
AGGREGATES.countDistinct = AGGREGATES.distinct;

export const AGGREGATE_NAMES = Object.keys(AGGREGATES);

const AGGREGATE_INDEX = new Map(AGGREGATE_NAMES.map((k): [string, AggFn] => [k.toLowerCase(), AGGREGATES[k]]));

type Fn = (args: unknown[], scope: Scope) => unknown;

const BUCKETS: DateBucket[] = ['day', 'week', 'month', 'quarter', 'year'];

function toBucket(v: unknown): DateBucket {
  const b = BUCKETS.find(x => x === String(v).toLowerCase());
  if (!b) throw new Error(`Unknown date unit '${String(v)}', use one of ${BUCKETS.join(', ')}`);
  return b;
}

const str = (v: unknown): string => (isNil(v) ? '' : v instanceof Date ? v.toISOString() : String(v));

function round(v: unknown, d: unknown): number | null {
  const n = toNumber(v);
  if (isNaN(n)) return null;
  const f = Math.pow(10, isNil(d) ? 0 : toNumber(d));
  return Math.round(n * f) / f;
}

const FUNCTIONS: Record<string, Fn> = {
  if: ([c, a, b]) => (truthy(c) ? a : b),
  iif: ([c, a, b]) => (truthy(c) ? a : b),
  coalesce: args => { for (const a of args) if (!isNil(a) && a !== '') return a; return null; },
  isEmpty: ([a]) => isNil(a) || a === '' || (Array.isArray(a) && a.length === 0),
  isNull: ([a]) => isNil(a),
  round: ([v, d]) => round(v, d),
  floor: ([v]) => Math.floor(toNumber(v)),
  ceil: ([v]) => Math.ceil(toNumber(v)),
  abs: ([v]) => Math.abs(toNumber(v)),
  sqrt: ([v]) => Math.sqrt(toNumber(v)),
  pow: ([a, b]) => Math.pow(toNumber(a), toNumber(b)),
  number: ([v]) => { const n = toNumber(v); return isNaN(n) ? null : n; },
  int: ([v]) => { const n = toNumber(v); return isNaN(n) ? null : Math.trunc(n); },
  clamp: ([v, lo, hi]) => Math.min(toNumber(hi), Math.max(toNumber(lo), toNumber(v))),
  safeDiv: ([a, b, fallback]) => { const d = toNumber(b); return d === 0 || isNaN(d) ? (fallback === undefined ? null : fallback) : toNumber(a) / d; },
  pct: ([a, b]) => { const d = toNumber(b); return d === 0 || isNaN(d) ? null : toNumber(a) / d; },
  string: ([v]) => str(v),
  lower: ([v]) => str(v).toLowerCase(),
  upper: ([v]) => str(v).toUpperCase(),
  trim: ([v]) => str(v).trim(),
  len: ([v]) => (Array.isArray(v) ? v.length : str(v).length),
  concat: args => args.map(str).join(''),
  contains: ([h, n]) => (Array.isArray(h) ? h.some(x => looseEq(x, n)) : str(h).toLowerCase().indexOf(str(n).toLowerCase()) >= 0),
  startsWith: ([s, p]) => str(s).toLowerCase().indexOf(str(p).toLowerCase()) === 0,
  endsWith: ([s, p]) => { const a = str(s).toLowerCase(), b = str(p).toLowerCase(); return a.length >= b.length && a.slice(a.length - b.length) === b; },
  replace: ([s, a, b]) => str(s).split(str(a)).join(str(b)),
  left: ([s, n]) => str(s).slice(0, toNumber(n)),
  right: ([s, n]) => { const x = str(s); return x.slice(Math.max(0, x.length - toNumber(n))); },
  split: ([s, sep]) => str(s).split(isNil(sep) ? ',' : str(sep)).map(x => x.trim()),
  join: ([a, sep]) => (Array.isArray(a) ? a.map(str).join(isNil(sep) ? ', ' : str(sep)) : str(a)),
  now: () => new Date(),
  today: () => startOf(new Date(), 'day'),
  date: ([v, m, d]) => (m === undefined ? toDate(v) : new Date(toNumber(v), toNumber(m) - 1, isNil(d) ? 1 : toNumber(d))),
  year: ([v]) => { const d = toDate(v); return d ? d.getFullYear() : null; },
  month: ([v]) => { const d = toDate(v); return d ? d.getMonth() + 1 : null; },
  day: ([v]) => { const d = toDate(v); return d ? d.getDate() : null; },
  weekday: ([v]) => { const d = toDate(v); return d ? d.getDay() : null; },
  quarter: ([v]) => { const d = toDate(v); return d ? Math.floor(d.getMonth() / 3) + 1 : null; },
  addDays: ([v, n]) => { const d = toDate(v); return d ? addDays(d, toNumber(n)) : null; },
  addMonths: ([v, n]) => { const d = toDate(v); return d ? addMonths(d, toNumber(n)) : null; },
  daysBetween: ([a, b]) => {
    const da = toDate(a), db = toDate(b);
    return da && db ? Math.round((startOf(db, 'day').getTime() - startOf(da, 'day').getTime()) / 86400000) : null;
  },
  startOf: ([v, unit]) => { const d = toDate(v); return d ? startOf(d, toBucket(unit)) : null; },
  bucket: ([v, unit]) => { const d = toDate(v); return d ? bucketKey(d, toBucket(unit)) : null; },
  formatDate: ([v, p]) => { const d = toDate(v); return d ? formatDatePattern(d, isNil(p) ? 'yyyy-MM-dd' : str(p)) : ''; },
};

export const FUNCTION_NAMES = Object.keys(FUNCTIONS);

const FUNCTION_INDEX = new Map(FUNCTION_NAMES.map((k): [string, Fn] => [k.toLowerCase(), FUNCTIONS[k]]));

export function isKnownFunction(name: string): boolean {
  const n = name.toLowerCase();
  return FUNCTION_INDEX.has(n) || AGGREGATE_INDEX.has(n);
}

function lookupId(name: string, scope: Scope): unknown {
  if (name.charAt(0) === '$') {
    const v = scope.vars ? getProp(scope.vars, name.slice(1)) : undefined;
    return v;
  }
  const row = scope.row;
  if (row && name in row) return row[name];
  const v = getProp(row, name);
  if (v !== undefined) return v;
  return scope.vars ? scope.vars[name] : undefined;
}

function evalNode(n: Node, scope: Scope): unknown {
  switch (n.k) {
    case 'lit': return n.v;
    case 'id': return lookupId(n.name, scope);
    case 'arr': return n.items.map(i => evalNode(i, scope));
    case 'get': return getProp(evalNode(n.obj, scope), evalNode(n.key, scope));
    case 'un': {
      const a = evalNode(n.a, scope);
      return n.op === '!' ? !truthy(a) : -toNumber(a);
    }
    case 'cond': return truthy(evalNode(n.c, scope)) ? evalNode(n.a, scope) : evalNode(n.b, scope);
    case 'bin': return evalBin(n.op, n.a, n.b, scope);
    case 'call': return evalCall(n, scope);
  }
}

function evalBin(op: string, an: Node, bn: Node, scope: Scope): unknown {
  if (op === '&&') { const a = evalNode(an, scope); return truthy(a) ? evalNode(bn, scope) : a; }
  if (op === '||') { const a = evalNode(an, scope); return truthy(a) ? a : evalNode(bn, scope); }
  if (op === '??') { const a = evalNode(an, scope); return isNil(a) ? evalNode(bn, scope) : a; }
  const a = evalNode(an, scope);
  const b = evalNode(bn, scope);
  switch (op) {
    case '+':
      if (typeof a === 'string' || typeof b === 'string') return str(a) + str(b);
      if (a instanceof Date) return addDays(a, toNumber(b));
      return toNumber(a) + toNumber(b);
    case '-':
      if (a instanceof Date && b instanceof Date) return (a.getTime() - b.getTime()) / 86400000;
      if (a instanceof Date) return addDays(a, -toNumber(b));
      return toNumber(a) - toNumber(b);
    case '*': return toNumber(a) * toNumber(b);
    case '/': { const d = toNumber(b); return d === 0 ? null : toNumber(a) / d; }
    case '%': return toNumber(a) % toNumber(b);
    case '==': return looseEq(a, b);
    case '!=': return !looseEq(a, b);
    case '<': return !isNil(a) && !isNil(b) && compare(a, b) < 0;
    case '<=': return !isNil(a) && !isNil(b) && compare(a, b) <= 0;
    case '>': return !isNil(a) && !isNil(b) && compare(a, b) > 0;
    case '>=': return !isNil(a) && !isNil(b) && compare(a, b) >= 0;
    case 'in':
      if (Array.isArray(b)) return b.some(x => looseEq(a, x));
      if (typeof b === 'string') return b.toLowerCase().indexOf(str(a).toLowerCase()) >= 0;
      return false;
    default: throw new Error(`Unknown operator ${op}`);
  }
}

function evalCall(n: Extract<Node, { k: 'call' }>, scope: Scope): unknown {
  const name = n.name.toLowerCase();
  const agg = AGGREGATE_INDEX.get(name);
  const fn = FUNCTION_INDEX.get(name);

  if (agg) {
    if ((name === 'min' || name === 'max') && n.args.length > 1) {
      return agg(n.args.map(a => evalNode(a, scope)), []);
    }
    const extra = n.args.slice(1).map(a => evalNode(a, scope));
    if (scope.rows) {
      const rowScope = (r: Row): Scope => ({ row: r, vars: scope.vars });
      if (n.args.length === 0) return name === 'count' ? scope.rows.length : agg(scope.rows, extra);
      const values = scope.rows.map(r => evalNode(n.args[0], rowScope(r)));
      return agg(values, extra);
    }
    if (n.args.length === 0) return name === 'count' ? 1 : null;
    const v = evalNode(n.args[0], scope);
    return agg(Array.isArray(v) ? v : [v], extra);
  }
  if (fn) return fn(n.args.map(a => evalNode(a, scope)), scope);
  throw new Error(`Unknown function '${n.name}'`);
}

export function evaluate(expression: string, scope: Scope): unknown {
  const node = parse(expression);
  try {
    return evalNode(node, scope);
  } catch (e) {
    if (e instanceof ExpressionError) throw e;
    throw new ExpressionError((e as Error).message, expression);
  }
}

export function evaluateAggregate(expression: string, rows: Row[], vars?: Record<string, unknown>): unknown {
  return evaluate(expression, { row: rows[0] || {}, rows, vars });
}

const PLAIN_PATH = /^[A-Za-z_@][A-Za-z0-9_@]*(\.[A-Za-z_@][A-Za-z0-9_@]*)*$/;

export function fieldOrExpr(expression: string, row: Row, vars?: Record<string, unknown>): unknown {
  if (PLAIN_PATH.test(expression) && !/^(true|false|null)$/i.test(expression)) return getPath(row, expression);
  return evaluate(expression, { row, vars });
}
