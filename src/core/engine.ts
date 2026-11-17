import { DatasetConfig, DatasetRef, GroupByConfig, JoinConfig, MetricsConfig, Row } from './types';
import { DataContext, loadDataSource } from './datasources';
import { compare, evaluate, evaluateAggregate, fieldOrExpr, isNil } from './expr/evaluate';
import { bucketKey, toDate } from './dates';

export class MetricsEngine {
  private readonly sources = new Map<string, Promise<Row[]>>();
  private readonly datasets = new Map<string, Promise<Row[]>>();
  public readonly vars: Record<string, unknown>;

  public constructor(private readonly config: MetricsConfig, private readonly ctx: DataContext) {
    this.vars = {
      user: ctx.user ? { name: ctx.user.displayName, email: ctx.user.email, login: ctx.user.loginName } : {},
      today: new Date(new Date().setHours(0, 0, 0, 0)),
      now: new Date(),
      webUrl: ctx.webUrl,
    };
  }

  public resolve(ref: DatasetRef): Promise<Row[]> {
    if (typeof ref === 'string') return this.byName(ref, []);
    return this.run(ref, []);
  }

  private byName(name: string, stack: string[]): Promise<Row[]> {
    const ds = this.config.datasets && this.config.datasets[name];
    if (ds) {
      if (stack.indexOf(name) >= 0) return Promise.reject(new Error(`Dataset cycle: ${stack.concat(name).join(' → ')}`));
      let p = this.datasets.get(name);
      if (!p) {
        p = this.run(ds, stack.concat(name)).catch(e => { throw new Error(`Dataset "${name}": ${(e as Error).message}`); });
        this.datasets.set(name, p);
      }
      return p;
    }
    const src = this.config.dataSources && this.config.dataSources[name];
    if (src) {
      let p = this.sources.get(name);
      if (!p) {
        p = loadDataSource(name, src, this.ctx);
        this.sources.set(name, p);
      }
      return p;
    }
    return Promise.reject(new Error(`Unknown dataset or data source "${name}"`));
  }

  private async run(ds: DatasetConfig, stack: string[]): Promise<Row[]> {
    const froms = Array.isArray(ds.from) ? ds.from : [ds.from];
    if (!froms.length || !froms[0]) throw new Error('"from" is required');
    const bases = await Promise.all(froms.map(f => this.byName(f, stack)));
    let rows: Row[] = [];
    for (const b of bases) for (const r of b) rows.push({ ...r });

    if (ds.join) {
      for (const j of ds.join) {
        const right = await this.byName(j.source, stack);
        rows = joinRows(rows, right, j);
      }
    }
    return transform(rows, ds, this.vars);
  }
}

function keyOf(v: unknown): string {
  if (v instanceof Date) return v.toISOString();
  if (v && typeof v === 'object') {
    const o = v as Record<string, unknown>;
    if ('Id' in o) return String(o.Id);
    if ('ID' in o) return String(o.ID);
  }
  return String(v);
}

export function joinRows(left: Row[], right: Row[], j: JoinConfig): Row[] {
  const [lf, rf] = j.on;
  const as = j.as || j.source;
  const index = new Map<string, Row[]>();
  for (const r of right) {
    const k = fieldOrExpr(rf, r);
    if (isNil(k)) continue;
    const key = keyOf(k);
    const list = index.get(key);
    if (list) list.push(r);
    else index.set(key, [r]);
  }
  const out: Row[] = [];
  for (const l of left) {
    const lv = fieldOrExpr(lf, l);
    const keys = Array.isArray(lv) ? lv : [lv];
    let matches: Row[] = [];
    for (const k of keys) {
      if (isNil(k)) continue;
      const m = index.get(keyOf(k));
      if (m) matches = matches.concat(m);
    }
    if (j.many) {
      if (!matches.length && j.type === 'inner') continue;
      out.push({ ...l, [as]: matches });
    } else if (!matches.length) {
      if (j.type !== 'inner') out.push({ ...l, [as]: null });
    } else {
      for (const m of matches) out.push({ ...l, [as]: m });
    }
  }
  return out;
}

function groupSpec(g: string | GroupByConfig): GroupByConfig {
  return typeof g === 'string' ? { field: g } : g;
}

export function groupRows(rows: Row[], groupBy: Array<string | GroupByConfig>, aggregate: Record<string, string>, vars: Record<string, unknown>): Row[] {
  const specs = groupBy.map(groupSpec);
  const groups = new Map<string, { key: Row; rows: Row[] }>();
  for (const r of rows) {
    const key: Row = {};
    const parts: string[] = [];
    for (const g of specs) {
      let v = fieldOrExpr(g.field, r, vars);
      if (g.bucket) {
        const d = toDate(v);
        v = d ? bucketKey(d, g.bucket) : null;
      }
      key[g.as || g.field] = v === undefined ? null : v;
      parts.push(keyOf(v));
    }
    const k = parts.join('\u0001');
    const grp = groups.get(k);
    if (grp) grp.rows.push(r);
    else groups.set(k, { key, rows: [r] });
  }
  if (!specs.length && !groups.size) groups.set('', { key: {}, rows: [] });
  const out: Row[] = [];
  groups.forEach(({ key, rows: grows }) => {
    const row: Row = { ...key };
    for (const name of Object.keys(aggregate)) {
      row[name] = evaluateAggregate(aggregate[name], grows, { ...vars, ...row });
    }
    out.push(row);
  });
  return out;
}

export function parseSort(sort: string | string[]): Array<{ field: string; desc: boolean }> {
  const items = Array.isArray(sort) ? sort : [sort];
  return items.map(s => {
    let field = s.trim();
    let desc = false;
    if (field.charAt(0) === '-') { desc = true; field = field.slice(1); }
    const m = /^(.*?)\s+(asc|desc)$/i.exec(field);
    if (m) { field = m[1]; desc = m[2].toLowerCase() === 'desc'; }
    return { field, desc };
  });
}

export function sortRows(rows: Row[], sort: string | string[]): Row[] {
  const keys = parseSort(sort);
  return rows.slice().sort((a, b) => {
    for (const k of keys) {
      const c = compare(fieldOrExpr(k.field, a), fieldOrExpr(k.field, b));
      if (c !== 0) return k.desc ? -c : c;
    }
    return 0;
  });
}

export function transform(input: Row[], ds: Omit<DatasetConfig, 'from' | 'join'>, vars: Record<string, unknown>): Row[] {
  let rows = input;
  if (ds.filter) {
    const f = ds.filter;
    rows = rows.filter(r => !!evaluate(f, { row: r, vars }));
  }
  const compute = ds.compute;
  if (compute) {
    const names = Object.keys(compute);
    rows = rows.map(r => {
      const o: Row = { ...r };
      for (const n of names) o[n] = evaluate(compute[n], { row: o, vars });
      return o;
    });
  }
  if (ds.groupBy || ds.aggregate) {
    const agg = ds.aggregate && Object.keys(ds.aggregate).length ? ds.aggregate : { count: 'count()' };
    rows = groupRows(rows, ds.groupBy || [], agg, vars);
  }
  if (ds.having) {
    const h = ds.having;
    rows = rows.filter(r => !!evaluate(h, { row: r, vars }));
  }
  if (ds.sort && (Array.isArray(ds.sort) ? ds.sort.length : ds.sort)) rows = sortRows(rows, ds.sort);
  if (ds.limit !== undefined && ds.limit >= 0) rows = rows.slice(0, ds.limit);
  return rows;
}
