import { DataSourceConfig, InlineSourceConfig, RestSourceConfig, Row, SpListSourceConfig } from '../types';
import { getPath } from '../expr/evaluate';

export interface HttpOptions {
  sharepoint: boolean;
  headers?: Record<string, string>;
}

export interface HttpPostOptions extends HttpOptions {
  body?: unknown;
}

export interface IHttp {
  getJson(url: string, options: HttpOptions): Promise<unknown>;
  postJson(url: string, options: HttpPostOptions): Promise<unknown>;
  getBinary(url: string, options: HttpOptions): Promise<ArrayBuffer>;
}

export interface DataContext {
  http: IHttp;
  webUrl: string;
  user?: { displayName: string; email: string; loginName?: string };
}

export class DataSourceError extends Error {
  public constructor(public readonly source: string, message: string) {
    super(`Data source "${source}": ${message}`);
    this.name = 'DataSourceError';
  }
}

function odataString(s: string): string {
  return s.replace(/'/g, "''");
}

export function resolveWebUrl(webUrl: string | undefined, current: string): string {
  if (!webUrl) return current.replace(/\/$/, '');
  if (/^https?:\/\//i.test(webUrl)) return webUrl.replace(/\/$/, '');
  const m = /^(https?:\/\/[^/]+)/i.exec(current);
  const origin = m ? m[1] : '';
  return (origin + (webUrl.charAt(0) === '/' ? '' : '/') + webUrl).replace(/\/$/, '');
}

export function listApiUrl(cfg: SpListSourceConfig, webUrl: string): string {
  const web = resolveWebUrl(cfg.webUrl, webUrl);
  if (cfg.listId) return `${web}/_api/web/lists(guid'${cfg.listId.replace(/[{}]/g, '')}')`;
  if (cfg.list) return `${web}/_api/web/lists/getbytitle('${encodeURIComponent(odataString(cfg.list))}')`;
  throw new Error('either "list" or "listId" is required');
}

export function expandFor(select: string[], expand: string[] = []): string[] {
  const out = expand.slice();
  for (const s of select) {
    const slash = s.indexOf('/');
    if (slash > 0 && out.indexOf(s.slice(0, slash)) < 0) out.push(s.slice(0, slash));
  }
  return out;
}

export function buildSpListUrl(cfg: SpListSourceConfig, ctx: DataContext): string {

  const params: string[] = [];
  const select = cfg.select && cfg.select.length ? cfg.select.slice() : [];
  const expand = expandFor(select, cfg.expand);
  if (select.length) params.push('$select=' + encodeURIComponent(select.join(',')));
  if (expand.length) params.push('$expand=' + encodeURIComponent(expand.join(',')));
  if (cfg.filter) params.push('$filter=' + encodeURIComponent(cfg.filter));
  if (cfg.orderBy) params.push('$orderby=' + encodeURIComponent(cfg.orderBy));
  const top = Math.max(1, Math.min(5000, cfg.pageSize || 1000, cfg.top || 5000));
  params.push('$top=' + top);
  return `${listApiUrl(cfg, ctx.webUrl)}/items?${params.join('&')}`;
}

export function normalizeSpItem(item: Record<string, unknown>, rename?: Record<string, string>): Row {
  const out: Row = {};
  for (const key of Object.keys(item)) {
    if (key.indexOf('odata.') === 0 || key.indexOf('@odata') === 0 || key === '__metadata') continue;
    let v = item[key];
    if (v && typeof v === 'object' && !Array.isArray(v)) {
      const o = v as Record<string, unknown>;
      if (Array.isArray(o.results)) v = (o.results as unknown[]).map(x => (x && typeof x === 'object' ? normalizeSpItem(x as Record<string, unknown>) : x));
      else if ('__deferred' in o) continue;
      else v = normalizeSpItem(o);
    } else if (Array.isArray(v)) {
      v = v.map(x => (x && typeof x === 'object' ? normalizeSpItem(x as Record<string, unknown>) : x));
    }
    out[rename && rename[key] ? rename[key] : key] = v;
  }
  return out;
}

export async function loadSpList(name: string, cfg: SpListSourceConfig, ctx: DataContext): Promise<Row[]> {
  const max = cfg.top || 5000;
  let url: string | undefined = buildSpListUrl(cfg, ctx);
  const rows: Row[] = [];
  let guard = 0;
  while (url && rows.length < max && guard++ < 200) {
    const res = (await ctx.http.getJson(url, { sharepoint: true, headers: { Accept: 'application/json;odata=nometadata' } })) as Record<string, unknown>;
    const d = (res && (res.d as Record<string, unknown>)) || res;
    const items = (d && ((d.value as unknown[]) || (d.results as unknown[]))) || [];
    if (!Array.isArray(items)) throw new DataSourceError(name, 'unexpected response shape');
    for (const it of items) rows.push(normalizeSpItem(it as Record<string, unknown>, cfg.rename));
    url = (res['odata.nextLink'] || res['@odata.nextLink'] || (d && d.__next)) as string | undefined;
  }
  return rows.slice(0, max);
}

function loadInline(cfg: InlineSourceConfig): Row[] {
  return Array.isArray(cfg.rows) ? cfg.rows.map(r => ({ ...r })) : [];
}

async function loadRest(name: string, cfg: RestSourceConfig, ctx: DataContext): Promise<Row[]> {
  const url = /^https?:\/\//i.test(cfg.url) ? cfg.url : resolveWebUrl(cfg.url, ctx.webUrl);
  const sharepoint = cfg.sharepoint !== undefined ? cfg.sharepoint : url.indexOf('/_api/') >= 0;
  const res = await ctx.http.getJson(url, { sharepoint, headers: cfg.headers });
  let data: unknown = cfg.path ? getPath(res, cfg.path) : res;
  if (!cfg.path && data && !Array.isArray(data) && typeof data === 'object') {
    const o = data as Record<string, unknown>;
    data = o.value || (o.d && ((o.d as Record<string, unknown>).results || o.d)) || o;
  }
  if (!Array.isArray(data)) {
    if (data && typeof data === 'object') return [data as Row];
    throw new DataSourceError(name, `response${cfg.path ? ` at "${cfg.path}"` : ''} is not an array`);
  }
  return data.map(x => (x && typeof x === 'object' ? (x as Row) : { value: x }));
}

export async function loadDataSource(name: string, cfg: DataSourceConfig, ctx: DataContext): Promise<Row[]> {
  try {
    switch (cfg.type) {
      case 'spList': return await loadSpList(name, cfg, ctx);
      case 'inline': return loadInline(cfg);
      case 'rest': return await loadRest(name, cfg, ctx);
      default: throw new Error(`unknown type "${(cfg as { type: string }).type}"`);
    }
  } catch (e) {
    if (e instanceof DataSourceError) throw e;
    throw new DataSourceError(name, (e as Error).message);
  }
}
