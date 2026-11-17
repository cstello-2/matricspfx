import { describe, it, expect } from 'vitest';
import { MetricsEngine } from '../../src/core/engine';
import { parseConfig } from '../../src/core/validate';
import { buildSpListUrl, normalizeSpItem, IHttp } from '../../src/core/datasources';
import { formatValue } from '../../src/core/format';
import { MetricsConfig } from '../../src/core/types';

const projects = [
  { Id: 1, Title: 'Apollo', Budget: 1000 },
  { Id: 2, Title: 'Gemini', Budget: 500 },
  { Id: 3, Title: 'Mercury', Budget: 200 },
];
const tasks = [
  { Id: 10, ProjectId: 1, Hours: 5, Done: true, Created: '2026-01-05' },
  { Id: 11, ProjectId: 1, Hours: 3, Done: false, Created: '2026-01-20' },
  { Id: 12, ProjectId: 2, Hours: 8, Done: true, Created: '2026-02-02' },
  { Id: 13, ProjectId: 99, Hours: 1, Done: true, Created: '2026-02-10' },
];

const noHttp: IHttp = { getJson: () => Promise.reject(new Error('no http')), postJson: () => Promise.reject(new Error('no http')), getBinary: () => Promise.reject(new Error('no http')) };
const ctx = { http: noHttp, webUrl: 'https://contoso.sharepoint.com/sites/pmo' };

function engine(datasets: MetricsConfig['datasets']): MetricsEngine {
  return new MetricsEngine({
    dataSources: { projects: { type: 'inline', rows: projects }, tasks: { type: 'inline', rows: tasks } },
    datasets,
    widgets: [],
  }, ctx);
}

describe('engine', () => {
  it('left joins, groups and aggregates', async () => {
    const rows = await engine({
      byProject: {
        from: 'tasks',
        join: [{ source: 'projects', on: ['ProjectId', 'Id'], as: 'project' }],
        groupBy: [{ field: "project.Title ?? '(none)'", as: 'project' }],
        aggregate: { hours: 'sum(Hours)', done: 'countIf(Done)' },
        sort: '-hours',
      },
    }).resolve('byProject');
    expect(rows).toEqual([
      { project: 'Apollo', hours: 8, done: 1 },
      { project: 'Gemini', hours: 8, done: 1 },
      { project: '(none)', hours: 1, done: 1 },
    ]);
  });

  it('inner joins drop unmatched rows', async () => {
    const rows = await engine({ t: { from: 'tasks', join: [{ source: 'projects', on: ['ProjectId', 'Id'], type: 'inner' }] } }).resolve('t');
    expect(rows).toHaveLength(3);
    expect((rows[0].projects as { Title: string }).Title).toBe('Apollo');
  });

  it('one-to-many joins nest arrays usable in compute', async () => {
    const rows = await engine({
      p: {
        from: 'projects',
        join: [{ source: 'tasks', on: ['Id', 'ProjectId'], as: 'tasks', many: true }],
        compute: { hours: 'sum(tasks.Hours)', burn: 'round(sum(tasks.Hours) * 100 / Budget, 2)' },
        sort: 'Title',
      },
    }).resolve('p');
    expect(rows.map(r => [r.Title, r.hours, r.burn])).toEqual([['Apollo', 8, 0.8], ['Gemini', 8, 1.6], ['Mercury', 0, 0]]);
  });

  it('buckets dates, filters, having, limit, chained datasets', async () => {
    const e = engine({
      monthly: { from: 'tasks', filter: 'Done', groupBy: [{ field: 'Created', bucket: 'month', as: 'month' }], aggregate: { hours: 'sum(Hours)' }, sort: 'month' },
      top: { from: 'monthly', having: 'hours > 5', limit: 1 },
    });
    expect(await e.resolve('monthly')).toEqual([{ month: '2026-01', hours: 5 }, { month: '2026-02', hours: 9 }]);
    expect(await e.resolve('top')).toEqual([{ month: '2026-02', hours: 9 }]);
  });

  it('aggregate without groupBy yields one row even when empty', async () => {
    expect(await engine({ x: { from: 'tasks', filter: 'Hours > 100', aggregate: { n: 'count()', h: 'sum(Hours)' } } }).resolve('x')).toEqual([{ n: 0, h: 0 }]);
  });

  it('resolves inline dataset definitions and unions', async () => {
    const rows = await engine({}).resolve({ from: ['projects', 'projects'], aggregate: { n: 'count()' } });
    expect(rows).toEqual([{ n: 6 }]);
  });

  it('detects cycles', async () => {
    await expect(engine({ a: { from: 'b' }, b: { from: 'a' } }).resolve('a')).rejects.toThrow(/cycle/);
  });
});

describe('SharePoint list source', () => {
  it('builds REST urls with expand derived from select', () => {
    const url = buildSpListUrl({ type: 'spList', list: "Bob's Tasks", select: ['Id', 'Project/Title', 'AssignedTo/EMail'], filter: "Status eq 'Open'", top: 200 }, ctx);
    expect(decodeURIComponent(url)).toBe(
      "https://contoso.sharepoint.com/sites/pmo/_api/web/lists/getbytitle('Bob''s Tasks')/items?$select=Id,Project/Title,AssignedTo/EMail&$expand=Project,AssignedTo&$filter=Status eq 'Open'&$top=200",
    );
    expect(buildSpListUrl({ type: 'spList', listId: '{ABC}', webUrl: '/sites/other' }, ctx)).toBe(
      "https://contoso.sharepoint.com/sites/other/_api/web/lists(guid'ABC')/items?$top=1000",
    );
  });

  it('pages through odata.nextLink and normalizes items', async () => {
    const pages: Record<string, unknown> = {
      first: { value: [{ 'odata.type': 'x', Id: 1, Tags: { results: ['a'] } }], 'odata.nextLink': 'second' },
      second: { value: [{ Id: 2, Project: { Title: 'P' } }] },
    };
    let n = 0;
    const http: IHttp = { getJson: url => Promise.resolve(pages[n++ === 0 ? 'first' : url]), postJson: () => Promise.reject(new Error('read only')), getBinary: () => Promise.reject(new Error('read only')) };
    const e = new MetricsEngine({ dataSources: { l: { type: 'spList', list: 'L' } }, widgets: [] }, { http, webUrl: 'https://x' });
    expect(await e.resolve('l')).toEqual([{ Id: 1, Tags: ['a'] }, { Id: 2, Project: { Title: 'P' } }]);
  });

  it('normalizes and renames', () => {
    expect(normalizeSpItem({ __metadata: {}, Due_x0020_Date: '2026', Author: { __deferred: {} } }, { Due_x0020_Date: 'Due' })).toEqual({ Due: '2026' });
  });
});

describe('validation', () => {
  it('reports JSON syntax errors with offsets', () => {
    const r = parseConfig('{ "widgets": [ }');
    expect(r.config).toBeUndefined();
    expect(r.diagnostics[0].from).toBeGreaterThan(0);
  });

  it('accepts comments and reports semantic problems at the right place', () => {
    const text = `{
      // my dashboard
      "dataSources": { "t": { "type": "inline", "rows": [] } },
      "widgets": [
        { "type": "kpi", "dataset": "nope", "value": "sum(Hours" },
        { "type": "chart", "chart": "bar", "dataset": "t", "x": "a", "y": "b" },
      ]
    }`;
    const r = parseConfig(text);
    expect(r.config).toBeDefined();
    const msgs = r.diagnostics.map(d => d.message);
    expect(msgs.some(m => /Unknown dataset or data source "nope"/.test(m))).toBe(true);
    expect(msgs.some(m => /Expected '\)'/.test(m))).toBe(true);
    const d = r.diagnostics.find(x => /nope/.test(x.message));
    expect(d && text.slice(d.from, d.to)).toBe('"nope"');
  });
});

describe('format', () => {
  it('formats numbers, currency, percent, compact, dates', () => {
    expect(formatValue(1234.5, undefined, 'en-US')).toBe('1,235');
    expect(formatValue(0.256, 'percent', 'en-US')).toBe('25.6%');
    expect(formatValue(1500000, 'compact', 'en-US')).toBe('1.5M');
    expect(formatValue(12.5, { type: 'currency', currency: 'EUR' }, 'en-US')).toBe('€12.50');
    expect(formatValue('2026-03', undefined, 'en-US')).toBe('Mar 2026');
    expect(formatValue(null)).toBe('–');
  });
});
