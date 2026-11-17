import { describe, it, expect } from 'vitest';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { odataFilterToExpr } from '../sp-mock/odata';
import { evaluate } from '../../src/core/expr/evaluate';
import { parseConfig } from '../../src/core/validate';
import { buildChartOption } from '../../src/components/chartOptions';
import { LIGHT_THEME } from '../../src/components/theme';
import { ChartWidgetConfig } from '../../src/core/types';

describe('OData $filter translation (workbench mock)', () => {
  const row = { Status: 'Open', Hours: 7, Title: "Bob's task", Created: '2026-03-01T10:00:00Z', Project: { Title: 'Apollo' } };
  const run = (f: string) => evaluate(odataFilterToExpr(f), { row });

  it('handles comparison and logic operators', () => {
    expect(run("Status eq 'Open' and Hours gt 5")).toBe(true);
    expect(run("Status ne 'Open' or (Hours le 7 and not (Hours lt 7))")).toBe(true);
  });
  it('handles escaped quotes, lookups, functions and datetimes', () => {
    expect(run("Title eq 'Bob''s task'")).toBe(true);
    expect(run("Project/Title eq 'Apollo'")).toBe(true);
    expect(run("substringof('task', Title)")).toBe(true);
    expect(run("startswith(Title, 'Bob')")).toBe(true);
    expect(run("Created ge datetime'2026-02-01T00:00:00Z'")).toBe(true);
    expect(run("Created lt datetime'2026-02-01T00:00:00Z'")).toBe(false);
  });
});

describe('samples', () => {
  const dir = join(__dirname, '..', '..', 'src', 'samples');
  for (const f of readdirSync(dir).filter(x => x.endsWith('.json'))) {
    it(`${f} validates without errors`, () => {
      const r = parseConfig(readFileSync(join(dir, f), 'utf8'));
      expect(r.diagnostics.filter(d => d.severity === 'error')).toEqual([]);
    });
  }
});

describe('chart options', () => {
  const rows = [
    { m: '2026-01', region: 'N', v: 1 }, { m: '2026-01', region: 'S', v: 2 },
    { m: '2026-02', region: 'N', v: 3 }, { m: '2026-02', region: 'S', v: 4 },
  ];
  it('pivots a series field into one series per value', () => {
    const w: ChartWidgetConfig = { type: 'chart', chart: 'column', dataset: 'x', x: 'm', y: 'v', series: 'region' };
    const o = buildChartOption(w, rows, LIGHT_THEME, 'en-US') as { series: Array<{ name: string; data: number[] }>; xAxis: { data: string[] }; legend: { show: boolean } };
    expect(o.series.map(s => [s.name, s.data])).toEqual([['N', [1, 3]], ['S', [2, 4]]]);
    expect(o.xAxis.data).toEqual(['Jan 2026', 'Feb 2026']);
    expect(o.legend.show).toBe(true);
  });
  it('folds pie slices beyond 8 into Other and merges the echarts escape hatch', () => {
    const many = Array.from({ length: 12 }, (_, i) => ({ k: 'c' + i, v: 12 - i }));
    const w: ChartWidgetConfig = { type: 'chart', chart: 'pie', dataset: 'x', x: 'k', y: 'v', echarts: { series: [{ radius: '50%' }], animation: false } };
    const o = buildChartOption(w, many, LIGHT_THEME) as { series: Array<{ data: Array<{ name: string; value: number }>; radius: string }>; animation: boolean };
    expect(o.series[0].radius).toBe('50%');
    expect(o.animation).toBe(false);
    const o2 = buildChartOption({ ...w, echarts: undefined }, many, LIGHT_THEME) as { series: Array<{ data: Array<{ name: string; value: number }> }> };
    expect(o2.series[0].data).toHaveLength(8);
    expect(o2.series[0].data[7]).toEqual({ name: 'Other', value: 15 });
  });
});
