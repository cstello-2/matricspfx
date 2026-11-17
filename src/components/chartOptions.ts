import type {
  EChartsOption,
  BarSeriesOption,
  LineSeriesOption,
  PieSeriesOption,
  ScatterSeriesOption,
  GaugeSeriesOption,
  FunnelSeriesOption,
  HeatmapSeriesOption,
  SankeySeriesOption,
  TreemapSeriesOption,
  RadarSeriesOption,
  XAXisComponentOption,
  YAXisComponentOption,
  TooltipComponentOption,
  LegendComponentOption,
  DefaultLabelFormatterCallbackParams,
  TooltipComponentFormatterCallbackParams,
} from 'echarts';
import { ChartWidgetConfig, Row, SeriesConfig } from '../core/types';
import { fieldOrExpr, evaluateAggregate, toNumber, compare } from '../core/expr/evaluate';
import { formatValue, formatAxis } from '../core/format';
import { bucketLabel } from '../core/dates';
import { MetricsTheme } from './theme';

type CartesianAxis = XAXisComponentOption & YAXisComponentOption;

interface BuildContext {
  w: ChartWidgetConfig;
  rows: Row[];
  t: MetricsTheme;
  locale?: string;
  width: number;
}

interface NameValue {
  name: string;
  value: number;
}

interface BuiltSeries {
  name: string;
  spec: SeriesConfig;
  data: Array<number | null>;
}

const MAX_SLICES = 8;

function seriesSpecs(y: ChartWidgetConfig['y']): SeriesConfig[] {
  if (!y) return [];
  const arr = Array.isArray(y) ? y : [y];
  return arr.map(s => (typeof s === 'string' ? { field: s } : s));
}

function label(v: unknown, locale?: string): string {
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) return bucketLabel(v, locale);
  return formatValue(v, undefined, locale);
}

function uniq(values: unknown[]): unknown[] {
  const seen = new Set<string>();
  const out: unknown[] = [];
  for (const v of values) {
    const k = String(v);
    if (!seen.has(k)) {
      seen.add(k);
      out.push(v);
    }
  }
  return out;
}

function num(v: unknown): number | null {
  const n = toNumber(v);
  return isNaN(n) ? null : n;
}

function paletteOf(c: BuildContext): string[] {
  return c.w.colors && c.w.colors.length ? c.w.colors : c.t.palette;
}

function isPlainObject(v: unknown): v is Record<string, unknown> {
  return !!v && typeof v === 'object' && !Array.isArray(v);
}

export function deepMerge<T extends object>(target: T, src: unknown): T {
  if (!isPlainObject(src)) return target;
  const out: Record<string, unknown> = { ...(target as Record<string, unknown>) };
  for (const k of Object.keys(src)) {
    const sv = src[k];
    const tv = out[k];
    out[k] = isPlainObject(sv) && isPlainObject(tv) ? deepMerge(tv, sv) : sv;
  }
  return out as T;
}

function tooltip(t: MetricsTheme, extra: TooltipComponentOption): TooltipComponentOption {
  return {
    backgroundColor: t.surface,
    borderColor: t.border,
    borderWidth: 1,
    textStyle: { color: t.text, fontSize: 12, fontFamily: t.fontFamily },
    extraCssText: 'box-shadow: 0 4px 12px rgba(0,0,0,0.12); border-radius: 8px;',
    confine: true,
    ...extra,
  };
}

function baseOption(t: MetricsTheme): EChartsOption {
  return {
    backgroundColor: 'transparent',
    textStyle: { fontFamily: t.fontFamily, color: t.textSecondary },
    animationDuration: 400,
  };
}

function legend(t: MetricsTheme, show: boolean, extra: LegendComponentOption = {}): LegendComponentOption {
  return {
    show,
    type: 'scroll',
    top: 0,
    left: 0,
    icon: 'roundRect',
    itemWidth: 10,
    itemHeight: 10,
    itemGap: 16,
    textStyle: { color: t.textSecondary, fontSize: 12 },
    pageTextStyle: { color: t.textMuted },
    ...extra,
  };
}

function valueAxis(c: BuildContext): CartesianAxis {
  return {
    type: 'value',
    axisLine: { show: false },
    axisTick: { show: false },
    axisLabel: { color: c.t.textMuted, fontSize: 11, formatter: (v: number) => formatAxis(v, c.w.format, c.locale) },
    splitLine: { show: true, lineStyle: { color: c.t.grid, width: 1, type: 'solid' } },
  };
}

function itemTooltip(c: BuildContext, render: (p: DefaultLabelFormatterCallbackParams) => string): TooltipComponentOption {
  return tooltip(c.t, {
    trigger: 'item',
    formatter: (params: TooltipComponentFormatterCallbackParams) => render(Array.isArray(params) ? params[0] : params),
  });
}

export function buildChartOption(w: ChartWidgetConfig, rows: Row[], t: MetricsTheme, locale?: string, width: number = 600): EChartsOption {
  const c: BuildContext = { w, rows, t, locale, width };
  let opt: EChartsOption;
  switch (w.chart) {
    case 'pie':
    case 'donut':
      opt = pieOption(c);
      break;
    case 'funnel':
      opt = funnelOption(c);
      break;
    case 'gauge':
      opt = gaugeOption(c);
      break;
    case 'heatmap':
      opt = heatmapOption(c);
      break;
    case 'sankey':
      opt = sankeyOption(c);
      break;
    case 'treemap':
      opt = treemapOption(c);
      break;
    case 'radar':
      opt = radarOption(c);
      break;
    case 'scatter':
      opt = scatterOption(c);
      break;
    default:
      opt = cartesianOption(c);
  }
  return w.echarts ? deepMerge(opt, w.echarts) : opt;
}

function categoriesOf(c: BuildContext): unknown[] {
  const x = c.w.x;
  return x ? uniq(c.rows.map(r => fieldOrExpr(x, r))) : c.rows.map((_, i) => i + 1);
}

function buildSeries(c: BuildContext, categories: unknown[]): BuiltSeries[] {
  const { w, rows } = c;
  const specs = seriesSpecs(w.y);
  const catIndex = new Map<string, number>();
  categories.forEach((cat, i) => catIndex.set(String(cat), i));
  const xOf = (r: Row, i: number): unknown => (w.x ? fieldOrExpr(w.x, r) : i + 1);
  const accumulate = (spec: SeriesConfig, include: (r: Row) => boolean): Array<number | null> => {
    const data: Array<number | null> = categories.map(() => null);
    rows.forEach((r, i) => {
      if (!include(r)) return;
      const idx = catIndex.get(String(xOf(r, i)));
      const v = num(fieldOrExpr(spec.field, r));
      if (idx !== undefined && v !== null) data[idx] = (data[idx] || 0) + v;
    });
    return data;
  };

  const seriesField = w.series;
  if (seriesField && specs.length) {
    const spec = specs[0];
    const names = uniq(rows.map(r => fieldOrExpr(seriesField, r)));
    return names.map(n => ({
      name: label(n, c.locale),
      spec: { ...spec, name: undefined },
      data: accumulate(spec, r => String(fieldOrExpr(seriesField, r)) === String(n)),
    }));
  }
  return specs.map(spec => ({ name: spec.name || spec.field, spec, data: accumulate(spec, () => true) }));
}

function lastIndex(data: Array<number | null>): number {
  for (let i = data.length - 1; i >= 0; i--) if (data[i] !== null) return i;
  return -1;
}

function cartesianOption(c: BuildContext): EChartsOption {
  const { w, t } = c;
  const horizontal = w.chart === 'bar';
  const categories = categoriesOf(c);
  const series = buildSeries(c, categories);
  const multi = series.length > 1;
  const showLegend = w.showLegend !== undefined ? w.showLegend : multi;
  const labelsMode = w.labels === undefined ? 'auto' : w.labels;
  const palette = paletteOf(c);
  const fmt = (v: unknown): string => formatValue(v, w.format, c.locale);

  const echSeries: Array<BarSeriesOption | LineSeriesOption> = series.map((s, si) => {
    const kind = s.spec.type || (w.chart === 'line' ? 'line' : w.chart === 'area' ? 'area' : 'bar');
    const color = s.spec.color || palette[si % palette.length];
    const stack = s.spec.stack || (w.stacked ? 'total' : undefined);
    if (kind === 'bar') {
      const lastInStack = !stack || si === series.length - 1;
      const radius: number | number[] = stack && !lastInStack ? 0 : horizontal ? [0, 4, 4, 0] : [4, 4, 0, 0];
      const bar: BarSeriesOption = {
        type: 'bar',
        name: s.name,
        data: s.data,
        stack,
        barMaxWidth: 24,
        barGap: '12%',
        itemStyle: { color, borderRadius: radius, borderColor: stack ? t.surface : undefined, borderWidth: stack ? 1 : 0 },
        emphasis: { focus: 'series' },
        label: {
          show: labelsMode === true || (labelsMode === 'auto' && !multi && categories.length <= 12),
          position: horizontal ? 'right' : 'top',
          color: t.textSecondary,
          fontSize: 11,
          formatter: (p: DefaultLabelFormatterCallbackParams) => fmt(p.value),
        },
      };
      return bar;
    }
    const last = lastIndex(s.data);
    const line: LineSeriesOption = {
      type: 'line',
      name: s.name,
      data: s.data,
      stack,
      smooth: !!w.smooth,
      connectNulls: true,
      symbol: 'circle',
      symbolSize: 8,
      showSymbol: categories.length <= 16,
      lineStyle: { width: 2, color, cap: 'round', join: 'round' },
      itemStyle: { color, borderColor: t.surface, borderWidth: 2 },
      areaStyle: kind === 'area' ? { color, opacity: 0.1 } : undefined,
      emphasis: { focus: 'series' },
      label: {
        show: labelsMode !== false && (labelsMode === true || series.length <= 4),
        position: 'right',
        color: t.textSecondary,
        fontSize: 11,
        formatter: (p: DefaultLabelFormatterCallbackParams) => (labelsMode === true || p.dataIndex === last ? fmt(p.value) : ''),
      },
    };
    return line;
  });

  const hasBars = echSeries.some(s => s.type === 'bar');
  const hasLines = echSeries.some(s => s.type === 'line');
  const categoryAxis: CartesianAxis = {
    type: 'category',
    data: categories.map(cat => label(cat, c.locale)),
    boundaryGap: hasBars,
    inverse: horizontal,
    axisLine: { show: true, lineStyle: { color: t.axis } },
    axisTick: { show: false },
    splitLine: { show: false },
    axisLabel: horizontal
      ? { color: t.textMuted, fontSize: 11, width: Math.round(Math.max(56, Math.min(140, c.width * 0.28))), overflow: 'truncate', interval: 0 }
      : {
          color: t.textMuted,
          fontSize: 11,
          hideOverlap: true,
          interval: categories.length <= 12 ? 0 : 'auto',
          width: Math.max(48, Math.min(96, (c.width - 60) / Math.max(1, categories.length))),
          overflow: 'break',
          lineHeight: 14,
        },
  };
  const amountAxis = valueAxis(c);

  return {
    ...baseOption(t),
    color: palette,
    legend: legend(t, showLegend),
    grid: { left: 8, right: hasLines ? 48 : 24, top: showLegend ? 36 : 12, bottom: 8, outerBoundsMode: 'same', outerBoundsContain: 'axisLabel' },
    tooltip: tooltip(t, {
      trigger: 'axis',
      axisPointer: hasLines
        ? { type: 'line', lineStyle: { color: t.axis } }
        : { type: 'shadow', shadowStyle: { color: t.isDark ? 'rgba(255,255,255,0.04)' : 'rgba(0,0,0,0.04)' } },
      valueFormatter: v => fmt(v),
    }),
    xAxis: horizontal ? amountAxis : categoryAxis,
    yAxis: horizontal ? categoryAxis : amountAxis,
    series: echSeries,
  };
}

function scatterOption(c: BuildContext): EChartsOption {
  const { w, t, rows } = c;
  const yField = seriesSpecs(w.y)[0];
  const seriesField = w.series;
  const groups = seriesField ? uniq(rows.map(r => fieldOrExpr(seriesField, r))) : [yField ? yField.name || yField.field : 'value'];
  const palette = paletteOf(c);
  const series: ScatterSeriesOption[] = groups.map((g, gi) => ({
    type: 'scatter',
    name: label(g, c.locale),
    symbolSize: 10,
    itemStyle: { color: palette[gi % palette.length], borderColor: t.surface, borderWidth: 2, opacity: 0.9 },
    data: rows
      .filter(r => !seriesField || String(fieldOrExpr(seriesField, r)) === String(g))
      .map(r => [num(w.x ? fieldOrExpr(w.x, r) : 0), num(yField ? fieldOrExpr(yField.field, r) : 0)]),
  }));
  const showLegend = w.showLegend !== undefined ? w.showLegend : series.length > 1;
  const xAxis: CartesianAxis = {
    type: 'value',
    name: w.x,
    nameLocation: 'middle',
    nameGap: 24,
    nameTextStyle: { color: t.textMuted },
    axisLine: { show: true, lineStyle: { color: t.axis } },
    axisTick: { show: false },
    axisLabel: { color: t.textMuted, fontSize: 11 },
    splitLine: { show: false },
  };
  return {
    ...baseOption(t),
    legend: legend(t, showLegend),
    grid: { left: 8, right: 24, top: showLegend ? 36 : 12, bottom: 8, outerBoundsMode: 'same', outerBoundsContain: 'axisLabel' },
    tooltip: itemTooltip(c, p => {
      const v = Array.isArray(p.value) ? p.value : [];
      return `${p.seriesName}<br/>${w.x}: <b>${formatValue(v[0], undefined, c.locale)}</b><br/>${yField ? yField.field : ''}: <b>${formatValue(v[1], w.format, c.locale)}</b>`;
    }),
    xAxis,
    yAxis: valueAxis(c),
    series,
  };
}

function nameValues(c: BuildContext): NameValue[] {
  const { w, rows } = c;
  const spec = seriesSpecs(w.y)[0];
  const field = spec ? spec.field : w.value || 'value';
  const map = new Map<string, number>();
  rows.forEach((r, i) => {
    const n = label(w.x ? fieldOrExpr(w.x, r) : i + 1, c.locale);
    const v = num(fieldOrExpr(field, r));
    if (v !== null) map.set(n, (map.get(n) || 0) + v);
  });
  const items: NameValue[] = [];
  map.forEach((value, name) => items.push({ name, value }));
  return items;
}

function foldOther(items: NameValue[], max: number): NameValue[] {
  const sorted = items.slice().sort((a, b) => b.value - a.value);
  if (sorted.length <= max) return sorted;
  const head = sorted.slice(0, max - 1);
  const rest = sorted.slice(max - 1).reduce((a, b) => a + b.value, 0);
  return head.concat([{ name: 'Other', value: rest }]);
}

function pieOption(c: BuildContext): EChartsOption {
  const { w, t } = c;
  const items = foldOther(nameValues(c), MAX_SLICES);
  const donut = w.chart === 'donut';
  const total = items.reduce((a, b) => a + b.value, 0);
  const showLegend = w.showLegend !== undefined ? w.showLegend : items.length > 1;
  const showLabels = w.labels === undefined || w.labels === 'auto' ? items.length <= 6 : w.labels;
  const series: PieSeriesOption = {
    type: 'pie',
    radius: donut ? ['52%', '78%'] : ['0%', '78%'],
    center: [showLegend ? '35%' : '50%', '50%'],
    data: items,
    padAngle: 1,
    itemStyle: { borderColor: t.surface, borderWidth: 2, borderRadius: 4 },
    label: { show: showLabels && !donut, color: t.textSecondary, fontSize: 11, formatter: '{d}%' },
    labelLine: { lineStyle: { color: t.axis } },
    emphasis: { scaleSize: 4 },
  };
  return {
    ...baseOption(t),
    color: paletteOf(c),
    legend: legend(t, showLegend, { top: 'middle', left: 'auto', right: 0, orient: 'vertical' }),
    tooltip: itemTooltip(c, p => `${p.name}<br/><b>${formatValue(p.value, w.format, c.locale)}</b> (${p.percent}%)`),
    title: donut
      ? {
          text: formatValue(total, w.format || 'compact', c.locale),
          subtext: 'Total',
          left: showLegend ? '34%' : 'center',
          top: 'center',
          textAlign: 'center',
          itemGap: 2,
          textStyle: { color: t.text, fontSize: 20, fontWeight: 600 },
          subtextStyle: { color: t.textMuted, fontSize: 11 },
        }
      : undefined,
    series: [series],
  };
}

function funnelOption(c: BuildContext): EChartsOption {
  const { w, t } = c;
  const items = nameValues(c).sort((a, b) => b.value - a.value);
  const ramp = t.ordinal;
  const colors = items.map((_, i) => ramp[Math.min(ramp.length - 1, Math.round((i * (ramp.length - 1)) / Math.max(1, items.length - 1)))]);
  const top = items.length ? items[0].value : 0;
  const series: FunnelSeriesOption = {
    type: 'funnel',
    left: 8,
    right: 8,
    top: 8,
    bottom: 8,
    minSize: '20%',
    gap: 2,
    sort: 'descending',
    data: items.map((it, i) => ({ ...it, itemStyle: { color: colors[i] } })),
    itemStyle: { borderWidth: 0, borderRadius: 4 },
    label: {
      show: true,
      position: 'inside',
      color: '#ffffff',
      fontSize: 12,
      formatter: (p: DefaultLabelFormatterCallbackParams) => `${p.name}  ${formatValue(p.value, w.format, c.locale)}`,
    },
  };
  return {
    ...baseOption(t),
    tooltip: itemTooltip(c, p => {
      const v = toNumber(p.value);
      return `${p.name}<br/><b>${formatValue(v, w.format, c.locale)}</b> (${top ? Math.round((v / top) * 100) : 0}% of first stage)`;
    }),
    series: [series],
  };
}

function treemapOption(c: BuildContext): EChartsOption {
  const { w, t } = c;
  const items = foldOther(nameValues(c), 12);
  const palette = paletteOf(c);
  const series: TreemapSeriesOption = {
    type: 'treemap',
    roam: false,
    nodeClick: false,
    breadcrumb: { show: false },
    left: 0,
    right: 0,
    top: 0,
    bottom: 0,
    data: items.map((it, i) => ({ ...it, itemStyle: { color: palette[i % palette.length] } })),
    itemStyle: { borderColor: t.surface, borderWidth: 2, gapWidth: 2, borderRadius: 4 },
    label: {
      show: true,
      color: '#ffffff',
      fontSize: 12,
      formatter: (p: DefaultLabelFormatterCallbackParams) => `${p.name}\n${formatValue(p.value, w.format || 'compact', c.locale)}`,
    },
  };
  return {
    ...baseOption(t),
    tooltip: itemTooltip(c, p => `${p.name}<br/><b>${formatValue(p.value, w.format, c.locale)}</b>`),
    series: [series],
  };
}

function isPercentFormat(w: ChartWidgetConfig): boolean {
  return w.format === 'percent' || (typeof w.format === 'object' && w.format.type === 'percent');
}

function gaugeOption(c: BuildContext): EChartsOption {
  const { w, t, rows } = c;
  let value: number | null;
  if (w.value) value = num(evaluateAggregate(w.value, rows));
  else {
    const spec = seriesSpecs(w.y)[0];
    value = rows.length && spec ? num(fieldOrExpr(spec.field, rows[0])) : null;
  }
  const min = w.min !== undefined ? w.min : 0;
  const max = w.max !== undefined ? w.max : isPercentFormat(w) ? 1 : Math.max(100, Math.ceil((value || 0) * 1.25));
  const color = (w.colors && w.colors[0]) || t.accent;
  const series: GaugeSeriesOption = {
    type: 'gauge',
    min,
    max,
    startAngle: 210,
    endAngle: -30,
    radius: '92%',
    center: ['50%', '58%'],
    progress: { show: true, width: 14, roundCap: true, itemStyle: { color } },
    axisLine: { roundCap: true, lineStyle: { width: 14, color: [[1, t.grid]] } },
    pointer: { show: false },
    axisTick: { show: false },
    splitLine: { show: false },
    axisLabel: { show: false },
    anchor: { show: false },
    title: { show: false },
    detail: {
      valueAnimation: true,
      offsetCenter: [0, 0],
      fontSize: 28,
      fontWeight: 600,
      color: t.text,
      formatter: (v: number) => formatValue(v, w.format, c.locale),
    },
    data: [{ value: value === null ? 0 : value }],
  };
  return { ...baseOption(t), series: [series] };
}

function heatmapOption(c: BuildContext): EChartsOption {
  const { w, t, rows } = c;
  const xField = w.x || '';
  const yField = w.series || '';
  const valueField = w.value || '';
  const xs = uniq(rows.map(r => fieldOrExpr(xField, r))).sort(compare);
  const ys = uniq(rows.map(r => fieldOrExpr(yField, r))).sort(compare);
  const xi = new Map(xs.map((x, i): [string, number] => [String(x), i]));
  const yi = new Map(ys.map((y, i): [string, number] => [String(y), i]));
  const cells = new Map<string, [number, number, number]>();
  for (const r of rows) {
    const a = xi.get(String(fieldOrExpr(xField, r)));
    const b = yi.get(String(fieldOrExpr(yField, r)));
    const v = num(fieldOrExpr(valueField, r));
    if (a === undefined || b === undefined || v === null) continue;
    const key = `${a}|${b}`;
    const cell = cells.get(key);
    if (cell) cell[2] += v;
    else cells.set(key, [a, b, v]);
  }
  const data = Array.from(cells.values());
  const values = data.map(d => d[2]);
  const lo = values.length ? Math.min(...values) : 0;
  const hi = values.length ? Math.max(...values) : 1;
  const categoryAxis = (vals: unknown[]): CartesianAxis => ({
    type: 'category',
    data: vals.map(v => label(v, c.locale)),
    axisLine: { show: false },
    axisTick: { show: false },
    axisLabel: { color: t.textMuted, fontSize: 11 },
    splitArea: { show: false },
  });
  const series: HeatmapSeriesOption = {
    type: 'heatmap',
    data,
    itemStyle: { borderColor: t.surface, borderWidth: 2, borderRadius: 4 },
    emphasis: { itemStyle: { borderColor: t.text, borderWidth: 1 } },
  };
  return {
    ...baseOption(t),
    tooltip: itemTooltip(c, p => {
      const v = Array.isArray(p.value) ? p.value.map(Number) : [];
      return `${label(xs[v[0]], c.locale)} · ${label(ys[v[1]], c.locale)}<br/><b>${formatValue(v[2], w.format, c.locale)}</b>`;
    }),
    grid: { left: 8, right: 8, top: 8, bottom: 40, outerBoundsMode: 'same', outerBoundsContain: 'axisLabel' },
    xAxis: categoryAxis(xs),
    yAxis: categoryAxis(ys),
    visualMap: {
      type: 'continuous',
      min: lo,
      max: hi,
      calculable: false,
      orient: 'horizontal',
      left: 'center',
      bottom: 0,
      itemHeight: 120,
      itemWidth: 10,
      text: [formatAxis(hi, w.format, c.locale), formatAxis(lo, w.format, c.locale)],
      textStyle: { color: t.textMuted, fontSize: 11 },
      inRange: { color: t.sequential },
    },
    series: [series],
  };
}

function sankeyOption(c: BuildContext): EChartsOption {
  const { w, t, rows } = c;
  const fromField = w.from || '';
  const toField = w.to || '';
  const links = new Map<string, { source: string; target: string; value: number }>();
  const nodes: string[] = [];
  const addNode = (n: string): void => {
    if (nodes.indexOf(n) < 0) nodes.push(n);
  };
  for (const r of rows) {
    const source = label(fieldOrExpr(fromField, r), c.locale);
    const target = label(fieldOrExpr(toField, r), c.locale);
    if (source === target) continue;
    const v = w.value ? num(fieldOrExpr(w.value, r)) : 1;
    if (v === null || v <= 0) continue;
    addNode(source);
    addNode(target);
    const key = JSON.stringify([source, target]);
    const link = links.get(key);
    if (link) link.value += v;
    else links.set(key, { source, target, value: v });
  }
  const data = Array.from(links.values());
  const depth = new Map<string, number>(nodes.map((n): [string, number] => [n, 0]));
  for (let pass = 0; pass < nodes.length; pass++) {
    let changed = false;
    for (const l of data) {
      const d = (depth.get(l.source) || 0) + 1;
      if (d > (depth.get(l.target) || 0) && d < nodes.length) {
        depth.set(l.target, d);
        changed = true;
      }
    }
    if (!changed) break;
  }
  const palette = paletteOf(c);
  const series: SankeySeriesOption = {
    type: 'sankey',
    left: 8,
    right: 96,
    top: 8,
    bottom: 8,
    nodeWidth: 12,
    nodeGap: 10,
    draggable: false,
    emphasis: { focus: 'adjacency' },
    data: nodes.map(name => ({ name, itemStyle: { color: palette[Math.min(depth.get(name) || 0, palette.length - 1)], borderRadius: 2 } })),
    links: data,
    lineStyle: { color: 'gradient', opacity: 0.25, curveness: 0.5 },
    label: { color: t.textSecondary, fontSize: 11 },
  };
  return {
    ...baseOption(t),
    tooltip: tooltip(t, { trigger: 'item', valueFormatter: v => formatValue(v, w.format, c.locale) }),
    series: [series],
  };
}

function radarOption(c: BuildContext): EChartsOption {
  const { w, t } = c;
  const categories = categoriesOf(c);
  const series = buildSeries(c, categories);
  const palette = paletteOf(c);
  const maxV = Math.max(1, ...series.map(s => Math.max(0, ...s.data.map(v => v || 0))));
  const showLegend = w.showLegend !== undefined ? w.showLegend : series.length > 1;
  const radar: RadarSeriesOption = {
    type: 'radar',
    symbolSize: 6,
    data: series.map((s, i) => ({
      name: s.name,
      value: s.data.map(v => v || 0),
      lineStyle: { width: 2, color: palette[i % palette.length] },
      itemStyle: { color: palette[i % palette.length] },
      areaStyle: { color: palette[i % palette.length], opacity: 0.1 },
    })),
  };
  return {
    ...baseOption(t),
    legend: legend(t, showLegend),
    tooltip: tooltip(t, { trigger: 'item' }),
    radar: {
      indicator: categories.map(cat => ({ name: label(cat, c.locale), max: maxV * 1.1 })),
      radius: '68%',
      center: ['50%', showLegend ? '56%' : '50%'],
      axisName: { color: t.textMuted, fontSize: 11 },
      splitLine: { lineStyle: { color: t.grid } },
      splitArea: { show: false },
      axisLine: { lineStyle: { color: t.grid } },
    },
    series: [radar],
  };
}
