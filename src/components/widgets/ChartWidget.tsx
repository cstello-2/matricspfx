import * as React from 'react';
import * as echarts from 'echarts/core';
import { BarChart, LineChart, PieChart, ScatterChart, GaugeChart, FunnelChart, HeatmapChart, SankeyChart, TreemapChart, RadarChart } from 'echarts/charts';
import { GridComponent, TooltipComponent, LegendComponent, VisualMapComponent, TitleComponent, RadarComponent } from 'echarts/components';
import { SVGRenderer } from 'echarts/renderers';
import { ChartWidgetConfig, Row } from '../../core/types';
import { useDashboard, useRows } from '../context';
import { ErrorBox, Loading, WidgetFrame } from '../WidgetFrame';
import { buildChartOption } from '../chartOptions';
import { DataTable } from './TableWidget';

echarts.use([
  BarChart, LineChart, PieChart, ScatterChart, GaugeChart, FunnelChart, HeatmapChart, SankeyChart, TreemapChart, RadarChart,
  GridComponent, TooltipComponent, LegendComponent, VisualMapComponent, TitleComponent, RadarComponent,
  SVGRenderer,
]);

function EChart({ w, rows, height }: { w: ChartWidgetConfig; rows: Row[]; height: number }): React.ReactElement {
  const { theme, locale } = useDashboard();
  const ref = React.useRef<HTMLDivElement>(null);
  const chart = React.useRef<echarts.ECharts | undefined>(undefined);
  const [error, setError] = React.useState<Error | undefined>(undefined);
  const [width, setWidth] = React.useState(0);

  React.useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const instance = echarts.init(el, undefined, { renderer: 'svg' });
    chart.current = instance;
    const measure = (): void => setWidth(Math.round(el.clientWidth / 40) * 40);
    measure();
    const observer = new ResizeObserver(() => {
      instance.resize();
      measure();
    });
    observer.observe(el);
    return () => {
      observer.disconnect();
      instance.dispose();
      chart.current = undefined;
    };
  }, []);

  React.useEffect(() => {
    if (!chart.current || !width) return;
    try {
      chart.current.setOption(buildChartOption(w, rows, theme, locale, width), true);
      setError(undefined);
    } catch (e) {
      setError(e as Error);
    }
  }, [w, rows, theme, locale, width]);

  return (
    <>
      {error && <ErrorBox error={error} />}
      <div ref={ref} style={{ width: '100%', height }} role="img" aria-label={w.title || `${w.chart} chart`} />
    </>
  );
}

function tableColumns(w: ChartWidgetConfig): string[] {
  const cols: string[] = [];
  const add = (f?: string): void => { if (f && cols.indexOf(f) < 0) cols.push(f); };
  add(w.x);
  add(w.series);
  add(w.from);
  add(w.to);
  const ys = w.y ? (Array.isArray(w.y) ? w.y : [w.y]) : [];
  ys.forEach(y => add(typeof y === 'string' ? y : y.field));
  add(w.value && /^[\w.]+$/.test(w.value) ? w.value : undefined);
  return cols;
}

export function ChartWidget({ w }: { w: ChartWidgetConfig }): React.ReactElement {
  const { styles } = useDashboard();
  const { rows, error } = useRows(w.dataset);
  const [asTable, setAsTable] = React.useState(false);
  const height = w.height || (w.chart === 'gauge' ? 200 : 280);

  const toggle = (
    <button type="button" className={styles.iconButton} onClick={() => setAsTable(!asTable)} aria-pressed={asTable} title={asTable ? 'Show chart' : 'Show data as table'}>
      {asTable ? 'Chart' : 'Table'}
    </button>
  );

  let body: React.ReactNode;
  if (error) body = <ErrorBox error={error} />;
  else if (!rows) body = <Loading height={height} />;
  else if (!rows.length && w.chart !== 'gauge') body = <div className={styles.empty}>No data</div>;
  else if (asTable) {
    const cols = tableColumns(w);
    body = <DataTable rows={rows} columns={cols.length ? cols.map(c => ({ field: c, format: c === cols[0] ? undefined : w.format })) : undefined} pageSize={10} maxHeight={height} />;
  } else body = <EChart w={w} rows={rows} height={height} />;

  return <WidgetFrame title={w.title} subtitle={w.subtitle} actions={w.chart !== 'gauge' ? toggle : undefined}>{body}</WidgetFrame>;
}
