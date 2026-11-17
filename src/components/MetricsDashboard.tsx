import * as React from 'react';
import { MetricsConfig, WidgetConfig } from '../core/types';
import { MetricsEngine } from '../core/engine';
import { DataContext } from '../core/datasources';
import { DashboardContext, DashboardContextValue } from './context';
import { MetricsTheme, withPalette } from './theme';
import { getStyles } from './styles';
import { ErrorBox, WidgetErrorBoundary, WidgetFrame } from './WidgetFrame';
import { KpiWidget } from './widgets/KpiWidget';
import { ChartWidget } from './widgets/ChartWidget';
import { TableWidget } from './widgets/TableWidget';
import { TextWidget } from './widgets/TextWidget';

export interface MetricsDashboardProps {
  config?: MetricsConfig;
  error?: string;
  data: DataContext;
  theme: MetricsTheme;
  locale?: string;
  reloadToken?: number;
  placeholder?: React.ReactNode;
  hideHeader?: boolean;
}

function defaultSpan(w: WidgetConfig, columns: number): number {
  if (w.span) return w.span;
  if (w.type === 'kpi') return Math.round(columns / 4);
  if (w.type === 'table') return columns;
  return Math.round(columns / 2);
}

function effectiveSpan(span: number, columns: number, width: number): number {
  if (width > 0 && width < 560) return columns;
  if (width > 0 && width < 900) return Math.min(columns, span * 2);
  return Math.min(columns, span);
}

function Widget({ w }: { w: WidgetConfig }): React.ReactElement {
  switch (w.type) {
    case 'kpi': return <KpiWidget w={w} />;
    case 'chart': return <ChartWidget w={w} />;
    case 'table': return <TableWidget w={w} />;
    case 'text': return <TextWidget w={w} />;
    default: return <WidgetFrame title="Unknown widget"><ErrorBox error={`Unknown widget type "${(w as { type: string }).type}"`} /></WidgetFrame>;
  }
}

export function MetricsDashboard(props: MetricsDashboardProps): React.ReactElement {
  const { config } = props;
  const rootRef = React.useRef<HTMLDivElement>(null);
  const [width, setWidth] = React.useState(0);
  const [tick, setTick] = React.useState(0);

  React.useEffect(() => {
    const el = rootRef.current;
    if (!el) return;
    setWidth(el.clientWidth);
    const observer = new ResizeObserver(entries => setWidth(entries[0].contentRect.width));
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

  const refresh = config && config.refreshSeconds ? Math.max(15, config.refreshSeconds) : 0;
  React.useEffect(() => {
    if (!refresh) return;
    const h = setInterval(() => setTick(t => t + 1), refresh * 1000);
    return () => clearInterval(h);
  }, [refresh]);

  const theme = React.useMemo(() => withPalette(props.theme, config && config.theme && config.theme.palette), [props.theme, config]);
  const styles = getStyles(theme, config && config.theme && config.theme.cardStyle);

  const [engine, setEngine] = React.useState<MetricsEngine | undefined>(undefined);
  React.useEffect(() => {
    setEngine(config ? new MetricsEngine(config, props.data) : undefined);
  }, [config, props.data, props.reloadToken, tick]);

  const ctxValue = React.useMemo<DashboardContextValue | undefined>(
    () => (engine ? { engine, theme, styles, locale: (config && config.locale) || props.locale } : undefined),
    [engine, theme, styles, config, props.locale],
  );

  const columns = (config && config.layout && config.layout.columns) || 12;
  const gap = config && config.layout && config.layout.gap !== undefined ? config.layout.gap : 16;

  let content: React.ReactNode;
  if (props.error) {
    content = <div className={styles.error} role="alert">{props.error}</div>;
  } else if (!config || !ctxValue) {
    content = config ? undefined : props.placeholder;
  } else {
    content = (
      <DashboardContext.Provider value={ctxValue}>
        {!props.hideHeader && (config.title || config.description) && (
          <header className={styles.header}>
            {config.title && <h2 className={styles.title}>{config.title}</h2>}
            {config.description && <p className={styles.description}>{config.description}</p>}
          </header>
        )}
        <div className={styles.grid} style={{ gridTemplateColumns: `repeat(${columns}, minmax(0, 1fr))`, gap }}>
          {(config.widgets || []).map((w, i) => (
            <div key={i} className={styles.cell} style={{ gridColumn: `span ${effectiveSpan(defaultSpan(w, columns), columns, width)}` }}>
              <WidgetErrorBoundary resetKey={JSON.stringify(w)} render={e => <WidgetFrame title={w.title}><ErrorBox error={e} /></WidgetFrame>}>
                <Widget w={w} />
              </WidgetErrorBoundary>
            </div>
          ))}
        </div>
      </DashboardContext.Provider>
    );
  }

  return <div ref={rootRef} className={styles.root}>{content}</div>;
}
