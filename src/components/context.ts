import * as React from 'react';
import { MetricsEngine } from '../core/engine';
import { DatasetRef, Row } from '../core/types';
import { MetricsTheme } from './theme';
import { Styles } from './styles';

export interface DashboardContextValue {
  engine: MetricsEngine;
  theme: MetricsTheme;
  styles: Styles;
  locale?: string;
}

export const DashboardContext = React.createContext<DashboardContextValue | undefined>(undefined);

export function useDashboard(): DashboardContextValue {
  const c = React.useContext(DashboardContext);
  if (!c) throw new Error('useDashboard must be used inside <MetricsDashboard>');
  return c;
}

export interface AsyncRows {
  rows?: Row[];
  error?: Error;
  loading: boolean;
}

export function useRows(ref: DatasetRef | undefined): AsyncRows {
  const { engine } = useDashboard();
  const [state, setState] = React.useState<AsyncRows>({ loading: !!ref });
  const key = typeof ref === 'string' ? ref : JSON.stringify(ref);
  React.useEffect(() => {
    if (!ref) { setState({ loading: false, rows: [] }); return; }
    let live = true;
    setState(s => ({ rows: s.rows, loading: true }));
    engine.resolve(ref).then(
      rows => { if (live) setState({ rows, loading: false }); },
      (error: Error) => { if (live) setState({ error, loading: false }); },
    );
    return () => { live = false; };
  }, [engine, key]);
  return state;
}
