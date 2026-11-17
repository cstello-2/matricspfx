import * as React from 'react';
import { useDashboard } from './context';

export interface WidgetFrameProps {
  title?: string;
  subtitle?: string;
  actions?: React.ReactNode;
  style?: React.CSSProperties;
  children?: React.ReactNode;
}

export function WidgetFrame(props: WidgetFrameProps): React.ReactElement {
  const { styles } = useDashboard();
  const hasHead = !!(props.title || props.subtitle || props.actions);
  return (
    <section className={styles.card} style={props.style} aria-label={props.title}>
      {hasHead && (
        <div className={styles.cardHead}>
          <div className={styles.cardTitles}>
            {props.title && <h3 className={styles.cardTitle}>{props.title}</h3>}
            {props.subtitle && <p className={styles.cardSubtitle}>{props.subtitle}</p>}
          </div>
          {props.actions}
        </div>
      )}
      <div className={styles.body}>{props.children}</div>
    </section>
  );
}

export function Loading({ height }: { height?: number }): React.ReactElement {
  const { styles } = useDashboard();
  return <div className={styles.shimmer} style={{ height: height || 64 }} aria-busy="true" aria-label="Loading" />;
}

export function ErrorBox({ error }: { error: Error | string }): React.ReactElement {
  const { styles } = useDashboard();
  return <div className={styles.error} role="alert">{typeof error === 'string' ? error : error.message}</div>;
}

interface BoundaryProps { children?: React.ReactNode; render: (e: Error) => React.ReactNode; resetKey?: unknown }
interface BoundaryState { error?: Error; key?: unknown }

export class WidgetErrorBoundary extends React.Component<BoundaryProps, BoundaryState> {
  public state: BoundaryState = {};

  public static getDerivedStateFromError(error: Error): Partial<BoundaryState> {
    return { error };
  }

  public static getDerivedStateFromProps(props: BoundaryProps, state: BoundaryState): Partial<BoundaryState> | null {
    if (props.resetKey !== state.key) return { key: props.resetKey, error: undefined };
    return null;
  }

  public render(): React.ReactNode {
    return this.state.error ? this.props.render(this.state.error) : this.props.children;
  }
}
