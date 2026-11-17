import * as React from 'react';
import { FormatName, TextWidgetConfig } from '../../core/types';
import { evaluateAggregate } from '../../core/expr/evaluate';
import { formatValue } from '../../core/format';
import { useDashboard, useRows } from '../context';
import { ErrorBox, Loading, WidgetFrame } from '../WidgetFrame';

export function TextWidget({ w }: { w: TextWidgetConfig }): React.ReactElement {
  const { styles, locale, engine } = useDashboard();
  const { rows, error } = useRows(w.dataset);

  let body: React.ReactNode;
  if (error) body = <ErrorBox error={error} />;
  else if (w.dataset && !rows) body = <Loading height={40} />;
  else {
    const data = rows || [];
    const fill = (s: string): string => s.replace(/\{\{([\s\S]*?)\}\}/g, (_, inner: string) => {
      const [expr, fmt] = inner.split('|').map(x => x.trim());
      try {
        return formatValue(evaluateAggregate(expr, data, engine.vars), (fmt as FormatName) || undefined, locale);
      } catch (e) {
        return `[${(e as Error).message}]`;
      }
    });
    const paras = (w.text || '').split(/\n\s*\n/);
    body = (
      <div className={styles.text}>
        {paras.map((p, i) => (
          <p key={i}>
            {fill(p).split(/(\*\*[^*]+\*\*)/g).map((part, j) =>
              /^\*\*[^*]+\*\*$/.test(part) ? <strong key={j}>{part.slice(2, -2)}</strong> : <React.Fragment key={j}>{part}</React.Fragment>,
            )}
          </p>
        ))}
      </div>
    );
  }
  return <WidgetFrame title={w.title} subtitle={w.subtitle}>{body}</WidgetFrame>;
}
