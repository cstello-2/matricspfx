import * as React from 'react';
import { KpiWidgetConfig } from '../../core/types';
import { evaluateAggregate, fieldOrExpr, toNumber } from '../../core/expr/evaluate';
import { formatValue } from '../../core/format';
import { useDashboard, useRows } from '../context';
import { ErrorBox, Loading, WidgetFrame } from '../WidgetFrame';
import { LEVEL_LABEL, StatusIcon, matchThreshold } from '../status';

function Sparkline({ values, color, muted, height = 32 }: { values: number[]; color: string; muted: string; height?: number }): React.ReactElement | null {
  const { styles } = useDashboard();
  if (values.length < 2) return null;
  const w = 200;
  const pad = 4;
  const min = Math.min(...values);
  const max = Math.max(...values);
  const span = max - min || 1;
  const pts = values.map((v, i) => [pad + (i * (w - pad * 2)) / (values.length - 1), pad + (1 - (v - min) / span) * (height - pad * 2)]);
  const d = pts.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join(' ');
  const last = pts[pts.length - 1];
  return (
    <svg className={styles.spark} viewBox={`0 0 ${w} ${height}`} preserveAspectRatio="none" height={height} aria-hidden="true">
      <path d={d} fill="none" stroke={muted} strokeWidth={2} strokeLinejoin="round" strokeLinecap="round" vectorEffect="non-scaling-stroke" />
      <circle cx={last[0]} cy={last[1]} r={3.5} fill={color} />
    </svg>
  );
}

export function KpiWidget({ w }: { w: KpiWidgetConfig }): React.ReactElement {
  const { theme, styles, locale, engine } = useDashboard();
  const main = useRows(w.dataset);
  const cmp = useRows(w.compare ? w.compare.dataset || w.dataset : undefined);
  const spark = useRows(w.sparkline ? w.sparkline.dataset || w.dataset : undefined);

  let body: React.ReactNode;
  if (main.error) body = <ErrorBox error={main.error} />;
  else if (!main.rows) body = <Loading height={72} />;
  else {
    try {
      const rows = main.rows;
      const value = evaluateAggregate(w.value, rows, engine.vars);
      const target = typeof w.target === 'string' ? evaluateAggregate(w.target, rows, engine.vars) : w.target;
      const level = matchThreshold(w.thresholds, value, target);

      let delta: React.ReactNode = undefined;
      if (w.compare && cmp.rows) {
        const prev = toNumber(evaluateAggregate(w.compare.value, cmp.rows, engine.vars));
        const cur = toNumber(value);
        if (!isNaN(prev) && !isNaN(cur)) {
          const diff = cur - prev;
          const pctMode = (w.compare.mode || 'percent') === 'percent';
          const shown = pctMode ? (prev === 0 ? null : diff / Math.abs(prev)) : diff;
          const up = diff > 0;
          const good = diff === 0 ? undefined : (w.compare.goodWhen || 'up') === 'up' ? up : !up;
          const color = good === undefined ? theme.textSecondary : good ? theme.deltaGood : theme.deltaBad;
          delta = (
            <span className={styles.delta} style={{ color }}>
              <span aria-hidden="true">{diff === 0 ? '–' : up ? '▲' : '▼'}</span>
              {shown === null ? 'n/a' : (diff > 0 ? '+' : '') + (pctMode ? formatValue(shown, { type: 'percent', decimals: 1 }, locale) : formatValue(shown, w.format, locale))}
              <span className={styles.deltaNote}>{w.compare.label || 'vs previous'}</span>
            </span>
          );
        }
      }

      const tNum = toNumber(target);
      const vNum = toNumber(value);
      const meter = !isNaN(tNum) && tNum > 0 && !isNaN(vNum);
      const meterColor = level ? theme.status[level.level] : theme.accent;

      let sparkValues: number[] = [];
      const sparkline = w.sparkline;
      if (sparkline && spark.rows) {
        sparkValues = spark.rows
          .map(r => toNumber(fieldOrExpr(sparkline.y, r)))
          .filter(n => !isNaN(n));
      }

      body = (
        <div>
          <p className={styles.kpiLabel}>{w.label || (w.title ? '' : w.value)}</p>
          <div className={styles.kpiValue} title={String(value)}>{formatValue(value, w.format, locale)}</div>
          <div className={styles.kpiRow}>
            {delta}
            {level && (
              <span className={styles.pill}>
                <StatusIcon level={level.level} color={theme.status[level.level]} />
                {level.label || LEVEL_LABEL[level.level]}
              </span>
            )}
          </div>
          {meter && (
            <>
              <div className={styles.meterTrack} role="meter" aria-valuenow={vNum} aria-valuemin={0} aria-valuemax={tNum} aria-label="Progress to target">
                <div className={styles.meterFill} style={{ width: `${Math.max(0, Math.min(100, (vNum / tNum) * 100))}%`, background: meterColor }} />
              </div>
              <div className={styles.meterLegend}>
                <span>{formatValue(vNum / tNum, { type: 'percent', decimals: 0 }, locale)} of target</span>
                <span>{formatValue(tNum, w.format, locale)}</span>
              </div>
            </>
          )}
          {sparkValues.length > 1 && <Sparkline values={sparkValues} color={theme.accent} muted={theme.textMuted} />}
        </div>
      );
    } catch (e) {
      body = <ErrorBox error={e as Error} />;
    }
  }

  return <WidgetFrame title={w.title} subtitle={w.subtitle}>{body}</WidgetFrame>;
}
