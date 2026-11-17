import * as React from 'react';
import { Level, Row, ThresholdConfig } from '../core/types';
import { evaluate } from '../core/expr/evaluate';

export const LEVEL_LABEL: Record<Level, string> = { good: 'On track', warning: 'At risk', serious: 'Off track', critical: 'Critical', neutral: 'Neutral', info: 'In progress' };

export function matchThreshold(thresholds: ThresholdConfig[] | undefined, value: unknown, target: unknown, row?: Row): ThresholdConfig | undefined {
  if (!thresholds) return undefined;
  const scope = { row: { ...(row || {}), value, target }, vars: { value, target } };
  return thresholds.find(t => !!evaluate(t.when, scope));
}

export function StatusIcon({ level, color }: { level: Level; color: string }): React.ReactElement {
  let mark: React.ReactNode = undefined;
  if (level === 'good') mark = <path d="M2 5.2 4.2 7.4 8 3" fill="none" stroke="#fff" strokeWidth={1.6} strokeLinecap="round" strokeLinejoin="round" />;
  else if (level === 'warning' || level === 'serious') mark = <path d="M5 2.4v3.2M5 7.4v.1" stroke={level === 'warning' ? '#0b0b0b' : '#fff'} strokeWidth={1.6} strokeLinecap="round" />;
  else if (level === 'critical') mark = <path d="M3.2 3.2 6.8 6.8M6.8 3.2 3.2 6.8" stroke="#fff" strokeWidth={1.6} strokeLinecap="round" />;
  return (
    <svg width={12} height={12} viewBox="0 0 10 10" aria-hidden="true" focusable="false">
      <circle cx={5} cy={5} r={5} fill={color} />
      {mark}
    </svg>
  );
}
