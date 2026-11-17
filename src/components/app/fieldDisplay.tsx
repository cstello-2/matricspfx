import * as React from 'react';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { Format, Level, ListViewConfig, Row } from '../../core/types';
import { getPath, isNil } from '../../core/expr/evaluate';
import { formatValue } from '../../core/format';
import { toDate } from '../../core/dates';
import { SpField, rootField } from '../../core/sharepoint/listService';
import { MetricsTheme } from '../theme';
import { relativeTime } from './time';

const DEFAULT_LEVELS: Array<[RegExp, Level]> = [
  [/^(done|completed?|closed|resolved|approved|shipped)$/i, 'good'],
  [/^(blocked|rejected|failed|critical|cancell?ed)$/i, 'critical'],
  [/^(at risk|on hold|waiting.*)$/i, 'warning'],
  [/^(in progress|active|assigned|in review|doing|open)$/i, 'info'],
];

export function statusLevel(value: unknown, colors: ListViewConfig['statusColors']): Level {
  const text = String(value);
  if (colors && colors[text]) return colors[text];
  const match = DEFAULT_LEVELS.find(([re]) => re.test(text));
  return match ? match[1] : 'neutral';
}

function hexToRgb(hex: string): [number, number, number] | undefined {
  const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex.trim());
  return m ? [parseInt(m[1], 16), parseInt(m[2], 16), parseInt(m[3], 16)] : undefined;
}

export function tint(color: string, alpha: number): string {
  const rgb = hexToRgb(color);
  return rgb ? `rgba(${rgb[0]}, ${rgb[1]}, ${rgb[2]}, ${alpha})` : color;
}

export function Lozenge({ value, level, theme }: { value: string; level: Level; theme: MetricsTheme }): React.ReactElement {
  const color = theme.status[level];
  return (
    <span
      style={{
        display: 'inline-block',
        padding: '2px 8px',
        borderRadius: 4,
        fontSize: 11,
        fontWeight: 700,
        lineHeight: '16px',
        textTransform: 'uppercase',
        letterSpacing: '0.03em',
        whiteSpace: 'nowrap',
        color: theme.text,
        background: tint(color, theme.isDark ? 0.3 : 0.16),
      }}
    >
      {value}
    </span>
  );
}

export function Pill({ value, theme }: { value: string; theme: MetricsTheme }): React.ReactElement {
  return (
    <span
      style={{
        display: 'inline-block',
        padding: '1px 8px',
        borderRadius: 10,
        fontSize: 12,
        lineHeight: '18px',
        whiteSpace: 'nowrap',
        color: theme.textSecondary,
        background: theme.hover,
        border: `1px solid ${theme.border}`,
      }}
    >
      {value}
    </span>
  );
}

type PriorityRank = 'highest' | 'high' | 'medium' | 'low' | 'lowest';

const PRIORITY_RANKS: Array<[RegExp, PriorityRank]> = [
  [/^(highest|critical|urgent|blocker|p0|1)\b/i, 'highest'],
  [/^(high|major|p1|2)\b/i, 'high'],
  [/^(medium|normal|moderate|p2|3)\b/i, 'medium'],
  [/^(low|minor|p3|4)\b/i, 'low'],
  [/^(lowest|trivial|p4|5)\b/i, 'lowest'],
];

export function priorityRank(value: string): PriorityRank | undefined {
  const match = PRIORITY_RANKS.find(([re]) => re.test(value.trim()));
  return match ? match[1] : undefined;
}

export function PriorityIcon({ rank, theme }: { rank: PriorityRank; theme: MetricsTheme }): React.ReactElement {
  const color = rank === 'highest' ? theme.status.critical : rank === 'high' ? theme.status.serious : rank === 'medium' ? theme.status.warning : theme.palette[0];
  const paths: Record<PriorityRank, string> = {
    highest: 'M3 9.5 8 5l5 4.5M3 13 8 8.5l5 4.5',
    high: 'M3 11 8 6l5 5',
    medium: 'M3 6.5h10M3 10.5h10',
    low: 'M3 6l5 5 5-5',
    lowest: 'M3 3.5 8 8l5-4.5M3 7l5 4.5L13 7',
  };
  return (
    <svg width={16} height={16} viewBox="0 0 16 16" aria-hidden="true" focusable="false" style={{ flex: '0 0 auto' }}>
      <path d={paths[rank]} fill="none" stroke={color} strokeWidth={2} strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

export interface PersonValue {
  name: string;
  email?: string;
}

export function people(value: unknown): PersonValue[] {
  const list = Array.isArray(value) ? value : isNil(value) ? [] : [value];
  return list
    .filter(v => v && typeof v === 'object')
    .map(v => {
      const o = v as Row;
      return { name: String(o.Title || o.name || ''), email: o.EMail ? String(o.EMail) : o.email ? String(o.email) : undefined };
    })
    .filter(p => p.name);
}

export function plainText(html: string): string {
  if (!/<[a-z][\s\S]*>/i.test(html)) return html;
  return new DOMParser().parseFromString(html, 'text/html').body.textContent || '';
}

export function displayText(field: SpField | undefined, path: string, row: Row, locale?: string): string {
  const value = getPath(row, path);
  if (isNil(value) || value === '') return '';
  if (!field || path !== field.name) return formatValue(value, undefined, locale);
  switch (field.kind) {
    case 'lookup':
    case 'multilookup': {
      const display = field.lookupField || 'Title';
      const items = Array.isArray(value) ? value : [value];
      return items.map(v => (v && typeof v === 'object' ? String((v as Row)[display] ?? '') : String(v))).filter(Boolean).join(', ');
    }
    case 'boolean':
      return value ? 'Yes' : 'No';
    case 'datetime':
      return formatValue(value, field.dateOnly ? 'date' : 'datetime', locale);
    case 'note':
      return plainText(String(value));
    case 'url': {
      const o = value as Row;
      return String(o.Description || o.Url || '');
    }
    case 'user':
    case 'multiuser':
      return people(value).map(p => p.name).join(', ');
    default:
      return formatValue(value, undefined, locale);
  }
}

const Empty = ({ theme }: { theme: MetricsTheme }): React.ReactElement => <span style={{ color: theme.textMuted }} aria-label="Empty">—</span>;

export interface FieldCellProps {
  path: string;
  field: SpField | undefined;
  row: Row;
  list: ListViewConfig;
  theme: MetricsTheme;
  locale?: string;
  format?: Format;
  dense?: boolean;
}

function isStatus(path: string, list: ListViewConfig, value: unknown): boolean {
  return rootField(path) === (list.form && list.form.status) || !!(list.statusColors && list.statusColors[String(value)]);
}

export function FieldValue({ path, field, row, list, theme, locale, format, dense }: FieldCellProps): React.ReactElement {
  const value = getPath(row, path);
  const own = !!field && path === field.name;
  if (isNil(value) || value === '' || (Array.isArray(value) && !value.length)) {
    if (own && field && field.kind === 'user') return <span style={{ color: theme.textMuted }}>Unassigned</span>;
    return <Empty theme={theme} />;
  }
  if (format) return <span>{formatValue(value, format, locale)}</span>;

  if (own && field && (field.kind === 'user' || field.kind === 'multiuser')) {
    const ps = people(value);
    return (
      <span style={{ display: 'inline-flex', flexDirection: dense ? 'row' : 'column', gap: 4, flexWrap: 'wrap' }}>
        {ps.map(p => <Persona key={p.name} text={p.name} size={PersonaSize.size24} styles={{ primaryText: { fontSize: 13 } }} />)}
      </span>
    );
  }

  if (field && (field.kind === 'choice' || field.kind === 'multichoice')) {
    const values = Array.isArray(value) ? value.map(String) : [String(value)];
    if (isStatus(path, list, values[0])) {
      return (
        <span style={{ display: 'inline-flex', gap: 4, flexWrap: 'wrap' }}>
          {values.map(v => <Lozenge key={v} value={v} level={statusLevel(v, list.statusColors)} theme={theme} />)}
        </span>
      );
    }
    const rank = values.length === 1 && /priority/i.test(field.name) ? priorityRank(values[0]) : undefined;
    if (rank) {
      return (
        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
          <PriorityIcon rank={rank} theme={theme} />
          {values[0]}
        </span>
      );
    }
    if (field.kind === 'multichoice') {
      return (
        <span style={{ display: 'inline-flex', gap: 4, flexWrap: 'wrap' }}>
          {values.map(v => <Pill key={v} value={v} theme={theme} />)}
        </span>
      );
    }
  }

  if (own && field && field.kind === 'datetime') {
    const date = toDate(value);
    if (date) {
      const full = date.toLocaleString(locale, field.dateOnly ? { dateStyle: 'medium' } : { dateStyle: 'medium', timeStyle: 'short' });
      if (field.dateOnly) return <span title={full}>{date.toLocaleDateString(locale, { month: 'short', day: 'numeric', year: 'numeric' })}</span>;
      return <span title={full}>{relativeTime(date.toISOString(), locale)}</span>;
    }
  }

  if (own && field && field.kind === 'boolean') return <span>{value ? 'Yes' : 'No'}</span>;

  if (own && field && field.kind === 'url' && typeof value === 'object') {
    const o = value as Row;
    const href = String(o.Url || '');
    if (/^(https?:\/\/|\/)/i.test(href)) {
      return <a href={href} target="_blank" rel="noopener noreferrer" style={{ color: theme.link }} onClick={e => e.stopPropagation()}>{String(o.Description || href)}</a>;
    }
  }

  const text = displayText(field, path, row, locale);
  if (!text) return <Empty theme={theme} />;
  if (dense && field && field.kind === 'note') {
    return <span style={{ display: 'block', maxWidth: 360, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={text}>{text}</span>;
  }
  return <span>{text}</span>;
}
