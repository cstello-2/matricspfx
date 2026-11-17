import { mergeStyleSets, keyframes, IStyle, IProcessedStyleSet } from '@fluentui/react/lib/Styling';
import { MetricsTheme } from './theme';
import { ThemeConfig } from '../core/types';

type StyleKey =
  | 'root' | 'header' | 'title' | 'description' | 'grid' | 'cell'
  | 'card' | 'cardHead' | 'cardTitles' | 'cardTitle' | 'cardSubtitle' | 'iconButton' | 'body'
  | 'shimmer' | 'error' | 'empty'
  | 'kpiLabel' | 'kpiValue' | 'kpiRow' | 'delta' | 'deltaNote' | 'pill' | 'meterTrack' | 'meterFill' | 'meterLegend' | 'spark'
  | 'tableTools' | 'search' | 'tableWrap' | 'table' | 'th' | 'sortMark' | 'td' | 'num' | 'row' | 'bar' | 'barTrack' | 'barFill'
  | 'link' | 'statusCell' | 'expandCell' | 'chevron' | 'chevronOpen' | 'detailCell' | 'detailTitle' | 'pager'
  | 'text';

export type Styles = IProcessedStyleSet<Record<StyleKey, IStyle>>;

const shimmer = keyframes({
  '0%': { backgroundPosition: '-400px 0' },
  '100%': { backgroundPosition: '400px 0' },
});

function styleSet(t: MetricsTheme, cardStyle: NonNullable<ThemeConfig['cardStyle']>): Record<StyleKey, IStyle> {
  return {
    root: {
      fontFamily: t.fontFamily,
      color: t.text,
      width: '100%',
      boxSizing: 'border-box',
      selectors: { '*': { boxSizing: 'border-box' } },
    },
    header: { marginBottom: 16 },
    title: { fontSize: 20, fontWeight: 600, margin: 0, color: t.text, lineHeight: '28px' },
    description: { fontSize: 14, color: t.textSecondary, margin: '4px 0 0' },
    grid: { display: 'grid', alignItems: 'stretch' },
    cell: { display: 'flex', flexDirection: 'column', minWidth: 0 },
    card: {
      background: t.surface,
      borderRadius: 8,
      padding: 16,
      minWidth: 0,
      flex: '1 1 auto',
      display: 'flex',
      flexDirection: 'column',
      border: cardStyle === 'outlined' ? `1px solid ${t.border}` : '1px solid transparent',
      boxShadow: cardStyle === 'elevated' ? '0 1.6px 3.6px rgba(0,0,0,0.13), 0 0.3px 0.9px rgba(0,0,0,0.11)' : 'none',
    },
    cardHead: { display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 12 },
    cardTitles: { flex: '1 1 auto', minWidth: 0 },
    cardTitle: { fontSize: 14, fontWeight: 600, color: t.text, margin: 0, lineHeight: '20px' },
    cardSubtitle: { fontSize: 12, color: t.textMuted, margin: '2px 0 0', lineHeight: '16px' },
    iconButton: {
      border: 'none',
      background: 'transparent',
      color: t.textMuted,
      cursor: 'pointer',
      borderRadius: 4,
      padding: '2px 6px',
      fontSize: 12,
      fontFamily: t.fontFamily,
      selectors: {
        ':hover': { background: t.hover, color: t.text },
        ':focus-visible': { outline: `2px solid ${t.accent}`, outlineOffset: '1px' },
        ':disabled': { opacity: 0.4, cursor: 'default' },
      },
    },
    body: { flex: '1 1 auto', minHeight: 0, position: 'relative' },
    shimmer: {
      height: '100%',
      minHeight: 48,
      borderRadius: 6,
      backgroundImage: `linear-gradient(90deg, ${t.grid} 0px, ${t.hover} 120px, ${t.grid} 240px)`,
      backgroundSize: '800px 100%',
      animationName: shimmer,
      animationDuration: '1.4s',
      animationIterationCount: 'infinite',
      animationTimingFunction: 'linear',
    },
    error: {
      fontSize: 12,
      color: t.text,
      borderLeft: `3px solid ${t.status.critical}`,
      padding: '8px 10px',
      whiteSpace: 'pre-wrap',
      wordBreak: 'break-word',
    },
    empty: { fontSize: 13, color: t.textMuted, padding: '12px 0' },

    kpiLabel: { fontSize: 13, color: t.textSecondary, margin: 0, lineHeight: '18px' },
    kpiValue: { fontSize: 32, fontWeight: 600, color: t.text, lineHeight: '40px', margin: '4px 0 2px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' },
    kpiRow: { display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', fontSize: 12, color: t.textSecondary, minHeight: 18 },
    delta: { display: 'inline-flex', alignItems: 'center', gap: 3, fontWeight: 600 },
    deltaNote: { color: t.textMuted, fontWeight: 400 },
    pill: { display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: 12, color: t.textSecondary },
    meterTrack: { height: 6, borderRadius: 3, background: t.grid, overflow: 'hidden', marginTop: 10 },
    meterFill: { height: '100%', borderRadius: 3, transition: 'width .4s ease' },
    meterLegend: { display: 'flex', justifyContent: 'space-between', fontSize: 11, color: t.textMuted, marginTop: 4 },
    spark: { marginTop: 8, display: 'block', width: '100%' },

    tableTools: { display: 'flex', justifyContent: 'flex-end', marginBottom: 8 },
    search: {
      fontFamily: t.fontFamily,
      fontSize: 13,
      padding: '5px 10px',
      borderRadius: 4,
      border: `1px solid ${t.axis}`,
      background: 'transparent',
      color: t.text,
      width: 200,
      maxWidth: '100%',
      selectors: { ':focus': { outline: `2px solid ${t.accent}`, outlineOffset: '-1px' } },
    },
    tableWrap: { overflow: 'auto', maxWidth: '100%' },
    table: { width: '100%', borderCollapse: 'collapse', fontSize: 13, color: t.text },
    th: {
      textAlign: 'left',
      fontWeight: 600,
      fontSize: 12,
      color: t.textSecondary,
      padding: '8px 10px',
      borderBottom: `1px solid ${t.axis}`,
      whiteSpace: 'nowrap',
      cursor: 'pointer',
      userSelect: 'none',
      position: 'sticky',
      top: 0,
      background: t.surface,
      selectors: { ':hover': { color: t.text }, ':focus-visible': { outline: `2px solid ${t.accent}`, outlineOffset: '-2px' } },
    },
    sortMark: { marginLeft: 4 },
    td: { padding: '8px 10px', borderBottom: `1px solid ${t.grid}`, verticalAlign: 'middle' },
    num: { fontVariantNumeric: 'tabular-nums', whiteSpace: 'nowrap' },
    row: { selectors: { ':hover td': { background: t.hover } } },
    bar: { display: 'flex', alignItems: 'center', gap: 8, justifyContent: 'flex-end' },
    barTrack: { flex: '0 0 72px', height: 6, borderRadius: 3, background: t.grid, overflow: 'hidden' },
    barFill: { display: 'block', height: '100%', background: t.palette[0], borderRadius: 3 },
    link: { color: t.link, textDecoration: 'none', selectors: { ':hover': { textDecoration: 'underline' } } },
    statusCell: { display: 'inline-flex', alignItems: 'center', gap: 6 },
    expandCell: { width: 24, padding: 4 },
    chevron: { display: 'inline-block', transition: 'transform .15s' },
    chevronOpen: { display: 'inline-block', transition: 'transform .15s', transform: 'rotate(90deg)' },
    detailCell: { padding: '4px 10px 12px 34px', background: t.hover, borderBottom: `1px solid ${t.grid}` },
    detailTitle: { fontSize: 12, fontWeight: 600, color: t.textSecondary, margin: '8px 0 4px' },
    pager: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 8, fontSize: 12, color: t.textMuted },

    text: { fontSize: 14, lineHeight: '22px', color: t.text, selectors: { p: { margin: '0 0 8px' }, strong: { fontWeight: 600 } } },
  };
}

const cache = new Map<string, Styles>();

export function getStyles(t: MetricsTheme, cardStyle: ThemeConfig['cardStyle'] = 'outlined'): Styles {
  const key = JSON.stringify(t) + cardStyle;
  let s = cache.get(key);
  if (!s) {
    s = mergeStyleSets(styleSet(t, cardStyle));
    cache.set(key, s);
  }
  return s;
}
