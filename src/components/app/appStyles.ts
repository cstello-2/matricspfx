import { mergeStyleSets, IStyle, IProcessedStyleSet } from '@fluentui/react/lib/Styling';
import { MetricsTheme } from '../theme';

type AppStyleKey =
  | 'app' | 'header' | 'headerRow' | 'avatar' | 'appTitle' | 'appDescription' | 'tabs' | 'tabBody'
  | 'toolbar' | 'toolbarSpacer' | 'count' | 'gridCard' | 'gridScroll' | 'table' | 'th' | 'thButton' | 'sortActive' | 'sortIdle'
  | 'td' | 'row' | 'rowStatic' | 'openCell' | 'openIcon' | 'primaryCell' | 'key' | 'pager' | 'pagerButtons' | 'pageNumber' | 'empty' | 'tableWrap'
  | 'item' | 'itemTop' | 'breadcrumb' | 'itemTopActions' | 'itemBody' | 'itemMain' | 'itemSide' | 'section' | 'sectionTitle'
  | 'sideCard' | 'sideCardTitle' | 'fieldGrid' | 'fieldGridTwo' | 'longText' | 'readOnlyField' | 'readOnlyLabel' | 'saveBar' | 'saveBarText' | 'meta' | 'keyBadge'
  | 'comment' | 'commentHead' | 'commentAuthor' | 'commentTime' | 'commentText' | 'composer' | 'composerActions';

export type AppStyles = IProcessedStyleSet<Record<AppStyleKey, IStyle>>;

function styleSet(t: MetricsTheme): Record<AppStyleKey, IStyle> {
  const focus = { outline: `2px solid ${t.accent}`, outlineOffset: '-2px' };
  const card = {
    background: t.surface,
    border: `1px solid ${t.border}`,
    borderRadius: 8,
    boxShadow: t.isDark ? 'none' : '0 1px 2px rgba(0,0,0,0.04)',
  };
  return {
    app: { fontFamily: t.fontFamily, color: t.text, width: '100%', boxSizing: 'border-box', selectors: { '*': { boxSizing: 'border-box' } } },
    header: { marginBottom: 20, borderBottom: `1px solid ${t.border}` },
    headerRow: { display: 'flex', alignItems: 'center', gap: 12, paddingBottom: 8 },
    avatar: {
      width: 36,
      height: 36,
      borderRadius: 8,
      background: t.accent,
      color: '#ffffff',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      fontWeight: 700,
      fontSize: 17,
      flex: '0 0 auto',
    },
    appTitle: { fontSize: 20, fontWeight: 600, margin: 0, lineHeight: '26px', color: t.text },
    appDescription: { fontSize: 13, margin: '2px 0 0', color: t.textSecondary, lineHeight: '18px' },
    tabs: { marginBottom: -1 },
    tabBody: {},

    toolbar: { display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', marginBottom: 12 },
    toolbarSpacer: { flex: '1 1 auto' },
    count: { fontSize: 13, color: t.textSecondary, whiteSpace: 'nowrap' },
    gridCard: { ...card, overflow: 'hidden' },
    gridScroll: { overflowX: 'auto' },
    tableWrap: { ...card, overflowX: 'auto' },
    table: { width: '100%', borderCollapse: 'separate', borderSpacing: '0', fontSize: 14, color: t.text },
    th: {
      position: 'sticky',
      top: 0,
      zIndex: 1,
      textAlign: 'left',
      fontSize: 12,
      fontWeight: 600,
      color: t.textSecondary,
      background: t.surface,
      borderBottom: `1px solid ${t.border}`,
      padding: 0,
      whiteSpace: 'nowrap',
    },
    thButton: {
      display: 'flex',
      alignItems: 'center',
      gap: 6,
      width: '100%',
      border: 'none',
      background: 'transparent',
      color: 'inherit',
      font: 'inherit',
      padding: '10px 14px',
      cursor: 'pointer',
      selectors: { ':hover': { color: t.text }, ':focus-visible': focus },
    },
    sortActive: { fontSize: 12, color: t.accent },
    sortIdle: { fontSize: 12, color: t.textMuted, opacity: 0.45 },
    td: { padding: '10px 14px', borderBottom: `1px solid ${t.grid}`, verticalAlign: 'middle', lineHeight: '20px' },
    row: {
      cursor: 'pointer',
      outline: 'none',
      selectors: {
        td: { transition: 'background-color .12s ease' },
        ':hover td': { background: t.hover },
        ':focus-visible td': { background: t.hover },
        ':focus-visible td:first-child': { boxShadow: `inset 3px 0 0 ${t.accent}` },
        ':last-child td': { borderBottom: 'none' },
        ':hover $openIcon': { opacity: 1, transform: 'translateX(0)' },
        ':focus-visible $openIcon': { opacity: 1, transform: 'translateX(0)' },
      },
    },
    openCell: { width: 36, padding: '0 12px 0 0', textAlign: 'right' },
    openIcon: { color: t.accent, fontSize: 12, opacity: 0, transform: 'translateX(-4px)', transition: 'opacity .12s ease, transform .12s ease' },
    rowStatic: { selectors: { ':hover td': { background: t.hover }, ':last-child td': { borderBottom: 'none' } } },
    primaryCell: { fontWeight: 600, color: t.text, minWidth: 220 },
    key: { color: t.link, fontSize: 13, whiteSpace: 'nowrap', fontWeight: 600, fontVariantNumeric: 'tabular-nums' },
    pager: {
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'space-between',
      gap: 8,
      flexWrap: 'wrap',
      padding: '4px 8px 4px 14px',
      borderTop: `1px solid ${t.border}`,
      fontSize: 13,
      color: t.textSecondary,
    },
    pagerButtons: { display: 'inline-flex', alignItems: 'center', gap: 2 },
    pageNumber: { padding: '0 8px', whiteSpace: 'nowrap' },
    empty: { padding: '48px 24px', textAlign: 'center', color: t.textMuted, fontSize: 14 },

    item: {
      fontFamily: t.fontFamily,
      color: t.text,
      background: t.surface,
      height: '100%',
      display: 'flex',
      flexDirection: 'column',
      boxSizing: 'border-box',
      selectors: { '*': { boxSizing: 'border-box' } },
    },
    itemTop: {
      display: 'flex',
      alignItems: 'center',
      gap: 12,
      padding: '0 16px 0 24px',
      borderBottom: `1px solid ${t.border}`,
      minHeight: 56,
      flex: '0 0 auto',
      flexWrap: 'wrap',
    },
    breadcrumb: { display: 'flex', alignItems: 'center', gap: 8, fontSize: 14, color: t.textSecondary, flex: '1 1 auto', minWidth: 0, whiteSpace: 'nowrap' },
    itemTopActions: { display: 'flex', alignItems: 'center', gap: 4, flex: '0 0 auto', whiteSpace: 'nowrap', marginLeft: 'auto' },
    keyBadge: { color: t.textSecondary, fontWeight: 600, fontVariantNumeric: 'tabular-nums' },
    itemBody: {
      flex: '1 1 auto',
      overflow: 'auto',
      display: 'grid',
      gridTemplateColumns: 'minmax(0, 1fr) minmax(300px, 380px)',
      gap: 40,
      padding: '28px 40px 40px',
      maxWidth: 1320,
      width: '100%',
      margin: '0 auto',
      alignContent: 'start',
      selectors: { '@media (max-width: 900px)': { gridTemplateColumns: 'minmax(0, 1fr)', gap: 24, padding: '20px 16px' } },
    },
    itemMain: { minWidth: 0, display: 'flex', flexDirection: 'column', gap: 28 },
    itemSide: { minWidth: 0, display: 'flex', flexDirection: 'column', gap: 16 },
    section: { display: 'flex', flexDirection: 'column', gap: 10 },
    sectionTitle: { fontSize: 14, fontWeight: 600, margin: 0, color: t.text },
    sideCard: { ...card, padding: 16, display: 'flex', flexDirection: 'column', gap: 14 },
    sideCardTitle: { fontSize: 12, fontWeight: 600, margin: 0, color: t.textSecondary, textTransform: 'uppercase', letterSpacing: '0.04em' },
    fieldGrid: { display: 'flex', flexDirection: 'column', gap: 14 },
    fieldGridTwo: {
      display: 'grid',
      gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
      gap: 14,
      alignItems: 'start',
      selectors: { '@media (max-width: 640px)': { gridTemplateColumns: 'minmax(0, 1fr)' } },
    },
    longText: { fontSize: 14, lineHeight: '22px', whiteSpace: 'pre-wrap', color: t.text },
    readOnlyField: { display: 'flex', flexDirection: 'column', gap: 4, fontSize: 14 },
    readOnlyLabel: { fontSize: 14, fontWeight: 600, color: t.text },
    saveBar: {
      flex: '0 0 auto',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'flex-end',
      gap: 8,
      flexWrap: 'wrap',
      padding: '12px 24px',
      borderTop: `1px solid ${t.border}`,
      background: t.surface,
      boxShadow: t.isDark ? 'none' : '0 -4px 16px rgba(0,0,0,0.06)',
    },
    saveBarText: { marginRight: 'auto', fontSize: 13, color: t.textSecondary },
    meta: { fontSize: 12, color: t.textSecondary, lineHeight: '20px' },

    comment: { display: 'grid', gridTemplateColumns: '32px minmax(0, 1fr)', gap: 12 },
    commentHead: { display: 'flex', alignItems: 'baseline', gap: 8, flexWrap: 'wrap' },
    commentAuthor: { fontWeight: 600, fontSize: 14 },
    commentTime: { fontSize: 12, color: t.textSecondary },
    commentText: { fontSize: 14, lineHeight: '22px', whiteSpace: 'pre-wrap', wordBreak: 'break-word', marginTop: 2 },
    composer: { display: 'grid', gridTemplateColumns: '32px minmax(0, 1fr)', gap: 12, alignItems: 'start' },
    composerActions: { display: 'flex', gap: 8, marginTop: 8, alignItems: 'center' },
  };
}

const cache = new Map<string, AppStyles>();

export function getAppStyles(t: MetricsTheme): AppStyles {
  const key = JSON.stringify(t);
  let s = cache.get(key);
  if (!s) {
    s = mergeStyleSets(styleSet(t));
    cache.set(key, s);
  }
  return s;
}
