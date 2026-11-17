import type { IFontStyles, IPalette, ISemanticColors } from '@fluentui/react/lib/Styling';
import { Level, ThemeConfig } from '../core/types';

export interface MetricsTheme {
  isDark: boolean;
  fontFamily: string;
  surface: string;
  text: string;
  textSecondary: string;
  textMuted: string;
  grid: string;
  axis: string;
  border: string;
  accent: string;
  link: string;
  hover: string;
  palette: string[];
  sequential: string[];
  ordinal: string[];
  themeShades: string[];
  status: Record<Level, string>;
  deltaGood: string;
  deltaBad: string;
}

const CATEGORICAL_LIGHT = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#6250d6', '#e34948'];
const CATEGORICAL_DARK = ['#3987e5', '#d95926', '#199e70', '#c98500', '#d55181', '#0ca30c', '#9085e9', '#e66767'];
const BLUE_RAMP = ['#cde2fb', '#b7d3f6', '#9ec5f4', '#86b6ef', '#6da7ec', '#5598e7', '#3987e5', '#2a78d6', '#256abf', '#1c5cab', '#184f95', '#104281', '#0d366b'];
const ORDINAL_LIGHT = ['#104281', '#184f95', '#1c5cab', '#256abf', '#2a78d6', '#3987e5', '#5598e7', '#6da7ec', '#86b6ef'];
const ORDINAL_DARK = ['#86b6ef', '#6da7ec', '#5598e7', '#3987e5', '#2a78d6', '#256abf', '#1c5cab', '#184f95'];

const STATUS: Record<Level, string> = {
  good: '#0ca30c',
  warning: '#fab219',
  serious: '#ec835a',
  critical: '#d03b3b',
  neutral: '#898781',
  info: '#2a78d6',
};

const FONT = '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif';

export const LIGHT_THEME: MetricsTheme = {
  isDark: false,
  fontFamily: FONT,
  surface: '#ffffff',
  text: '#323130',
  textSecondary: '#605e5c',
  textMuted: '#8a8886',
  grid: '#edebe9',
  axis: '#c8c6c4',
  border: '#edebe9',
  accent: '#0078d4',
  link: '#0078d4',
  hover: 'rgba(0,0,0,0.04)',
  palette: CATEGORICAL_LIGHT,
  sequential: BLUE_RAMP,
  ordinal: ORDINAL_LIGHT,
  themeShades: ['#eff6fc', '#deecf9', '#c7e0f4', '#71afe5', '#2b88d8', '#0078d4', '#106ebe', '#005a9e', '#004578'],
  status: STATUS,
  deltaGood: '#107c10',
  deltaBad: '#d03b3b',
};

export const DARK_THEME: MetricsTheme = {
  ...LIGHT_THEME,
  isDark: true,
  surface: '#1b1a19',
  text: '#f3f2f1',
  textSecondary: '#c8c6c4',
  textMuted: '#a19f9d',
  grid: '#323130',
  axis: '#484644',
  border: '#323130',
  hover: 'rgba(255,255,255,0.06)',
  palette: CATEGORICAL_DARK,
  sequential: BLUE_RAMP.slice(1, 11).reverse(),
  ordinal: ORDINAL_DARK,
  deltaGood: '#0ca30c',
  deltaBad: '#e66767',
};

export interface SectionTheme {
  readonly isInverted?: boolean;
  readonly palette?: Readonly<Partial<IPalette>>;
  readonly semanticColors?: Readonly<Partial<ISemanticColors>>;
  readonly fonts?: Readonly<Partial<IFontStyles>>;
}

export function themeFromSharePoint(theme: SectionTheme | undefined): MetricsTheme {
  if (!theme) return LIGHT_THEME;
  const base = theme.isInverted ? DARK_THEME : LIGHT_THEME;
  const p = theme.palette || {};
  const s = theme.semanticColors || {};
  const font = theme.fonts && theme.fonts.medium && theme.fonts.medium.fontFamily;
  const shades = [p.themeLighterAlt, p.themeLighter, p.themeLight, p.themeTertiary, p.themeSecondary, p.themePrimary, p.themeDarkAlt, p.themeDark, p.themeDarker];
  return {
    ...base,
    fontFamily: font || base.fontFamily,
    surface: s.cardStandoutBackground || s.bodyBackground || base.surface,
    text: s.bodyText || base.text,
    textSecondary: s.bodySubtext || base.textSecondary,
    textMuted: p.neutralSecondaryAlt || p.neutralSecondary || base.textMuted,
    grid: s.bodyDivider || base.grid,
    axis: p.neutralTertiaryAlt || base.axis,
    border: s.bodyDivider || base.border,
    accent: p.themePrimary || base.accent,
    link: s.link || base.link,
    hover: s.listItemBackgroundHovered || base.hover,
    status: { ...base.status, info: p.themePrimary || base.status.info },
    themeShades: shades.every((c): c is string => !!c) ? shades : base.themeShades,
  };
}

export function withPalette(t: MetricsTheme, mode: ThemeConfig['palette']): MetricsTheme {
  if (Array.isArray(mode) && mode.length) return { ...t, palette: mode };
  if (mode === 'theme') {
    return {
      ...t,
      palette: [t.accent].concat(t.palette.slice(1)),
      sequential: t.themeShades,
      ordinal: t.themeShades.slice(3).reverse(),
    };
  }
  return t;
}
