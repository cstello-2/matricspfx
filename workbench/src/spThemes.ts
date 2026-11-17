import { createTheme, ITheme } from '@fluentui/react/lib/Styling';
import { BaseSlots, ThemeGenerator, themeRulesStandardCreator } from '@fluentui/react/lib/ThemeGenerator';
import { getNeutralVariant, getSoftVariant, getStrongVariant } from '@fluentui/scheme-utilities';

export interface SiteTheme {
  name: string;
  primary: string;
  text: string;
  background: string;
  inverted: boolean;
}

export const SITE_THEMES: SiteTheme[] = [
  { name: 'Blue', primary: '#0078d4', text: '#323130', background: '#ffffff', inverted: false },
  { name: 'Teal', primary: '#03787c', text: '#323130', background: '#ffffff', inverted: false },
  { name: 'Purple', primary: '#8764b8', text: '#323130', background: '#ffffff', inverted: false },
  { name: 'Red', primary: '#a4262c', text: '#323130', background: '#ffffff', inverted: false },
  { name: 'Dark blue', primary: '#3a96dd', text: '#ffffff', background: '#1f1f1f', inverted: true },
  { name: 'Dark yellow', primary: '#fce100', text: '#ffffff', background: '#1f1f1f', inverted: true },
];

export type SectionBackground = 'none' | 'neutral' | 'soft' | 'strong';

export const SECTION_BACKGROUNDS: SectionBackground[] = ['none', 'neutral', 'soft', 'strong'];

export function siteTheme(t: SiteTheme): ITheme {
  const rules = themeRulesStandardCreator();
  ThemeGenerator.insureSlots(rules, t.inverted);
  ThemeGenerator.setSlot(rules[BaseSlots[BaseSlots.primaryColor]], t.primary, t.inverted, true, true);
  ThemeGenerator.setSlot(rules[BaseSlots[BaseSlots.foregroundColor]], t.text, t.inverted, true, true);
  ThemeGenerator.setSlot(rules[BaseSlots[BaseSlots.backgroundColor]], t.background, t.inverted, true, true);
  return createTheme({ palette: ThemeGenerator.getThemeAsJson(rules), isInverted: t.inverted });
}

export function sectionTheme(site: ITheme, background: SectionBackground): ITheme {
  switch (background) {
    case 'neutral':
      return getNeutralVariant(site);
    case 'soft':
      return getSoftVariant(site);
    case 'strong':
      return getStrongVariant(site);
    default:
      return site;
  }
}

export const WIDTHS: Array<{ name: string; px: number | undefined }> = [
  { name: 'Full width', px: undefined },
  { name: 'One column', px: 1204 },
  { name: 'Two-thirds', px: 790 },
  { name: 'Half', px: 586 },
  { name: 'One-third', px: 380 },
  { name: 'Phone', px: 343 },
];
