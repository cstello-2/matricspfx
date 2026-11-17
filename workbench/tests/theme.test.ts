import { describe, it, expect } from 'vitest';
import { themeFromSharePoint, withPalette, LIGHT_THEME } from '../../src/components/theme';
import { SITE_THEMES, sectionTheme, siteTheme } from '../src/spThemes';

describe('theme from SharePoint', () => {
  const teal = siteTheme(SITE_THEMES[1]);

  it('takes colors and font from the section theme', () => {
    const t = themeFromSharePoint(teal);
    expect(t.isDark).toBe(false);
    expect(t.accent).toBe(teal.palette.themePrimary);
    expect(t.text).toBe(teal.semanticColors.bodyText);
    expect(t.surface).toBe(teal.semanticColors.cardStandoutBackground);
    expect(t.fontFamily).toBe(teal.fonts.medium.fontFamily);
    expect(t.themeShades[5]).toBe(teal.palette.themePrimary);
  });

  it('follows section background variants', () => {
    const strong = sectionTheme(teal, 'strong');
    const t = themeFromSharePoint(strong);
    expect(strong.isInverted).toBe(true);
    expect(t.isDark).toBe(true);
    expect(t.text).toBe(strong.semanticColors.bodyText);
    expect(themeFromSharePoint(sectionTheme(teal, 'neutral')).surface).toBe(sectionTheme(teal, 'neutral').semanticColors.cardStandoutBackground);
  });

  it('uses the dark palette steps on dark site themes', () => {
    const t = themeFromSharePoint(siteTheme(SITE_THEMES[4]));
    expect(t.isDark).toBe(true);
    expect(t.palette[0]).toBe('#3987e5');
  });

  it('falls back to defaults when SharePoint provides no theme', () => {
    expect(themeFromSharePoint(undefined)).toBe(LIGHT_THEME);
  });

  it('can lead the palette and ramps with the site color', () => {
    const t = withPalette(themeFromSharePoint(teal), 'theme');
    expect(t.palette[0]).toBe(teal.palette.themePrimary);
    expect(t.sequential).toEqual(t.themeShades);
  });
});
