import * as React from 'react';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { TagSuggestion } from '../../core/ai/autotag';
import { MetricsTheme } from '../theme';
import { tint } from './fieldDisplay';

function chipStyles(t: MetricsTheme): { chip: string; chosen: string; plain: string; score: string } {
  return mergeStyleSets({
    chip: {
      display: 'inline-flex',
      alignItems: 'center',
      gap: 6,
      padding: '2px 10px',
      borderRadius: 12,
      border: `1px solid ${t.axis}`,
      background: 'transparent',
      color: t.text,
      font: 'inherit',
      fontSize: 12,
      lineHeight: '20px',
      cursor: 'pointer',
      selectors: { ':focus-visible': { outline: `2px solid ${t.accent}`, outlineOffset: '1px' }, ':disabled': { cursor: 'default', opacity: 0.6 } },
    },
    chosen: { background: tint(t.accent, t.isDark ? 0.35 : 0.14), borderColor: t.accent },
    plain: {
      display: 'inline-block',
      padding: '2px 10px',
      borderRadius: 12,
      background: t.hover,
      color: t.text,
      fontSize: 12,
      lineHeight: '20px',
    },
    score: { color: t.textSecondary, fontVariantNumeric: 'tabular-nums' },
  });
}

export function TagList({ tags, theme }: { tags: string[]; theme: MetricsTheme }): React.ReactElement {
  const s = chipStyles(theme);
  if (!tags.length) return <span style={{ color: theme.textMuted, fontSize: 12 }}>None</span>;
  return (
    <span style={{ display: 'inline-flex', flexWrap: 'wrap', gap: 4 }}>
      {tags.map(t => <span key={t} className={s.plain}>{t}</span>)}
    </span>
  );
}

export interface SuggestionChipsProps {
  suggestions: TagSuggestion[];
  theme: MetricsTheme;
  disabled?: boolean;
  onToggle: (tag: string) => void;
}

export function SuggestionChips({ suggestions, theme, disabled, onToggle }: SuggestionChipsProps): React.ReactElement {
  const s = chipStyles(theme);
  if (!suggestions.length) return <span style={{ color: theme.textMuted, fontSize: 12 }}>No new tags above the score threshold</span>;
  return (
    <span style={{ display: 'inline-flex', flexWrap: 'wrap', gap: 4 }}>
      {suggestions.map(sg => (
        <button
          key={sg.tag}
          type="button"
          className={sg.chosen ? `${s.chip} ${s.chosen}` : s.chip}
          aria-pressed={sg.chosen}
          disabled={disabled}
          onClick={() => onToggle(sg.tag)}
          title={`Similarity ${sg.score.toFixed(3)}`}
        >
          <span aria-hidden="true">{sg.chosen ? '✓' : '+'}</span>
          {sg.tag}
          <span className={s.score}>{Math.round(sg.score * 100)}%</span>
        </button>
      ))}
    </span>
  );
}
