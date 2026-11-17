import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { ChoiceGroup, IChoiceGroupOption } from '@fluentui/react/lib/ChoiceGroup';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Icon } from '@fluentui/react/lib/Icon';
import { Row } from '../../core/types';
import { AutoTagger, ItemSuggestion, currentTags } from '../../core/ai/autotag';
import { MetricsTheme } from '../theme';
import { AppStyles } from './appStyles';
import { ListModel, itemKey } from './useListModel';
import { SuggestionChips, TagList } from './TagChips';

export interface AutotagPanelProps {
  model: ListModel;
  tagger: AutoTagger;
  visibleRows: Row[];
  theme: MetricsTheme;
  styles: AppStyles;
  onDismiss: () => void;
  onApplied: (itemId: number) => Promise<void>;
}

type Scope = 'untagged' | 'visible';
type ApplyState = 'pending' | 'saving' | 'saved' | 'failed';

interface ResultRow extends ItemSuggestion {
  state?: ApplyState;
  message?: string;
}

export function AutotagPanel(props: AutotagPanelProps): React.ReactElement {
  const { model, tagger, visibleRows, theme, styles } = props;
  const maxItems = tagger.config.maxItems || 100;
  const untagged = React.useMemo(() => visibleRows.filter(r => !currentTags(tagger.field, r).length), [visibleRows, tagger]);
  const [scope, setScope] = React.useState<Scope>(untagged.length ? 'untagged' : 'visible');
  const [progress, setProgress] = React.useState<number | undefined>(undefined);
  const [results, setResults] = React.useState<ResultRow[] | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [applying, setApplying] = React.useState(false);
  const auto = tagger.config.mode === 'auto';

  const targets = (scope === 'untagged' ? untagged : visibleRows).slice(0, maxItems);
  const rowById = React.useMemo(() => new Map((model.rows || []).map((r): [number, Row] => [Number(r.Id), r])), [model.rows]);

  const options: IChoiceGroupOption[] = [
    { key: 'untagged', text: `Items without ${tagger.field.title.toLowerCase()} (${untagged.length})`, disabled: !untagged.length },
    { key: 'visible', text: `All items currently shown in the list (${visibleRows.length})` },
  ];

  const applyAll = async (rows: ResultRow[]): Promise<void> => {
    setApplying(true);
    const next = rows.slice();
    for (let i = 0; i < next.length; i++) {
      const chosen = next[i].suggested.filter(s => s.chosen).map(s => s.tag);
      const row = rowById.get(next[i].itemId);
      if (!chosen.length || !row || next[i].state === 'saved') continue;
      next[i] = { ...next[i], state: 'saving' };
      setResults(next.slice());
      try {
        await tagger.apply(row, chosen);
        await props.onApplied(next[i].itemId);
        next[i] = { ...next[i], state: 'saved' };
      } catch (e) {
        next[i] = { ...next[i], state: 'failed', message: (e as Error).message };
      }
      setResults(next.slice());
    }
    setApplying(false);
  };

  const run = async (): Promise<void> => {
    setError(undefined);
    setResults(undefined);
    setProgress(0);
    try {
      const suggestions = await tagger.suggest(targets, (done, total) => setProgress(done / total));
      const rows: ResultRow[] = suggestions.map(s => ({ ...s, state: s.suggested.length ? 'pending' : undefined }));
      setResults(rows);
      setProgress(undefined);
      if (auto) await applyAll(rows);
    } catch (e) {
      setError((e as Error).message);
      setProgress(undefined);
    }
  };

  const toggle = (itemId: number, tag: string): void => {
    setResults(rows =>
      rows && rows.map(r => (r.itemId !== itemId ? r : { ...r, suggested: r.suggested.map(s => (s.tag === tag ? { ...s, chosen: !s.chosen } : s)) })),
    );
  };

  const chosenCount = (results || []).filter(r => r.state !== 'saved' && r.suggested.some(s => s.chosen)).length;
  const running = progress !== undefined;

  const footer = (): React.ReactElement => (
    <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
      {!results && <PrimaryButton text={running ? 'Analyzing…' : `Suggest tags for ${targets.length} item${targets.length === 1 ? '' : 's'}`} onClick={run} disabled={running || !targets.length} />}
      {results && !auto && <PrimaryButton text={applying ? 'Applying…' : `Apply to ${chosenCount} item${chosenCount === 1 ? '' : 's'}`} onClick={() => applyAll(results)} disabled={applying || !chosenCount} />}
      {results && <DefaultButton text="Start over" onClick={() => setResults(undefined)} disabled={applying} />}
      <DefaultButton text="Close" onClick={props.onDismiss} disabled={applying} />
    </Stack>
  );

  return (
    <Panel isOpen type={PanelType.large} headerText="Auto-tag" onDismiss={props.onDismiss} isFooterAtBottom onRenderFooterContent={footer} isBlocking={applying}>
      <Stack tokens={{ childrenGap: 16 }} styles={{ root: { paddingTop: 8 } }}>
        <Text>
          Suggests values for <strong>{tagger.field.title}</strong> by comparing each item&apos;s text with the available tags using embeddings
          {tagger.config.api.chatModel ? ', then lets the chat model pick from the closest tags' : ''}.{' '}
          {auto ? 'Suggestions are applied automatically.' : 'Review the suggestions, then apply them.'}
        </Text>
        {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
        {!results && (
          <ChoiceGroup label="Which items?" options={options} selectedKey={scope} onChange={(_, o) => o && setScope(o.key === 'visible' ? 'visible' : 'untagged')} disabled={running} />
        )}
        {!results && (scope === 'untagged' ? untagged.length : visibleRows.length) > maxItems && (
          <MessageBar>Only the first {maxItems} items are analyzed per run (autotag.maxItems).</MessageBar>
        )}
        {running && <ProgressIndicator label="Analyzing items" percentComplete={progress} />}
        {results && (
          <div className={styles.tableWrap}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th className={styles.th} style={{ padding: '8px 12px', width: 96 }}>Key</th>
                  <th className={styles.th} style={{ padding: '8px 12px' }}>Item</th>
                  <th className={styles.th} style={{ padding: '8px 12px' }}>Current</th>
                  <th className={styles.th} style={{ padding: '8px 12px' }}>Suggested</th>
                  <th className={styles.th} style={{ padding: '8px 12px', width: 90 }}>Status</th>
                </tr>
              </thead>
              <tbody>
                {results.map(r => {
                  const row = rowById.get(r.itemId);
                  return (
                    <tr key={r.itemId}>
                      <td className={styles.td}><span className={styles.key}>{row ? itemKey(model, row) : r.itemId}</span></td>
                      <td className={styles.td}>
                        {r.title}
                        {r.usedFileText && <Icon iconName="Page" title="File text was used" styles={{ root: { marginLeft: 6, color: theme.textSecondary } }} />}
                      </td>
                      <td className={styles.td}><TagList tags={r.current} theme={theme} /></td>
                      <td className={styles.td}>
                        {r.error ? (
                          <span style={{ color: theme.status.critical, fontSize: 12 }}>{r.error}</span>
                        ) : (
                          <SuggestionChips suggestions={r.suggested} theme={theme} disabled={applying || r.state === 'saved'} onToggle={tag => toggle(r.itemId, tag)} />
                        )}
                      </td>
                      <td className={styles.td} style={{ fontSize: 12 }}>
                        {r.state === 'saving' && 'Saving…'}
                        {r.state === 'saved' && <span style={{ color: theme.deltaGood }}>Applied</span>}
                        {r.state === 'failed' && <span style={{ color: theme.status.critical }} title={r.message}>Failed</span>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </Stack>
    </Panel>
  );
}
