import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { useTheme } from '@fluentui/react/lib/Theme';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { JsonConfigEditor } from '../../components/JsonConfigEditor';
import { parseConfig } from '../../core/validate';

export interface ConfigEditorPanelProps {
  initial: string;
  onSave: (text: string) => void;
  onDismiss: () => void;
}

const classes = mergeStyleSets({
  editor: { flex: '1 1 auto', minHeight: 400, height: 'calc(100vh - 220px)', overflow: 'hidden' },
});

export function ConfigEditorPanel(props: ConfigEditorPanelProps): React.ReactElement {
  const theme = useTheme();
  const [text, setText] = React.useState(props.initial);
  const result = React.useMemo(() => parseConfig(text), [text]);
  const errors = result.diagnostics.filter(d => d.severity === 'error');

  const format = (): void => {
    if (result.config) setText(JSON.stringify(result.config, undefined, 2));
  };

  const footer = (): React.ReactElement => (
    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
      <PrimaryButton text="Save" onClick={() => props.onSave(text)} disabled={!result.config} />
      <DefaultButton text="Cancel" onClick={props.onDismiss} />
      <DefaultButton text="Format" onClick={format} disabled={!result.config} />
      <Stack.Item grow>
        <Text variant="small" styles={{ root: { color: errors.length ? theme.semanticColors.errorText : theme.semanticColors.bodySubtext } }}>
          {errors.length ? `${errors.length} problem(s). Hover the underlined text for details.` : 'No problems found'}
        </Text>
      </Stack.Item>
    </Stack>
  );

  return (
    <Panel isOpen type={PanelType.large} headerText="Metrics configuration" onDismiss={props.onDismiss} isFooterAtBottom onRenderFooterContent={footer}>
      <div className={classes.editor} style={{ border: `1px solid ${theme.semanticColors.inputBorder}` }}>
        <JsonConfigEditor value={text} onChange={setText} dark={theme.isInverted} />
      </div>
    </Panel>
  );
}
