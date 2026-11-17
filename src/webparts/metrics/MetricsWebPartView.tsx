import * as React from 'react';
import { ThemeProvider } from '@fluentui/react/lib/Theme';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { MetricsApp } from '../../components/app/MetricsApp';
import { SectionTheme, themeFromSharePoint } from '../../components/theme';
import { MetricsConfig } from '../../core/types';
import { DataContext } from '../../core/datasources';
import { ConfigEditorPanel } from './ConfigEditorPanel';

export interface MetricsWebPartViewProps {
  config?: MetricsConfig;
  error?: string;
  data: DataContext;
  sectionTheme?: SectionTheme;
  locale?: string;
  reloadToken: number;
  instanceKey: string;
  isEditMode: boolean;
  editorOpen: boolean;
  editorText: string;
  onOpenEditor: () => void;
  onCloseEditor: () => void;
  onSaveConfig: (text: string) => void;
  onUseStarter: () => void;
}

function Placeholder(props: { isEditMode: boolean; onOpenEditor: () => void; onUseStarter: () => void }): React.ReactElement {
  return (
    <Stack horizontalAlign="center" tokens={{ childrenGap: 12, padding: 24 }}>
      <Text variant="large">Metrics</Text>
      <Text>
        {props.isEditMode
          ? 'Describe KPIs, charts and tables over your SharePoint lists with a JSON configuration.'
          : 'This dashboard has not been configured yet.'}
      </Text>
      {props.isEditMode && (
        <Stack horizontal wrap horizontalAlign="center" tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="Open JSON editor" onClick={props.onOpenEditor} />
          <DefaultButton text="Use starter example" onClick={props.onUseStarter} />
        </Stack>
      )}
    </Stack>
  );
}

export function MetricsWebPartView(props: MetricsWebPartViewProps): React.ReactElement {
  const theme = React.useMemo(() => themeFromSharePoint(props.sectionTheme), [props.sectionTheme]);
  return (
    <>
      <ThemeProvider theme={props.sectionTheme} applyTo="none">
        {props.error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline>
            <span style={{ whiteSpace: 'pre-wrap' }}>{props.error}</span>
          </MessageBar>
        )}
        {!props.error && !props.config && <Placeholder isEditMode={props.isEditMode} onOpenEditor={props.onOpenEditor} onUseStarter={props.onUseStarter} />}
        {!props.error && props.config && (
          <MetricsApp
            config={props.config}
            data={props.data}
            theme={theme}
            locale={props.locale}
            reloadToken={props.reloadToken}
            instanceKey={props.instanceKey}
          />
        )}
      </ThemeProvider>
      {props.editorOpen && (
        <ConfigEditorPanel initial={props.editorText} onSave={props.onSaveConfig} onDismiss={props.onCloseEditor} />
      )}
    </>
  );
}
