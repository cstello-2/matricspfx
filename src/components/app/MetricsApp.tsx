import * as React from 'react';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Modal } from '@fluentui/react/lib/Modal';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { MetricsConfig } from '../../core/types';
import { DataContext } from '../../core/datasources';
import { MetricsDashboard } from '../MetricsDashboard';
import { MetricsTheme } from '../theme';
import { getAppStyles } from './appStyles';
import { listConfigFor, spSource, useListModel, useSpListModel } from './useListModel';
import { useAutotag } from './useAutotag';
import { ListView } from './ListView';
import { ItemView } from './ItemView';
import { DataView } from './DataView';
import { RecordRoute, RecordView } from './RecordView';

interface ItemBoundaryProps {
  children: React.ReactNode;
  onClose: () => void;
}

interface ItemBoundaryState {
  error?: Error;
}

class ItemErrorBoundary extends React.Component<ItemBoundaryProps, ItemBoundaryState> {
  public state: ItemBoundaryState = {};

  public static getDerivedStateFromError(error: Error): ItemBoundaryState {
    return { error };
  }

  public render(): React.ReactNode {
    if (!this.state.error) return this.props.children;
    return (
      <Modal isOpen onDismiss={this.props.onClose} isBlocking={false} styles={{ main: { maxWidth: 520, padding: 24 } }}>
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          The item form could not be shown: {this.state.error.message}
        </MessageBar>
        <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 16 }}>
          <DefaultButton text="Close" onClick={this.props.onClose} />
        </div>
      </Modal>
    );
  }
}

export interface MetricsAppProps {
  config: MetricsConfig;
  data: DataContext;
  theme: MetricsTheme;
  locale?: string;
  reloadToken: number;
  instanceKey: string;
}

type Tab = 'list' | 'metrics' | 'data';

const TABS: Tab[] = ['list', 'metrics', 'data'];

export interface ItemRoute {
  source: string;
  id: number | undefined;
}

export function parseItemRoute(value: string | null, listSourceName: string | undefined): ItemRoute | undefined {
  if (!value) return undefined;
  const sep = value.lastIndexOf(':');
  const source = sep > 0 ? value.slice(0, sep) : listSourceName;
  const idPart = sep > 0 ? value.slice(sep + 1) : value;
  if (!source) return undefined;
  if (idPart === 'new') return { source, id: undefined };
  const id = Number(idPart);
  return isNaN(id) ? undefined : { source, id };
}

export function formatItemRoute(route: ItemRoute, listSourceName: string | undefined): string {
  const id = route.id === undefined ? 'new' : String(route.id);
  return route.source === listSourceName ? id : `${route.source}:${id}`;
}

function useItemRoute(param: string, listSourceName: string | undefined): [ItemRoute | undefined, (route: ItemRoute | undefined, replace?: boolean) => void] {
  const read = React.useCallback(() => parseItemRoute(new URLSearchParams(window.location.search).get(param), listSourceName), [param, listSourceName]);
  const [route, setRoute] = React.useState<ItemRoute | undefined>(read);
  const pushed = React.useRef(0);
  React.useEffect(() => {
    const onPop = (): void => {
      const next = read();
      if (!next) pushed.current = 0;
      setRoute(next);
    };
    window.addEventListener('popstate', onPop);
    return () => window.removeEventListener('popstate', onPop);
  }, [read]);
  const navigate = React.useCallback(
    (next: ItemRoute | undefined, replace?: boolean) => {
      if (!next && pushed.current > 0) {
        pushed.current = 0;
        window.history.back();
        return;
      }
      if (next && !replace) pushed.current += 1;
      setRoute(next);
      const url = new URL(window.location.href);
      if (next) url.searchParams.set(param, formatItemRoute(next, listSourceName));
      else url.searchParams.delete(param);
      if (replace) window.history.replaceState(window.history.state, '', url.toString());
      else window.history.pushState(window.history.state, '', url.toString());
    },
    [param, listSourceName],
  );
  return [route, navigate];
}

export function MetricsApp(props: MetricsAppProps): React.ReactElement {
  const { config, theme } = props;
  const styles = getAppStyles(theme);
  const [dataVersion, setDataVersion] = React.useState(0);
  const model = useListModel(config, props.data, props.reloadToken);
  const autotag = useAutotag(model, props.data);
  const [tab, setTab] = React.useState<Tab>(config.list ? 'list' : 'metrics');
  const listSourceName = config.list ? config.list.source : undefined;
  const [route, navigate] = useItemRoute(`item-${props.instanceKey}`, listSourceName);
  const [order, setOrder] = React.useState<number[]>([]);
  const [record, setRecord] = React.useState<RecordRoute | undefined>(undefined);

  const otherSource = route && route.source !== listSourceName ? route.source : undefined;
  const otherList = React.useMemo(() => listConfigFor(config, otherSource), [config, otherSource]);
  const otherModel = useSpListModel(otherList, spSource(config, otherSource), props.data, props.reloadToken + dataVersion, false);
  const itemModel = route ? (route.source === listSourceName ? model : otherModel) : undefined;

  const showData = config.dataView !== false;
  React.useEffect(() => {
    if ((tab === 'list' && !config.list) || (tab === 'data' && !showData)) setTab(config.list ? 'list' : 'metrics');
  }, [config.list, showData, tab]);

  const open = (source: string, id: number | undefined, ids: number[]): void => {
    setOrder(ids);
    navigate({ source, id });
  };

  const title = config.title || (model && model.info ? model.info.title : 'Metrics');
  const listLabel = (config.list && config.list.title) || (model && model.info ? model.info.title : 'List');
  const changed = (): void => setDataVersion(v => v + 1);

  return (
    <div className={styles.app}>
      <header className={styles.header}>
        <div className={styles.headerRow}>
          <div className={styles.avatar} aria-hidden="true">{title.charAt(0).toUpperCase()}</div>
          <div>
            <h2 className={styles.appTitle}>{title}</h2>
            {config.description && <p className={styles.appDescription}>{config.description}</p>}
          </div>
        </div>
        <Pivot
          className={styles.tabs}
          selectedKey={tab}
          onLinkClick={item => {
            const key = TABS.find(t => item !== undefined && t === item.props.itemKey);
            if (key) setTab(key);
          }}
          headersOnly
          aria-label="Views"
        >
          {config.list && <PivotItem headerText={listLabel} itemKey="list" itemIcon="BulletedList" />}
          <PivotItem headerText="Metrics" itemKey="metrics" itemIcon="BarChartVertical" />
          {showData && <PivotItem headerText={(config.dataView && config.dataView.title) || 'Data'} itemKey="data" itemIcon="Table" />}
        </Pivot>
      </header>

      <div className={styles.tabBody} role="tabpanel">
        {tab === 'list' && model && listSourceName && (
          <ListView
            model={model}
            theme={theme}
            styles={styles}
            locale={props.locale}
            onOpen={(id, ids) => open(listSourceName, id, ids)}
            onCreate={() => open(listSourceName, undefined, [])}
            autotag={autotag}
            onChanged={changed}
          />
        )}
        {tab === 'metrics' && (
          <MetricsDashboard config={config} data={props.data} theme={theme} locale={props.locale} reloadToken={props.reloadToken + dataVersion} hideHeader />
        )}
        {tab === 'data' && showData && (
          <DataView config={config} data={props.data} theme={theme} styles={styles} locale={props.locale} reloadToken={props.reloadToken + dataVersion} onOpen={open} onOpenRecord={setRecord} />
        )}
      </div>

      {record && (
        <RecordView record={record} locale={props.locale} onClose={() => setRecord(undefined)} onMove={index => setRecord({ ...record, index })} />
      )}

      {route && itemModel && itemModel.fields && (
        <ItemErrorBoundary key={`${route.source}:${route.id === undefined ? 'new' : route.id}`} onClose={() => navigate(undefined)}>
        <ItemView
          model={itemModel}
          id={route.id}
          order={order.length ? order : (itemModel.rows || []).map(r => Number(r.Id))}
          webUrl={props.data.webUrl}
          currentUser={props.data.user ? props.data.user.displayName : undefined}
          locale={props.locale}
          autotag={route.source === listSourceName ? autotag : {}}
          onClose={() => navigate(undefined)}
          onNavigate={next => navigate({ source: route.source, id: next }, true)}
          onCreated={id => navigate({ source: route.source, id }, true)}
          onChanged={changed}
        />
        </ItemErrorBoundary>
      )}
    </div>
  );
}
