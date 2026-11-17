import * as React from 'react';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Shimmer } from '@fluentui/react/lib/Shimmer';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { ColumnConfig, DataViewConfig, DataViewSource, MetricsConfig, Row } from '../../core/types';
import { DataContext } from '../../core/datasources';
import { MetricsEngine } from '../../core/engine';
import { toCsv } from '../../core/csv';
import { SpField, SpListService, rootField } from '../../core/sharepoint/listService';
import { DashboardContext, DashboardContextValue, useRows } from '../context';
import { MetricsTheme } from '../theme';
import { getStyles } from '../styles';
import { AppStyles } from './appStyles';
import { FieldValue, displayText } from './fieldDisplay';
import { GridColumn, GridTable } from './GridTable';
import { fieldColumns } from './ListView';
import { listConfigFor, spSource } from './useListModel';
import { RecordRoute } from './RecordView';

export interface DataViewProps {
  config: MetricsConfig;
  data: DataContext;
  theme: MetricsTheme;
  styles: AppStyles;
  locale?: string;
  reloadToken: number;
  onOpen: (source: string, id: number, order: number[]) => void;
  onOpenRecord: (record: RecordRoute) => void;
}

const HIDDEN_FIELD = /^(odata\.|@|__|OData__|ID$|AuthorId$|EditorId$|ContentTypeId$|ServerRedirectedEmbed|FileSystemObjectType$|GUID$|ComplianceAssetId$|Attachments$)/;

export function dataViewSources(config: MetricsConfig): DataViewSource[] {
  const view: DataViewConfig = config.dataView || {};
  if (view.sources && view.sources.length) return view.sources.map(s => (typeof s === 'string' ? { name: s } : s));
  return Object.keys(config.dataSources || {}).map(name => ({ name }));
}

export function dataViewTitle(source: DataViewSource, config: MetricsConfig): string {
  if (source.title) return source.title;
  const ds = config.dataSources && config.dataSources[source.name];
  if (ds && ds.type === 'spList' && ds.list) return ds.list;
  return source.name.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/^./, c => c.toUpperCase());
}

function isDisplayable(value: Row): boolean {
  return 'Title' in value || 'title' in value || 'Url' in value;
}

export function inferColumns(rows: Row[]): ColumnConfig[] {
  const names: string[] = [];
  for (const r of rows.slice(0, 50)) {
    for (const k of Object.keys(r)) {
      if (names.indexOf(k) >= 0 || HIDDEN_FIELD.test(k)) continue;
      const v = r[k];
      if (Array.isArray(v) && v.some(x => x && typeof x === 'object')) continue;
      if (v && typeof v === 'object' && !Array.isArray(v) && !isDisplayable(v as Row)) continue;
      names.push(k);
    }
  }
  return names.map(field => ({ field }));
}

function fileName(title: string): string {
  return `${title.replace(/[^\w\- ]+/g, '').trim() || 'data'}.csv`;
}

function download(name: string, text: string): void {
  const url = URL.createObjectURL(new Blob([text], { type: 'text/csv;charset=utf-8' }));
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function useFields(config: MetricsConfig, name: string, data: DataContext): Map<string, SpField> | undefined {
  const source = spSource(config, name);
  const [fields, setFields] = React.useState<Map<string, SpField> | undefined>(undefined);
  React.useEffect(() => {
    setFields(undefined);
    if (!source) return;
    let live = true;
    new SpListService(data, source).fields().then(
      list => {
        if (live) setFields(new Map(list.map((f): [string, SpField] => [f.name, f])));
      },
      () => {
        if (live) setFields(new Map());
      },
    );
    return () => {
      live = false;
    };
  }, [source, data]);
  return fields;
}

interface SourceTableProps {
  config: MetricsConfig;
  source: DataViewSource;
  title: string;
  view: DataViewConfig;
  data: DataContext;
  theme: MetricsTheme;
  styles: AppStyles;
  locale?: string;
  onOpen: (source: string, id: number, order: number[]) => void;
  onOpenRecord: (record: RecordRoute) => void;
}

function SourceTable(props: SourceTableProps): React.ReactElement {
  const { config, source, styles, theme, locale } = props;
  const { rows, error } = useRows(source.name);
  const isList = !!spSource(config, source.name);
  const fieldMap = useFields(config, source.name, props.data);
  const listConfig = React.useMemo(() => listConfigFor(config, source.name) || { source: source.name }, [config, source.name]);
  const [query, setQuery] = React.useState('');

  const columns = React.useMemo(
    () => (source.columns && source.columns.length ? source.columns.map(c => (typeof c === 'string' ? { field: c } : c)) : inferColumns(rows || [])),
    [source.columns, rows],
  );
  const map = React.useMemo(() => fieldMap || new Map<string, SpField>(), [fieldMap]);

  const gridColumns = React.useMemo<GridColumn[]>(
    () =>
      fieldColumns(columns, map, (c, field, row) => (
        <FieldValue path={c.field} field={field} row={row} list={listConfig} theme={theme} locale={locale} format={c.format} dense />
      )),
    [columns, map, listConfig, theme, locale],
  );

  const filteredRows = React.useMemo(() => {
    if (!rows || !query) return rows || [];
    const q = query.toLowerCase();
    return rows.filter(r => columns.some(c => displayText(map.get(rootField(c.field)), c.field, r, locale).toLowerCase().indexOf(q) >= 0));
  }, [rows, query, columns, map, locale]);

  if (error) return <MessageBar messageBarType={MessageBarType.error}>{error.message}</MessageBar>;
  const ready = rows && (!isList || fieldMap);
  const editable = isList && !!rows && rows.some(r => typeof r.Id === 'number');
  const open = (row: Row, ordered: Row[]): void => {
    if (editable) {
      props.onOpen(source.name, Number(row.Id), ordered.map(r => Number(r.Id)));
      return;
    }
    props.onOpenRecord({
      title: props.title,
      rows: ordered,
      index: Math.max(0, ordered.indexOf(row)),
      columns: columns.map(c => ({ column: c, title: (gridColumns.find(g => g.key === c.field) || { title: c.field }).title })),
    });
  };

  return (
    <div>
      <div className={styles.toolbar}>
        <SearchBox placeholder={`Search ${props.title.toLowerCase()}`} value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { width: 260 } }} />
        <span className={styles.toolbarSpacer} />
        {rows && <span className={styles.count}>{query ? `${filteredRows.length} of ${rows.length}` : rows.length} rows</span>}
        {props.view.export !== false && (
          <DefaultButton
            text="Export CSV"
            iconProps={{ iconName: 'Download' }}
            onClick={() => rows && download(fileName(props.title), toCsv(filteredRows, columns.map(c => ({ ...c, title: c.title || gridColumns.find(g => g.key === c.field)?.title })), locale))}
            disabled={!rows || !rows.length}
          />
        )}
      </div>
      {!ready ? (
        <div className={styles.gridCard} style={{ padding: 16, display: 'flex', flexDirection: 'column', gap: 14 }} aria-busy="true">
          {[0, 1, 2, 3, 4].map(i => <Shimmer key={i} width={`${90 - i * 7}%`} />)}
        </div>
      ) : (
        <GridTable
          rows={filteredRows}
          columns={gridColumns}
          styles={styles}
          pageSize={props.view.pageSize || 50}
          rowKey={row => String(row.Id ?? JSON.stringify(row))}
          rowLabel={row => `Open ${String(row.Title || row.title || row.Id || 'row')}`}
          onOpen={open}
          empty={query ? 'No rows match your search.' : 'No rows.'}
          ariaLabel={props.title}
        />
      )}
    </div>
  );
}

export function DataView({ config, data, theme, styles, locale, reloadToken, onOpen, onOpenRecord }: DataViewProps): React.ReactElement {
  const view: DataViewConfig = config.dataView || {};
  const sources = React.useMemo(() => dataViewSources(config), [config]);
  const [selected, setSelected] = React.useState<string | undefined>(sources.length ? sources[0].name : undefined);
  const engine = React.useMemo(() => new MetricsEngine(config, data), [config, data, reloadToken]);
  const tableStyles = getStyles(theme, config.theme && config.theme.cardStyle);
  const ctx = React.useMemo<DashboardContextValue>(
    () => ({ engine, theme, styles: tableStyles, locale: config.locale || locale }),
    [engine, theme, tableStyles, config.locale, locale],
  );

  React.useEffect(() => {
    if (!sources.some(s => s.name === selected)) setSelected(sources.length ? sources[0].name : undefined);
  }, [sources, selected]);

  if (!sources.length) return <div className={styles.empty}>This dashboard has no data sources.</div>;
  const current = sources.find(s => s.name === selected) || sources[0];

  return (
    <DashboardContext.Provider value={ctx}>
      {sources.length > 1 && (
        <Pivot
          selectedKey={current.name}
          onLinkClick={item => item && item.props.itemKey && setSelected(item.props.itemKey)}
          headersOnly
          linkFormat="tabs"
          linkSize="normal"
          aria-label="Data sources"
          styles={{ root: { marginBottom: 12 } }}
        >
          {sources.map(s => <PivotItem key={s.name} itemKey={s.name} headerText={dataViewTitle(s, config)} />)}
        </Pivot>
      )}
      <SourceTable
        key={`${current.name}-${reloadToken}`}
        config={config}
        source={current}
        title={dataViewTitle(current, config)}
        view={view}
        data={data}
        theme={theme}
        styles={styles}
        locale={locale}
        onOpen={onOpen}
        onOpenRecord={onOpenRecord}
      />
    </DashboardContext.Provider>
  );
}
