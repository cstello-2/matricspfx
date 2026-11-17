import * as React from 'react';
import { ColumnConfig, FormEntry, ListViewConfig, MetricsConfig, Row, SpListSourceConfig } from '../../core/types';
import { DataContext } from '../../core/datasources';
import { SpField, SpListInfo, SpListService, rootField } from '../../core/sharepoint/listService';

export interface ListModel {
  service: SpListService;
  config: ListViewConfig;
  info?: SpListInfo;
  fields?: SpField[];
  fieldMap: Map<string, SpField>;
  columns: ColumnConfig[];
  rows?: Row[];
  error?: Error;
  replaceRow: (row: Row) => void;
}

const DEFAULT_COLUMN_KINDS = ['choice', 'user', 'datetime', 'number', 'currency', 'lookup'];

export function spSource(config: MetricsConfig, name: string | undefined): SpListSourceConfig | undefined {
  const source = name && config.dataSources ? config.dataSources[name] : undefined;
  return source && source.type === 'spList' ? source : undefined;
}

export function listSource(config: MetricsConfig): SpListSourceConfig | undefined {
  return config.list ? spSource(config, config.list.source) : undefined;
}

export function listConfigFor(config: MetricsConfig, sourceName: string | undefined): ListViewConfig | undefined {
  if (!sourceName || !spSource(config, sourceName)) return undefined;
  if (config.list && config.list.source === sourceName) return config.list;
  return { source: sourceName, form: { comments: true } };
}

export function columnsFor(list: ListViewConfig, fields: SpField[] | undefined): ColumnConfig[] {
  if (list.columns && list.columns.length) return list.columns.map(c => (typeof c === 'string' ? { field: c } : c));
  if (!fields) return [{ field: 'Title' }];
  const extra = fields
    .filter(f => f.name !== 'Title' && DEFAULT_COLUMN_KINDS.indexOf(f.kind) >= 0 && !f.readOnly)
    .slice(0, 5)
    .map(f => ({ field: f.name }));
  return [{ field: 'Title' }].concat(extra);
}

const SYSTEM_FIELDS = ['ID', 'Id', 'Created', 'Modified', 'Author', 'Editor'];

export interface FormSectionLayout {
  title?: string;
  fields: string[];
  columns: 1 | 2;
}

export interface FormLayout {
  main: FormSectionLayout[];
  side: FormSectionLayout[];
}

function toSections(entries: FormEntry[], defaultTitle: string | undefined, used: Set<string>): FormSectionLayout[] {
  const sections: FormSectionLayout[] = [];
  let loose: string[] = [];
  const take = (names: string[]): string[] => {
    const out: string[] = [];
    for (const name of names.map(rootField)) {
      if (used.has(name)) continue;
      used.add(name);
      out.push(name);
    }
    return out;
  };
  const flush = (): void => {
    if (loose.length) sections.push({ title: sections.length ? undefined : defaultTitle, fields: loose, columns: 1 });
    loose = [];
  };
  for (const entry of entries) {
    if (typeof entry === 'string') {
      loose = loose.concat(take([entry]));
      continue;
    }
    flush();
    const fields = take(entry.fields || []);
    if (fields.length) sections.push({ title: entry.title, fields, columns: entry.columns === 2 ? 2 : 1 });
  }
  flush();
  return sections;
}

export function formLayout(list: ListViewConfig, columns: ColumnConfig[], schema?: SpField[]): FormLayout {
  const form = list.form || {};
  const used = new Set<string>([form.title || 'Title']);
  const main = toSections(form.main || (form.description ? [form.description] : []), undefined, used);
  const everything = !list.columns && schema ? schema.filter(f => f.kind !== 'other' && SYSTEM_FIELDS.indexOf(f.name) < 0).map(f => f.name) : undefined;
  const sideEntries: FormEntry[] = form.side || form.fields || everything || columns.map(c => rootField(c.field));
  const statusPlaced = !!form.status && (used.has(form.status) || sideEntries.some(e => (typeof e === 'string' ? e === form.status : e.fields.indexOf(form.status as string) >= 0)));
  const withStatus: FormEntry[] = form.status && !statusPlaced ? [form.status as FormEntry].concat(sideEntries) : sideEntries;
  const side = toSections(withStatus, 'Details', used);
  return { main, side };
}

export function formFields(list: ListViewConfig, columns: ColumnConfig[], schema?: SpField[]): string[] {
  const layout = formLayout(list, columns, schema);
  return layout.main.concat(layout.side).reduce<string[]>((all, s) => all.concat(s.fields), []);
}

function pathsFor(list: ListViewConfig, columns: ColumnConfig[], isLibrary: boolean, schema?: SpField[]): string[] {
  const form = list.form || {};
  const autotag = list.autotag;
  const paths = columns.map(c => c.field)
    .concat(formFields(list, columns, schema))
    .concat([form.title || 'Title'])
    .concat(form.description ? [form.description] : [])
    .concat(form.status ? [form.status] : [])
    .concat(list.quickFilters || [])
    .concat(autotag ? [autotag.field].concat(autotag.sourceFields || ['Title', 'Description']) : [])
    .concat(isLibrary ? ['FileRef', 'FileLeafRef'] : []);
  return paths.filter((p, i) => paths.indexOf(p) === i);
}

export function useListModel(config: MetricsConfig, data: DataContext, reloadToken: number): ListModel | undefined {
  return useSpListModel(config.list, listSource(config), data, reloadToken, true);
}

export function useSpListModel(
  list: ListViewConfig | undefined,
  source: SpListSourceConfig | undefined,
  data: DataContext,
  reloadToken: number,
  loadItems: boolean,
): ListModel | undefined {
  const service = React.useMemo(() => (source ? new SpListService(data, source) : undefined), [source, data]);
  const [info, setInfo] = React.useState<SpListInfo | undefined>(undefined);
  const [fields, setFields] = React.useState<SpField[] | undefined>(undefined);
  const [rows, setRows] = React.useState<Row[] | undefined>(undefined);
  const [error, setError] = React.useState<Error | undefined>(undefined);

  const columns = React.useMemo(() => (list ? columnsFor(list, fields) : []), [list, fields]);
  const fieldMap = React.useMemo(() => new Map((fields || []).map((f): [string, SpField] => [f.name, f])), [fields]);

  React.useEffect(() => {
    if (!service || !list) return;
    let live = true;
    setError(undefined);
    setRows(undefined);
    Promise.all([service.info(), service.fields()])
      .then(([listInfo, listFields]) => {
        if (!live) return undefined;
        setInfo(listInfo);
        setFields(listFields);
        if (!loadItems) return [];
        const known = new Set(listFields.map(f => f.name).concat(['Id', 'Title', 'FileRef', 'FileLeafRef']));
        return service.items(pathsFor(list, columnsFor(list, listFields), listInfo.isLibrary, listFields).filter(p => known.has(rootField(p))));
      })
      .then(
        items => {
          if (live && items) setRows(items);
        },
        (e: Error) => {
          if (live) setError(e);
        },
      );
    return () => {
      live = false;
    };
  }, [service, list, reloadToken, loadItems]);

  const replaceRow = React.useCallback((row: Row) => {
    setRows(current => (current ? current.map(r => (r.Id === row.Id ? row : r)) : current));
  }, []);

  return React.useMemo(
    () => (service && list ? { service, config: list, info, fields, fieldMap, columns, rows, error, replaceRow } : undefined),
    [service, list, info, fields, fieldMap, columns, rows, error, replaceRow],
  );
}

export function itemPaths(model: ListModel): string[] {
  const known = new Set(Array.from(model.fieldMap.keys()).concat(['Id', 'Title', 'FileRef', 'FileLeafRef']));
  return pathsFor(model.config, model.columns, !!model.info && model.info.isLibrary, model.fields).filter(p => known.has(rootField(p)));
}

export function itemKey(model: ListModel, row: Row): string {
  const prefix = model.config.keyPrefix || (model.info ? model.info.title : model.service.source.list || 'ITEM').replace(/[^A-Za-z]/g, '').slice(0, 4).toUpperCase();
  return `${prefix}-${String(row.Id)}`;
}
