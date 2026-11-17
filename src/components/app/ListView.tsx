import * as React from 'react';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { ActionButton, DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Shimmer } from '@fluentui/react/lib/Shimmer';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { ColumnConfig, Row } from '../../core/types';
import { compare, evaluate, getPath } from '../../core/expr/evaluate';
import { parseSort } from '../../core/engine';
import { SpField, isEditable, rootField } from '../../core/sharepoint/listService';
import { MetricsTheme } from '../theme';
import { AppStyles } from './appStyles';
import { ListModel, itemKey, itemPaths } from './useListModel';
import { AutotagState } from './useAutotag';
import { AutotagPanel } from './AutotagPanel';
import { FieldValue, displayText } from './fieldDisplay';
import { GridColumn, GridTable } from './GridTable';

export interface ListViewProps {
  model: ListModel;
  theme: MetricsTheme;
  styles: AppStyles;
  locale?: string;
  onOpen: (id: number, order: number[]) => void;
  onCreate: () => void;
  autotag: AutotagState;
  onChanged: () => void;
}

export function columnTitle(fieldMap: Map<string, SpField>, c: ColumnConfig): string {
  if (c.title) return c.title;
  const field = fieldMap.get(rootField(c.field));
  if (!field) return c.field.replace(/([a-z])([A-Z])/g, '$1 $2');
  return c.field === field.name ? field.title : `${field.title} ${c.field.split('.').slice(1).join(' ')}`;
}

function filterOptions(model: ListModel, path: string, rows: Row[]): IDropdownOption[] {
  const field = model.fieldMap.get(rootField(path));
  if (field && field.choices.length && path === field.name) return field.choices.map(c => ({ key: c, text: c }));
  const seen = new Set<string>();
  for (const r of rows) {
    const text = displayText(field, path, r);
    if (text) text.split(', ').forEach(t => seen.add(t));
  }
  return Array.from(seen).sort((a, b) => a.localeCompare(b)).map(t => ({ key: t, text: t }));
}

export function fieldColumns(
  columns: ColumnConfig[],
  fieldMap: Map<string, SpField>,
  render: (c: ColumnConfig, field: SpField | undefined, row: Row) => React.ReactNode,
): GridColumn[] {
  return columns.map((c, i) => {
    const field = fieldMap.get(rootField(c.field));
    const numeric = !!field && c.field === field.name && (field.kind === 'number' || field.kind === 'currency');
    return {
      key: c.field,
      title: columnTitle(fieldMap, c),
      width: c.width,
      align: c.align || (numeric ? 'right' : undefined),
      primary: c.field === 'Title' || (i === 0 && !fieldMap.has('Title')),
      sortValue: (row: Row) => {
        const v = getPath(row, c.field);
        if (v && typeof v === 'object' && !Array.isArray(v)) return displayText(field, c.field, row);
        return Array.isArray(v) ? displayText(field, c.field, row) : v;
      },
      render: (row: Row) => render(c, field, row),
    };
  });
}

function LoadingRows({ styles }: { styles: AppStyles }): React.ReactElement {
  return (
    <div className={styles.gridCard} style={{ padding: 16, display: 'flex', flexDirection: 'column', gap: 14 }} aria-busy="true" aria-label="Loading items">
      {[0, 1, 2, 3, 4, 5].map(i => <Shimmer key={i} width={`${90 - i * 6}%`} />)}
    </div>
  );
}

export function ListView({ model, theme, styles, locale, onOpen, onCreate, autotag, onChanged }: ListViewProps): React.ReactElement {
  const [tagging, setTagging] = React.useState(false);
  const { config, columns, rows, error } = model;
  const [query, setQuery] = React.useState('');
  const [filters, setFilters] = React.useState<Record<string, string[]>>({});

  const visible = React.useMemo(() => {
    if (!rows) return [];
    let out = rows;
    if (config.filter) {
      const expr = config.filter;
      out = out.filter(r => !!evaluate(expr, { row: r }));
    }
    for (const path of Object.keys(filters)) {
      const selected = filters[path];
      if (!selected.length) continue;
      const field = model.fieldMap.get(rootField(path));
      out = out.filter(r => displayText(field, path, r).split(', ').some(v => selected.indexOf(v) >= 0));
    }
    if (query) {
      const q = query.toLowerCase();
      out = out.filter(r =>
        itemKey(model, r).toLowerCase().indexOf(q) >= 0 ||
        columns.some(c => displayText(model.fieldMap.get(rootField(c.field)), c.field, r, locale).toLowerCase().indexOf(q) >= 0),
      );
    }
    const keys = config.sort ? parseSort(config.sort) : [{ field: 'Id', desc: true }];
    return out.slice().sort((a, b) => {
      for (const k of keys) {
        const c = compare(getPath(a, k.field), getPath(b, k.field));
        if (c) return k.desc ? -c : c;
      }
      return 0;
    });
  }, [rows, config, filters, query, columns, model, locale]);

  const gridColumns = React.useMemo<GridColumn[]>(() => {
    const keyColumn: GridColumn = {
      key: 'Id',
      title: 'Key',
      width: 104,
      sortValue: row => Number(row.Id),
      render: row => <span className={styles.key}>{itemKey(model, row)}</span>,
    };
    return [keyColumn].concat(
      fieldColumns(columns, model.fieldMap, (c, field, row) => (
        <FieldValue path={c.field} field={field} row={row} list={config} theme={theme} locale={locale} format={c.format} dense />
      )),
    );
  }, [columns, model, config, theme, locale, styles]);

  const filtered = !!query || Object.keys(filters).some(k => filters[k].length > 0);
  const canCreate =
    !!model.fields &&
    !!model.info &&
    !model.info.isLibrary &&
    config.form?.readOnly !== true &&
    model.fields.some(f => f.name === 'Title' && isEditable(f));

  if (error) return <MessageBar messageBarType={MessageBarType.error}>{error.message}</MessageBar>;

  return (
    <div>
      <div className={styles.toolbar}>
        {config.search !== false && (
          <SearchBox
            placeholder={`Search ${model.info ? model.info.title.toLowerCase() : 'items'}`}
            value={query}
            onChange={(_, v) => setQuery(v || '')}
            styles={{ root: { width: 260 } }}
            underlined={false}
          />
        )}
        {(config.quickFilters || []).map(path => {
          const field = model.fieldMap.get(rootField(path));
          const selected = filters[path] || [];
          return (
            <Dropdown
              key={path}
              placeholder={field ? field.title : path}
              multiSelect
              options={rows ? filterOptions(model, path, rows) : []}
              selectedKeys={selected}
              onRenderTitle={() => <span>{field ? field.title : path}{selected.length ? ` · ${selected.length}` : ''}</span>}
              onChange={(_, option) => {
                if (!option) return;
                const key = String(option.key);
                setFilters({ ...filters, [path]: option.selected ? selected.concat(key) : selected.filter(k => k !== key) });
              }}
              styles={{ root: { minWidth: 130 }, title: selected.length ? { borderColor: theme.accent, color: theme.text } : undefined }}
              ariaLabel={`Filter by ${field ? field.title : path}`}
            />
          );
        })}
        {filtered && <ActionButton iconProps={{ iconName: 'ClearFilter' }} text="Clear" onClick={() => { setQuery(''); setFilters({}); }} />}
        <span className={styles.toolbarSpacer} />
        {rows && <span className={styles.count}>{filtered ? `${visible.length} of ${rows.length}` : rows.length} items</span>}
        {autotag.tagger && autotag.access && autotag.access.allowed && (
          <DefaultButton text="Auto-tag" iconProps={{ iconName: 'Tag' }} onClick={() => setTagging(true)} title={`Suggest ${autotag.tagger.field.title} with AI (${autotag.access.reason})`} />
        )}
        {canCreate && <PrimaryButton text="New item" iconProps={{ iconName: 'Add' }} onClick={onCreate} />}
      </div>
      {autotag.error && <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { marginBottom: 12 } }}>{autotag.error}</MessageBar>}

      {!rows ? (
        <LoadingRows styles={styles} />
      ) : (
        <GridTable
          rows={visible}
          columns={gridColumns}
          styles={styles}
          pageSize={config.pageSize || 25}
          rowKey={row => String(row.Id)}
          rowLabel={row => `Open ${itemKey(model, row)}`}
          onOpen={(row, ordered) => onOpen(Number(row.Id), ordered.map(r => Number(r.Id)))}
          empty={filtered ? 'No items match your search or filters.' : 'This list has no items yet.'}
          ariaLabel={model.info ? model.info.title : 'Items'}
        />
      )}

      {tagging && autotag.tagger && (
        <AutotagPanel
          model={model}
          tagger={autotag.tagger}
          visibleRows={visible}
          theme={theme}
          styles={styles}
          onDismiss={() => setTagging(false)}
          onApplied={async id => {
            model.replaceRow(await model.service.item(id, itemPaths(model)));
            onChanged();
          }}
        />
      )}
    </div>
  );
}
