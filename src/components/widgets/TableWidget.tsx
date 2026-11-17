import * as React from 'react';
import { ColumnConfig, Row, TableWidgetConfig } from '../../core/types';
import { compare, evaluate, fieldOrExpr, toNumber, isNil } from '../../core/expr/evaluate';
import { formatValue } from '../../core/format';
import { useDashboard, useRows } from '../context';
import { ErrorBox, Loading, WidgetFrame } from '../WidgetFrame';
import { matchThreshold, StatusIcon } from '../status';

const SP_SYSTEM_FIELD = /^(odata\.|@|__|OData__|AuthorId$|EditorId$|ContentTypeId$|ServerRedirectedEmbed|FileSystemObjectType$|GUID$|ComplianceAssetId$)/;

function normCols(cols: Array<string | ColumnConfig> | undefined, rows: Row[]): ColumnConfig[] {
  if (cols && cols.length) return cols.map(c => (typeof c === 'string' ? { field: c } : c));
  const names: string[] = [];
  for (const r of rows.slice(0, 20)) {
    for (const k of Object.keys(r)) {
      if (names.indexOf(k) >= 0 || SP_SYSTEM_FIELD.test(k) || Array.isArray(r[k])) continue;
      if (k === 'ID' && 'Id' in r) continue;
      names.push(k);
    }
  }
  return names.slice(0, 8).map(field => ({ field }));
}

function isNumericCol(rows: Row[], field: string): boolean {
  let seen = 0;
  for (const r of rows.slice(0, 50)) {
    const v = fieldOrExpr(field, r);
    if (isNil(v) || v === '') continue;
    if (typeof v !== 'number') return false;
    seen++;
  }
  return seen > 0;
}

function headerText(c: ColumnConfig): string {
  if (c.title) return c.title;
  const last = c.field.split('.').pop() || c.field;
  return last.replace(/_x0020_/g, ' ').replace(/([a-z])([A-Z])/g, '$1 $2').replace(/^./, s => s.toUpperCase());
}

function safeUrl(v: unknown): string | undefined {
  return typeof v === 'string' && /^(https?:\/\/|\/)/i.test(v) ? v : undefined;
}

export interface DataTableProps {
  rows: Row[];
  columns?: Array<string | ColumnConfig>;
  pageSize?: number;
  search?: boolean;
  maxHeight?: number;
  renderDetail?: (row: Row) => React.ReactNode;
}

interface SortState {
  index: number;
  desc: boolean;
}

export function DataTable(props: DataTableProps): React.ReactElement {
  const { styles, theme, locale, engine } = useDashboard();
  const { rows, renderDetail } = props;
  const cols = React.useMemo(() => normCols(props.columns, rows), [props.columns, rows]);
  const numeric = React.useMemo(() => cols.map(c => isNumericCol(rows, c.field)), [cols, rows]);
  const maxes = React.useMemo(
    () => cols.map(c => (c.bar ? Math.max(0, ...rows.map(r => toNumber(fieldOrExpr(c.field, r))).filter(n => !isNaN(n))) : 0)),
    [cols, rows],
  );
  const [sort, setSort] = React.useState<SortState | undefined>(undefined);
  const [query, setQuery] = React.useState('');
  const [page, setPage] = React.useState(0);
  const [open, setOpen] = React.useState<Record<number, boolean>>({});
  const pageSize = props.pageSize || 10;

  const indexed = React.useMemo(() => rows.map((r, i) => ({ r, i })), [rows]);
  const filtered = React.useMemo(() => {
    if (!query) return indexed;
    const q = query.toLowerCase();
    return indexed.filter(({ r }) => cols.some(c => formatValue(fieldOrExpr(c.field, r), c.format, locale).toLowerCase().indexOf(q) >= 0));
  }, [indexed, query, cols, locale]);
  const sorted = React.useMemo(() => {
    if (!sort) return filtered;
    const field = cols[sort.index].field;
    return filtered.slice().sort((a, b) => {
      const c = compare(fieldOrExpr(field, a.r), fieldOrExpr(field, b.r));
      return sort.desc ? -c : c;
    });
  }, [filtered, sort, cols]);

  const pages = Math.max(1, Math.ceil(sorted.length / pageSize));
  const current = Math.min(page, pages - 1);
  const visible = sorted.slice(current * pageSize, current * pageSize + pageSize);

  const nextSort = (index: number): SortState | undefined => {
    if (!sort || sort.index !== index) return { index, desc: false };
    return sort.desc ? undefined : { index, desc: true };
  };

  const alignOf = (c: ColumnConfig, i: number): 'left' | 'right' | 'center' => c.align || (numeric[i] ? 'right' : 'left');

  const cell = (c: ColumnConfig, ci: number, r: Row): React.ReactNode => {
    const v = fieldOrExpr(c.field, r);
    const text = formatValue(v, c.format, locale);
    let content: React.ReactNode = text;
    if (c.link) {
      const href = safeUrl(evaluate(c.link, { row: r, vars: { ...engine.vars, value: v } }));
      if (href) content = <a className={styles.link} href={href} target="_blank" rel="noopener noreferrer">{text}</a>;
    }
    if (c.status) {
      const level = matchThreshold(c.status, v, undefined, r);
      if (level) {
        content = (
          <span className={styles.statusCell}>
            <StatusIcon level={level.level} color={theme.status[level.level]} />
            {level.label || content}
          </span>
        );
      }
    }
    if (c.bar) {
      const n = toNumber(v);
      const pct = maxes[ci] > 0 && !isNaN(n) ? Math.max(0, Math.min(100, (n / maxes[ci]) * 100)) : 0;
      content = (
        <span className={styles.bar}>
          <span>{content}</span>
          <span className={styles.barTrack}><span className={styles.barFill} style={{ width: `${pct}%` }} /></span>
        </span>
      );
    }
    return content;
  };

  return (
    <div>
      {props.search && (
        <div className={styles.tableTools}>
          <input
            className={styles.search}
            type="search"
            placeholder="Search"
            value={query}
            onChange={e => { setQuery(e.target.value); setPage(0); }}
            aria-label="Search table"
          />
        </div>
      )}
      <div className={styles.tableWrap} style={{ maxHeight: props.maxHeight }}>
        <table className={styles.table}>
          <thead>
            <tr>
              {renderDetail && <th className={styles.th} aria-label="Details" />}
              {cols.map((c, i) => {
                const active = sort !== undefined && sort.index === i;
                return (
                  <th
                    key={c.field}
                    className={styles.th}
                    style={{ textAlign: alignOf(c, i), width: c.width }}
                    aria-sort={active ? (sort.desc ? 'descending' : 'ascending') : 'none'}
                    tabIndex={0}
                    onClick={() => setSort(nextSort(i))}
                    onKeyDown={e => {
                      if (e.key === 'Enter' || e.key === ' ') {
                        e.preventDefault();
                        setSort(nextSort(i));
                      }
                    }}
                  >
                    {headerText(c)}
                    {active && <span className={styles.sortMark} aria-hidden="true">{sort.desc ? '↓' : '↑'}</span>}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {visible.map(({ r, i }) => (
              <React.Fragment key={i}>
                <tr className={styles.row}>
                  {renderDetail && (
                    <td className={`${styles.td} ${styles.expandCell}`}>
                      <button
                        type="button"
                        className={styles.iconButton}
                        onClick={() => setOpen({ ...open, [i]: !open[i] })}
                        aria-expanded={!!open[i]}
                        aria-label={open[i] ? 'Collapse row' : 'Expand row'}
                      >
                        <span className={open[i] ? styles.chevronOpen : styles.chevron}>›</span>
                      </button>
                    </td>
                  )}
                  {cols.map((c, ci) => (
                    <td key={c.field} className={numeric[ci] ? `${styles.td} ${styles.num}` : styles.td} style={{ textAlign: alignOf(c, ci) }}>
                      {cell(c, ci, r)}
                    </td>
                  ))}
                </tr>
                {renderDetail && open[i] && (
                  <tr>
                    <td colSpan={cols.length + 1} className={styles.detailCell}>{renderDetail(r)}</td>
                  </tr>
                )}
              </React.Fragment>
            ))}
            {!visible.length && (
              <tr>
                <td className={styles.td} colSpan={cols.length + (renderDetail ? 1 : 0)}>No rows</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      {pages > 1 && (
        <div className={styles.pager}>
          <span>{current * pageSize + 1}–{Math.min(sorted.length, (current + 1) * pageSize)} of {sorted.length}</span>
          <span>
            <button type="button" className={styles.iconButton} disabled={current === 0} onClick={() => setPage(current - 1)}>‹ Prev</button>
            <button type="button" className={styles.iconButton} disabled={current >= pages - 1} onClick={() => setPage(current + 1)}>Next ›</button>
          </span>
        </div>
      )}
    </div>
  );
}

export function TableWidget({ w }: { w: TableWidgetConfig }): React.ReactElement {
  const { styles } = useDashboard();
  const { rows, error } = useRows(w.dataset);
  const detailRows = useRows(w.detail ? w.detail.dataset : undefined);
  const detail = w.detail;

  const renderDetail = React.useMemo(() => {
    if (!detail) return undefined;
    return (row: Row): React.ReactNode => {
      if (detailRows.error) return <ErrorBox error={detailRows.error} />;
      if (!detailRows.rows) return <Loading height={40} />;
      const key = String(fieldOrExpr(detail.on[0], row));
      const children = detailRows.rows.filter(c => {
        const v = fieldOrExpr(detail.on[1], c);
        return Array.isArray(v) ? v.some(x => String(x) === key) : String(v) === key;
      });
      return (
        <div>
          {detail.title && <p className={styles.detailTitle}>{detail.title} ({children.length})</p>}
          <DataTable rows={children} columns={detail.columns} pageSize={5} />
        </div>
      );
    };
  }, [detail, detailRows.rows, detailRows.error, styles]);

  let body: React.ReactNode;
  if (error) body = <ErrorBox error={error} />;
  else if (!rows) body = <Loading height={w.height || 160} />;
  else body = <DataTable rows={rows} columns={w.columns} pageSize={w.pageSize} search={w.search} maxHeight={w.height} renderDetail={renderDetail} />;

  return <WidgetFrame title={w.title} subtitle={w.subtitle}>{body}</WidgetFrame>;
}
