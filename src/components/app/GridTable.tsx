import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { Row } from '../../core/types';
import { compare } from '../../core/expr/evaluate';
import { AppStyles } from './appStyles';

export interface GridColumn {
  key: string;
  title: string;
  width?: number;
  minWidth?: number;
  align?: 'left' | 'right' | 'center';
  primary?: boolean;
  sortValue: (row: Row) => unknown;
  render: (row: Row) => React.ReactNode;
}

export interface GridSort {
  key: string;
  desc: boolean;
}

export interface GridTableProps {
  rows: Row[];
  columns: GridColumn[];
  styles: AppStyles;
  pageSize: number;
  rowKey: (row: Row) => string;
  rowLabel?: (row: Row) => string;
  onOpen?: (row: Row, ordered: Row[]) => void;
  initialSort?: GridSort;
  empty: React.ReactNode;
  ariaLabel: string;
}

export function GridTable(props: GridTableProps): React.ReactElement {
  const { rows, columns, styles, pageSize, onOpen } = props;
  const [sort, setSort] = React.useState<GridSort | undefined>(props.initialSort);
  const [page, setPage] = React.useState(0);
  const bodyRef = React.useRef<HTMLTableSectionElement>(null);

  React.useEffect(() => setPage(0), [rows]);

  const sorted = React.useMemo(() => {
    const column = sort && columns.find(c => c.key === sort.key);
    if (!sort || !column) return rows;
    return rows.slice().sort((a, b) => {
      const c = compare(column.sortValue(a), column.sortValue(b));
      return sort.desc ? -c : c;
    });
  }, [rows, columns, sort]);

  const pages = Math.max(1, Math.ceil(sorted.length / pageSize));
  const current = Math.min(page, pages - 1);
  const visible = sorted.slice(current * pageSize, (current + 1) * pageSize);

  const toggleSort = (key: string): void => {
    setSort(s => (!s || s.key !== key ? { key, desc: false } : s.desc ? undefined : { key, desc: true }));
  };

  const focusRow = (index: number): void => {
    const body = bodyRef.current;
    if (!body) return;
    const target = body.querySelectorAll<HTMLTableRowElement>('tr[data-row]')[index];
    if (target) target.focus();
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLTableRowElement>, row: Row, index: number): void => {
    if ((e.key === 'Enter' || e.key === ' ') && onOpen) {
      e.preventDefault();
      onOpen(row, sorted);
    } else if (e.key === 'ArrowDown') {
      e.preventDefault();
      focusRow(index + 1);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      focusRow(index - 1);
    }
  };

  return (
    <div className={styles.gridCard}>
      <div className={styles.gridScroll}>
        <table className={styles.table} aria-label={props.ariaLabel} aria-rowcount={sorted.length}>
          <thead>
            <tr>
              {columns.map(c => {
                const active = sort !== undefined && sort.key === c.key;
                return (
                  <th
                    key={c.key}
                    className={styles.th}
                    style={{ width: c.width, minWidth: c.minWidth, textAlign: c.align }}
                    aria-sort={active ? (sort.desc ? 'descending' : 'ascending') : 'none'}
                    scope="col"
                  >
                    <button type="button" className={styles.thButton} style={{ justifyContent: c.align === 'right' ? 'flex-end' : 'flex-start' }} onClick={() => toggleSort(c.key)}>
                      <span>{c.title}</span>
                      <Icon iconName={active ? (sort.desc ? 'SortDown' : 'SortUp') : 'Sort'} className={active ? styles.sortActive : styles.sortIdle} />
                    </button>
                  </th>
                );
              })}
              {onOpen && <th className={styles.th} aria-hidden="true" />}
            </tr>
          </thead>
          <tbody ref={bodyRef}>
            {visible.map((row, i) => (
              <tr
                key={props.rowKey(row)}
                data-row=""
                className={onOpen ? styles.row : styles.rowStatic}
                tabIndex={onOpen ? 0 : undefined}
                onClick={onOpen ? () => onOpen(row, sorted) : undefined}
                onKeyDown={onOpen ? e => onKeyDown(e, row, i) : undefined}
                aria-label={props.rowLabel ? props.rowLabel(row) : undefined}
              >
                {columns.map(c => (
                  <td key={c.key} className={c.primary ? `${styles.td} ${styles.primaryCell}` : styles.td} style={{ textAlign: c.align }}>
                    {c.render(row)}
                  </td>
                ))}
                {onOpen && (
                  <td className={`${styles.td} ${styles.openCell}`} aria-hidden="true">
                    <Icon iconName="ChevronRight" className={styles.openIcon} />
                  </td>
                )}
              </tr>
            ))}
          </tbody>
        </table>
        {!visible.length && <div className={styles.empty}>{props.empty}</div>}
      </div>
      {sorted.length > pageSize && (
        <div className={styles.pager}>
          <span>
            {current * pageSize + 1}–{Math.min(sorted.length, (current + 1) * pageSize)} of {sorted.length}
          </span>
          <span className={styles.pagerButtons}>
            <IconButton iconProps={{ iconName: 'DoubleChevronLeft' }} title="First page" ariaLabel="First page" disabled={current === 0} onClick={() => setPage(0)} />
            <IconButton iconProps={{ iconName: 'ChevronLeft' }} title="Previous page" ariaLabel="Previous page" disabled={current === 0} onClick={() => setPage(current - 1)} />
            <span className={styles.pageNumber}>Page {current + 1} of {pages}</span>
            <IconButton iconProps={{ iconName: 'ChevronRight' }} title="Next page" ariaLabel="Next page" disabled={current >= pages - 1} onClick={() => setPage(current + 1)} />
            <IconButton iconProps={{ iconName: 'DoubleChevronRight' }} title="Last page" ariaLabel="Last page" disabled={current >= pages - 1} onClick={() => setPage(pages - 1)} />
          </span>
        </div>
      )}
    </div>
  );
}
