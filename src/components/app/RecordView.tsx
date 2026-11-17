import * as React from 'react';
import { Modal } from '@fluentui/react/lib/Modal';
import { DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { getTheme } from '@fluentui/react/lib/Styling';
import { ThemeProvider } from '@fluentui/react/lib/Theme';
import { ColumnConfig, ListViewConfig, Row } from '../../core/types';
import { themeFromSharePoint } from '../theme';
import { getAppStyles } from './appStyles';
import { FieldValue } from './fieldDisplay';

export interface RecordRoute {
  title: string;
  rows: Row[];
  index: number;
  columns: Array<{ column: ColumnConfig; title: string }>;
}

export interface RecordViewProps {
  record: RecordRoute;
  locale?: string;
  onClose: () => void;
  onMove: (index: number) => void;
}

const NO_LIST: ListViewConfig = { source: '' };

function heading(row: Row, columns: RecordRoute['columns']): string {
  for (const key of ['Title', 'title', 'Name', 'name']) {
    if (row[key] !== undefined && row[key] !== null && row[key] !== '') return String(row[key]);
  }
  const first = columns[0] ? row[columns[0].column.field] : undefined;
  return first === undefined || first === null ? 'Record' : String(first);
}

export function RecordView({ record, locale, onClose, onMove }: RecordViewProps): React.ReactElement {
  const siteTheme = getTheme();
  const theme = React.useMemo(() => themeFromSharePoint(siteTheme), [siteTheme]);
  const styles = getAppStyles(theme);
  const row = record.rows[record.index];
  const previous = record.index > 0 ? record.index - 1 : undefined;
  const next = record.index < record.rows.length - 1 ? record.index + 1 : undefined;

  return (
    <Modal
      isOpen
      onDismiss={onClose}
      isBlocking={false}
      titleAriaId="metrics-record-title"
      styles={{
        main: { position: 'fixed', top: 0, left: 0, width: '100vw', height: '100vh', maxWidth: '100vw', maxHeight: '100vh', borderRadius: 0 },
        scrollableContent: { height: '100%', maxHeight: '100vh', overflow: 'hidden' },
      }}
    >
      <ThemeProvider theme={siteTheme} applyTo="none" style={{ height: '100%' }}>
        <div className={styles.item}>
          <div className={styles.itemTop}>
            <DefaultButton iconProps={{ iconName: 'Back' }} text={`Back to ${record.title}`} onClick={onClose} styles={{ root: { flex: '0 0 auto' } }} />
            <nav className={styles.breadcrumb} aria-label="Breadcrumb">
              <span>{record.title}</span>
              <span aria-hidden="true">/</span>
              <span className={styles.keyBadge}>
                Row {record.index + 1} of {record.rows.length}
              </span>
            </nav>
            <div className={styles.itemTopActions}>
              <IconButton iconProps={{ iconName: 'ChevronUp' }} title="Previous row" ariaLabel="Previous row" disabled={previous === undefined} onClick={() => previous !== undefined && onMove(previous)} />
              <IconButton iconProps={{ iconName: 'ChevronDown' }} title="Next row" ariaLabel="Next row" disabled={next === undefined} onClick={() => next !== undefined && onMove(next)} />
              <IconButton iconProps={{ iconName: 'Cancel' }} title="Close" ariaLabel="Close" onClick={onClose} />
            </div>
          </div>
          <div className={styles.itemBody} style={{ gridTemplateColumns: 'minmax(0, 760px)', justifyContent: 'center' }}>
            <main className={styles.itemMain}>
              <h1 id="metrics-record-title" style={{ fontSize: 24, fontWeight: 600, lineHeight: '32px', margin: 0 }}>
                {heading(row, record.columns)}
              </h1>
              <section className={styles.sideCard} aria-label="Fields">
                <h2 className={styles.sideCardTitle}>Details</h2>
                <dl style={{ display: 'grid', gridTemplateColumns: 'minmax(120px, 220px) minmax(0, 1fr)', columnGap: 24, rowGap: 14, margin: 0 }}>
                  {record.columns.map(({ column, title }) => (
                    <React.Fragment key={column.field}>
                      <dt style={{ fontSize: 13, fontWeight: 600, color: theme.textSecondary }}>{title}</dt>
                      <dd style={{ margin: 0, fontSize: 14, minWidth: 0, wordBreak: 'break-word' }}>
                        <FieldValue path={column.field} field={undefined} row={row} list={NO_LIST} theme={theme} locale={locale} format={column.format} />
                      </dd>
                    </React.Fragment>
                  ))}
                </dl>
              </section>
              <p className={styles.meta}>This row is calculated or comes from a non-SharePoint source, so it can be viewed but not edited.</p>
            </main>
          </div>
        </div>
      </ThemeProvider>
    </Modal>
  );
}
