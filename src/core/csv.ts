import { ColumnConfig, Row } from './types';
import { fieldOrExpr } from './expr/evaluate';
import { formatValue } from './format';

function quote(value: string): string {
  return /[",\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

export function toCsv(rows: Row[], columns: ColumnConfig[], locale?: string): string {
  const header = columns.map(c => quote(c.title || c.field)).join(',');
  const lines = rows.map(r =>
    columns
      .map(c => {
        const value = fieldOrExpr(c.field, r);
        if (value === null || value === undefined) return '';
        if (typeof value === 'number' && !c.format) return String(value);
        return quote(formatValue(value, c.format, locale));
      })
      .join(','),
  );
  return `﻿${[header].concat(lines).join('\r\n')}\r\n`;
}
