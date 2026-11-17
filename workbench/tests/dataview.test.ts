import { describe, it, expect } from 'vitest';
import { toCsv } from '../../src/core/csv';
import { parseConfig } from '../../src/core/validate';
import { dataViewSources, dataViewTitle, inferColumns } from '../../src/components/app/DataView';
import { MetricsConfig } from '../../src/core/types';

const config: MetricsConfig = {
  dataSources: { tasks: { type: 'spList', list: 'Team Tasks' }, rates: { type: 'inline', rows: [] } },
  datasets: { openTasks: { from: 'tasks' } },
  widgets: [],
};

describe('data tab', () => {
  it('shows every data source by default, titled by list name', () => {
    const sources = dataViewSources(config);
    expect(sources.map(s => s.name)).toEqual(['tasks', 'rates']);
    expect(sources.map(s => dataViewTitle(s, config))).toEqual(['Team Tasks', 'Rates']);
  });

  it('uses configured sources, titles and datasets', () => {
    const custom = { ...config, dataView: { sources: ['openTasks', { name: 'tasks', title: 'All tasks' }] } };
    expect(dataViewSources(custom).map(s => dataViewTitle(s, custom))).toEqual(['Open Tasks', 'All tasks']);
  });

  it('infers columns, skipping SharePoint system fields and nested collections', () => {
    const cols = inferColumns([{ 'odata.etag': 'x', Id: 1, ID: 1, Title: 'A', AuthorId: 3, AssignedTo: { Title: 'Ann' }, t: { tasks: 3 }, Subtasks: [{ Id: 2 }], Tags: ['a'] }]);
    expect(cols.map(c => c.field)).toEqual(['Id', 'Title', 'AssignedTo', 'Tags']);
  });

  it('validates sources and allows hiding the tab', () => {
    const msgs = (dataView: unknown): string[] => parseConfig(JSON.stringify({ ...config, dataView })).diagnostics.map(d => d.message);
    expect(msgs(false)).toEqual([]);
    expect(msgs({ sources: ['tasks', 'openTasks'] })).toEqual([]);
    expect(msgs({ sources: ['nope'] })).toEqual(['Unknown dataset or data source "nope"']);
  });
});

describe('CSV export', () => {
  it('quotes values, formats them and adds a BOM for Excel', () => {
    const csv = toCsv(
      [
        { Title: 'Plan, "phase 1"', Hours: 4.5, Owner: { Title: 'Ann' }, Due: '2026-03-01', Budget: 1200 },
        { Title: 'Line\nbreak', Hours: null, Owner: null, Due: null, Budget: null },
      ],
      [{ field: 'Title' }, { field: 'Hours' }, { field: 'Owner', title: 'Owner name' }, { field: 'Due', format: 'date' }, { field: 'Budget', format: { type: 'currency', currency: 'EUR' } }],
      'en-US',
    );
    expect(csv).toBe('﻿Title,Hours,Owner name,Due,Budget\r\n"Plan, ""phase 1""",4.5,Ann,"Mar 1, 2026","€1,200"\r\n"Line\nbreak",,,,\r\n');
  });
});
