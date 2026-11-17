import { describe, it, expect } from 'vitest';
import { SpListService, editValue, isEditable, toSpField, toVerbosePayload } from '../../src/core/sharepoint/listService';
import { parseConfig } from '../../src/core/validate';
import { IHttp, HttpPostOptions } from '../../src/core/datasources';
import { columnsFor, formFields, formLayout } from '../../src/components/app/useListModel';
import { statusLevel } from '../../src/components/app/fieldDisplay';
import { relativeTime } from '../../src/components/app/time';

const rawFields = [
  { InternalName: 'Title', Title: 'Title', TypeAsString: 'Text', Required: true },
  { InternalName: 'Status', Title: 'Status', TypeAsString: 'Choice', Choices: ['Open', 'Done'] },
  { InternalName: 'AssignedTo', Title: 'Assigned to', TypeAsString: 'User' },
  { InternalName: 'Project', Title: 'Project', TypeAsString: 'Lookup', LookupList: '{abc}', LookupField: 'Title' },
  { InternalName: 'Due', Title: 'Due', TypeAsString: 'DateTime', DisplayFormat: 0 },
  { InternalName: 'Created', Title: 'Created', TypeAsString: 'DateTime', ReadOnlyField: true },
];

function fakeHttp(): { http: IHttp; posts: Array<{ url: string; options: HttpPostOptions }>; gets: string[] } {
  const posts: Array<{ url: string; options: HttpPostOptions }> = [];
  const gets: string[] = [];
  const http: IHttp = {
    getJson: url => {
      gets.push(url);
      if (url.indexOf('ListItemEntityTypeFullName') >= 0) return Promise.resolve({ Id: 'abc', Title: 'Tasks', ListItemEntityTypeFullName: 'SP.Data.TasksListItem' });
      if (url.indexOf('/fields') >= 0) return Promise.resolve({ value: rawFields });
      if (url.indexOf('/Comments') >= 0) return Promise.resolve({ value: [{ id: '1', text: 'hi', createdDate: '2026-01-01T00:00:00Z', author: { name: 'A', email: 'a@x' } }] });
      if (/items\(\d+\)/.test(url)) return Promise.resolve({ Id: 5, Title: 'T', AssignedTo: { Id: 3, Title: 'Ann', EMail: 'ann@x' } });
      return Promise.resolve({ value: [] });
    },
    postJson: (url, options) => {
      posts.push({ url, options });
      return Promise.resolve(url.indexOf('Comments') >= 0 ? { id: '9', text: 'new', createdDate: '2026-01-02T00:00:00Z', author: { name: 'Me' } } : undefined);
    },
    getBinary: () => Promise.reject(new Error('no files')),
  };
  return { http, posts, gets };
}

describe('SharePoint list service', () => {
  it('maps field schema to kinds', () => {
    const fields = rawFields.map(toSpField);
    expect(fields.map(f => f.kind)).toEqual(['text', 'choice', 'user', 'lookup', 'datetime', 'datetime']);
    expect(fields[1].choices).toEqual(['Open', 'Done']);
    expect(fields[3].lookupList).toBe('abc');
    expect(fields[4].dateOnly).toBe(true);
    expect(isEditable(fields[5])).toBe(false);
    expect(isEditable(fields[2])).toBe(true);
    expect(isEditable(fields[3])).toBe(true);
  });

  it('builds select and expand from column paths', async () => {
    const { http } = fakeHttp();
    const service = new SpListService({ http, webUrl: 'https://t/sites/a' }, { type: 'spList', list: 'Tasks' });
    const select = await service.selectFor(['Status', 'AssignedTo', 'Project', 'Project.Code']);
    expect(select).toEqual(expect.arrayContaining(['Status', 'AssignedTo/Id', 'AssignedTo/Title', 'AssignedTo/EMail', 'Project/Id', 'Project/Title', 'Project/Code']));
  });

  it('updates items with a typed verbose MERGE', async () => {
    const { http, posts, gets } = fakeHttp();
    const service = new SpListService({ http, webUrl: 'https://t/sites/a' }, { type: 'spList', list: 'Tasks' });
    const project = toSpField(rawFields[3]);
    await service.update(5, editValue(project, 7));
    expect(posts[0].url).toBe("https://t/sites/a/_api/web/lists/getbytitle('Tasks')/items(5)");
    expect(posts[0].options.headers).toMatchObject({
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*',
      'Content-Type': 'application/json;odata=verbose',
      Accept: 'application/json;odata=verbose',
      'odata-version': '',
    });
    expect(posts[0].options.body).toEqual({ __metadata: { type: 'SP.Data.TasksListItem' }, ProjectId: 7 });
    expect(gets.some(u => u.endsWith('/fields?$filter=Hidden eq false'))).toBe(true);
  });

  it('wraps multi-value fields in typed results collections', () => {
    const fields = [
      toSpField({ InternalName: 'Tags', TypeAsString: 'MultiChoice' }),
      toSpField({ InternalName: 'Related', TypeAsString: 'LookupMulti' }),
      toSpField({ InternalName: 'Reviewers', TypeAsString: 'UserMulti' }),
    ];
    expect(toVerbosePayload({ Tags: ['a', 'b'], RelatedId: [], ReviewersId: [4], Title: 'x' }, 'SP.Data.TasksListItem', fields)).toEqual({
      __metadata: { type: 'SP.Data.TasksListItem' },
      Tags: { __metadata: { type: 'Collection(Edm.String)' }, results: ['a', 'b'] },
      RelatedId: { __metadata: { type: 'Collection(Edm.Int32)' }, results: [] },
      ReviewersId: { __metadata: { type: 'Collection(Edm.Int32)' }, results: [4] },
      Title: 'x',
    });
  });

  it('reads documents by path so % and # in names work', async () => {
    const urls: string[] = [];
    const http = { ...fakeHttp().http, getBinary: (url: string) => { urls.push(url); return Promise.resolve(new ArrayBuffer(0)); } };
    const service = new SpListService({ http, webUrl: 'https://t/sites/a' }, { type: 'spList', list: 'Docs' });
    await service.fileContent("/sites/a/Docs/100% #1 Bob's.docx");
    expect(decodeURIComponent(urls[0])).toBe("https://t/sites/a/_api/web/GetFileByServerRelativePath(decodedurl='/sites/a/Docs/100% #1 Bob''s.docx')/$value");
  });

  it('reads and adds comments', async () => {
    const { http, posts } = fakeHttp();
    const service = new SpListService({ http, webUrl: 'https://t/sites/a' }, { type: 'spList', list: 'Tasks' });
    expect(await service.comments(5)).toEqual([{ id: '1', text: 'hi', createdDate: '2026-01-01T00:00:00Z', authorName: 'A', authorEmail: 'a@x' }]);
    const added = await service.addComment(5, 'new');
    expect(posts[0].url).toMatch(/items\(5\)\/Comments\(\)$/);
    expect(posts[0].options.body).toEqual({ text: 'new' });
    expect(added.authorName).toBe('Me');
  });

  it('converts edited values to SharePoint payloads', () => {
    const fields = rawFields.map(toSpField);
    expect(editValue(fields[1], 'Done')).toEqual({ Status: 'Done' });
    expect(editValue(fields[3], null)).toEqual({ ProjectId: null });
    expect(editValue(fields[4], '2026-03-01T00:00:00.000Z')).toEqual({ Due: '2026-03-01T00:00:00.000Z' });
  });
});

describe('list view config', () => {
  const fields = rawFields.map(toSpField);

  it('keeps configured column order and derives form fields', () => {
    const list = { source: 's', columns: ['Status', { field: 'Title', title: 'Name' }, 'AssignedTo'], form: { status: 'Status' } };
    const columns = columnsFor(list, fields);
    expect(columns.map(c => c.field)).toEqual(['Status', 'Title', 'AssignedTo']);
    expect(formFields(list, columns)).toEqual(['Status', 'AssignedTo']);
  });

  it('builds main and side sections from the form config', () => {
    const list = {
      source: 's',
      columns: ['Title', 'Status'],
      form: {
        status: 'Status',
        main: ['Description', { title: 'Effort', columns: 2 as const, fields: ['Estimate', 'Actual'] }, 'Notes'],
        side: ['AssignedTo', 'Project.Code', { title: 'Dates', fields: ['Due', 'AssignedTo'] }, 'Priority'],
      },
    };
    const layout = formLayout(list, columnsFor(list, fields));
    expect(layout.main).toEqual([
      { title: undefined, fields: ['Description'], columns: 1 },
      { title: 'Effort', fields: ['Estimate', 'Actual'], columns: 2 },
      { title: undefined, fields: ['Notes'], columns: 1 },
    ]);
    expect(layout.side).toEqual([
      { title: 'Details', fields: ['Status', 'AssignedTo', 'Project'], columns: 1 },
      { title: 'Dates', fields: ['Due'], columns: 1 },
      { title: undefined, fields: ['Priority'], columns: 1 },
    ]);
  });

  it('keeps the status where the config places it', () => {
    const list = { source: 's', form: { status: 'Status', main: ['Status'], side: ['AssignedTo'] } };
    const layout = formLayout(list, columnsFor(list, fields));
    expect(layout.main[0].fields).toEqual(['Status']);
    expect(layout.side).toEqual([{ title: 'Details', fields: ['AssignedTo'], columns: 1 }]);
  });

  it('maps the older description and fields options onto the layout', () => {
    const list = { source: 's', form: { description: 'Description', fields: ['AssignedTo', 'Due'] } };
    const layout = formLayout(list, columnsFor(list, fields));
    expect(layout.main).toEqual([{ title: undefined, fields: ['Description'], columns: 1 }]);
    expect(layout.side).toEqual([{ title: 'Details', fields: ['AssignedTo', 'Due'], columns: 1 }]);
  });

  it('rejects malformed form sections', () => {
    const base = { dataSources: { s: { type: 'spList', list: 'S' } }, widgets: [] };
    const bad = JSON.stringify({ ...base, list: { source: 's', form: { main: [{ title: 'x' }], side: [{ fields: ['A'], columns: 3 }], sideWidth: 9000 } } });
    const messages = parseConfig(bad).diagnostics.map(d => d.message);
    expect(messages).toContain('A section needs "fields": an array of field names');
    expect(messages).toContain('"columns" must be 1 or 2');
    expect(messages).toContain('"sideWidth" must be a number of pixels between 260 and 640');
  });

  it('falls back to Title plus useful fields', () => {
    expect(columnsFor({ source: 's' }, fields).map(c => c.field)).toEqual(['Title', 'Status', 'AssignedTo', 'Project', 'Due']);
  });

  it('validates the list section', () => {
    const base = { dataSources: { t: { type: 'spList', list: 'Tasks' }, i: { type: 'inline', rows: [] } }, widgets: [] };
    const msgs = (list: unknown): string[] => parseConfig(JSON.stringify({ ...base, list })).diagnostics.map(d => d.message);
    expect(msgs({ source: 't', columns: ['Title'] })).toEqual([]);
    expect(msgs({ source: 'nope' })).toEqual(['Unknown data source "nope"']);
    expect(msgs({ source: 'i' })).toEqual(['"i" must be an spList data source to show it as a list']);
    expect(msgs({ source: 't', statusColors: { Done: 'green' } })[0]).toMatch(/Level must be one of/);
  });

  it('colors statuses from config, then common names', () => {
    expect(statusLevel('Done', undefined)).toBe('good');
    expect(statusLevel('In progress', undefined)).toBe('info');
    expect(statusLevel('Waiting on user', { 'Waiting on user': 'serious' })).toBe('serious');
    expect(statusLevel('Triage', undefined)).toBe('neutral');
  });

  it('formats relative comment times', () => {
    const now = new Date('2026-11-16T12:00:00Z');
    expect(relativeTime('2026-11-16T10:00:00Z', 'en-US', now)).toBe('2 hours ago');
    expect(relativeTime('2026-11-13T12:00:00Z', 'en-US', now)).toBe('3 days ago');
  });
});
