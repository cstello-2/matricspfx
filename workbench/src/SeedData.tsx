import * as React from 'react';

type Item = Record<string, unknown>;

interface FieldDef {
  type: string;
  title?: string;
  choices?: string[];
  list?: string;
  field?: string;
  required?: boolean;
  readOnly?: boolean;
  dateOnly?: boolean;
  description?: string;
}

interface SeedList {
  title: string;
  id: string;
  baseTemplate: number;
  fields: Record<string, FieldDef>;
  items: Item[];
  comments: Record<string, number>;
  changed: number[];
}

interface Seed {
  lists: SeedList[];
  users: Item[];
}

type View = { kind: 'list'; title: string; tab: 'items' | 'schema' } | { kind: 'users' };

const PAGE = 100;
const SYSTEM = ['ID', 'Created', 'Modified', 'Author', 'Editor'];
const REFERENCE = ['Lookup', 'LookupMulti', 'User', 'UserMulti'];

const css = {
  root: { display: 'flex', height: '100%', minHeight: 0, background: '#faf9f8', color: '#323130', fontSize: 13 } as React.CSSProperties,
  nav: { flex: '0 0 220px', borderRight: '1px solid #e1dfdd', background: '#fff', overflowY: 'auto', padding: '12px 0' } as React.CSSProperties,
  navTitle: { padding: '4px 16px 8px', fontSize: 11, fontWeight: 600, textTransform: 'uppercase', color: '#605e5c', letterSpacing: '0.04em' } as React.CSSProperties,
  main: { flex: '1 1 auto', minWidth: 0, display: 'flex', flexDirection: 'column', padding: 16, gap: 12 } as React.CSSProperties,
  toolbar: { display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' } as React.CSSProperties,
  button: { fontFamily: 'inherit', fontSize: 13, padding: '4px 12px', borderRadius: 4, border: '1px solid #c8c6c4', background: '#fff', color: '#323130', cursor: 'pointer' } as React.CSSProperties,
  input: { fontFamily: 'inherit', fontSize: 13, padding: '5px 10px', borderRadius: 4, border: '1px solid #c8c6c4', width: 240 } as React.CSSProperties,
  tableWrap: { flex: '1 1 auto', minHeight: 0, overflow: 'auto', border: '1px solid #e1dfdd', borderRadius: 4, background: '#fff' } as React.CSSProperties,
  th: { position: 'sticky', top: 0, background: '#f3f2f1', textAlign: 'left', fontWeight: 600, padding: '6px 10px', borderBottom: '1px solid #e1dfdd', whiteSpace: 'nowrap', zIndex: 1 } as React.CSSProperties,
  td: { padding: '5px 10px', borderBottom: '1px solid #f3f2f1', whiteSpace: 'nowrap', maxWidth: 320, overflow: 'hidden', textOverflow: 'ellipsis', verticalAlign: 'top' } as React.CSSProperties,
  muted: { color: '#a19f9d' } as React.CSSProperties,
  badge: { display: 'inline-block', marginLeft: 6, padding: '0 6px', borderRadius: 8, fontSize: 11, background: '#fff4ce', color: '#8a6100' } as React.CSSProperties,
};

function navButton(active: boolean): React.CSSProperties {
  return {
    display: 'flex',
    alignItems: 'center',
    width: '100%',
    gap: 6,
    padding: '6px 16px',
    border: 'none',
    borderLeft: `3px solid ${active ? '#0078d4' : 'transparent'}`,
    background: active ? '#edebe9' : 'transparent',
    color: '#323130',
    font: 'inherit',
    textAlign: 'left',
    cursor: 'pointer',
  };
}

function tabButton(active: boolean): React.CSSProperties {
  return { ...css.button, background: active ? '#0078d4' : '#fff', color: active ? '#fff' : '#323130', borderColor: active ? '#0078d4' : '#c8c6c4' };
}

function rawText(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (Array.isArray(v)) return v.map(rawText).join(', ');
  if (typeof v === 'object') return JSON.stringify(v);
  return String(v);
}

function Cell({ value }: { value: React.ReactNode }): React.ReactElement {
  if (value === '' || value === null || value === undefined) return <span style={css.muted}>null</span>;
  return <>{value}</>;
}

export function SeedData({ reloadToken, onReset }: { reloadToken: number; onReset: () => void }): React.ReactElement {
  const [seed, setSeed] = React.useState<Seed | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [view, setView] = React.useState<View>({ kind: 'list', title: 'Tasks', tab: 'items' });
  const [query, setQuery] = React.useState('');
  const [page, setPage] = React.useState(0);
  const [changedOnly, setChangedOnly] = React.useState(false);
  const [refresh, setRefresh] = React.useState(0);
  const [inspect, setInspect] = React.useState<Item | undefined>(undefined);

  React.useEffect(() => {
    let live = true;
    fetch('/__sp-mock/seed')
      .then(r => (r.ok ? r.json() : Promise.reject(new Error(`HTTP ${r.status}`))))
      .then(
        (data: Seed) => {
          if (live) {
            setSeed(data);
            setError(undefined);
          }
        },
        (e: Error) => {
          if (live) setError(e.message);
        },
      );
    return () => {
      live = false;
    };
  }, [reloadToken, refresh]);

  React.useEffect(() => {
    setPage(0);
    setInspect(undefined);
  }, [view, query, changedOnly]);

  const reset = (): void => {
    fetch('/__sp-mock/reset', { method: 'POST' }).then(() => {
      setRefresh(r => r + 1);
      onReset();
    }, (e: Error) => setError(e.message));
  };

  const lookupName = React.useMemo(() => {
    const byTitle = new Map((seed ? seed.lists : []).map((l): [string, SeedList] => [l.title, l]));
    const users = new Map((seed ? seed.users : []).map((u): [unknown, string] => [u.Id, String(u.Title)]));
    return (def: FieldDef, id: unknown): string => {
      if (def.type === 'User' || def.type === 'UserMulti') return users.get(id) || '?';
      const target = def.list ? byTitle.get(def.list) : undefined;
      const item = target && target.items.find(it => it.Id === id);
      return item ? String(item[def.field || 'Title']) : '?';
    };
  }, [seed]);

  if (error) return <div style={{ padding: 24, color: '#a4262c' }}>Could not load seed data: {error}</div>;
  if (!seed) return <div style={{ padding: 24 }}>Loading seed data…</div>;

  const list = view.kind === 'list' ? seed.lists.find(l => l.title === view.title) || seed.lists[0] : undefined;

  let content: React.ReactNode;
  if (view.kind === 'users') {
    const columns = Array.from(new Set(seed.users.flatMap(u => Object.keys(u))));
    content = (
      <div style={css.tableWrap}>
        <table style={{ borderCollapse: 'collapse', width: '100%' }}>
          <thead><tr>{columns.map(c => <th key={c} style={css.th}>{c}</th>)}</tr></thead>
          <tbody>
            {seed.users.map(u => (
              <tr key={String(u.Id)}>{columns.map(c => <td key={c} style={css.td}><Cell value={rawText(u[c])} /></td>)}</tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  } else if (list && view.tab === 'schema') {
    content = (
      <div style={css.tableWrap}>
        <table style={{ borderCollapse: 'collapse', width: '100%' }}>
          <thead>
            <tr>{['Internal name', 'Display name', 'Type', 'Required', 'Read-only', 'Choices / target', 'Description'].map(h => <th key={h} style={css.th}>{h}</th>)}</tr>
          </thead>
          <tbody>
            {Object.keys(list.fields).map(name => {
              const f = list.fields[name];
              const target = f.choices && f.choices.length ? f.choices.join(', ') : f.list ? `${f.list}.${f.field || 'Title'}` : f.type === 'DateTime' ? (f.dateOnly ? 'date only' : 'date and time') : '';
              return (
                <tr key={name}>
                  <td style={{ ...css.td, fontFamily: 'Consolas, monospace' }}>{name}</td>
                  <td style={css.td}>{f.title || name}</td>
                  <td style={css.td}>{f.type}</td>
                  <td style={css.td}>{f.required ? 'Yes' : ''}</td>
                  <td style={css.td}>{f.readOnly ? 'Yes' : ''}</td>
                  <td style={{ ...css.td, whiteSpace: 'normal' }}><Cell value={target} /></td>
                  <td style={{ ...css.td, whiteSpace: 'normal' }}>{f.description || ''}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  } else if (list) {
    const fieldNames = Object.keys(list.fields).filter(n => SYSTEM.indexOf(n) < 0);
    const changed = new Set(list.changed);
    const q = query.toLowerCase();
    const rows = list.items
      .filter(it => !changedOnly || changed.has(Number(it.Id)))
      .filter(it => !q || Object.values(it).some(v => rawText(v).toLowerCase().indexOf(q) >= 0));
    const pages = Math.max(1, Math.ceil(rows.length / PAGE));
    const visible = rows.slice(page * PAGE, (page + 1) * PAGE);
    const value = (it: Item, name: string): React.ReactNode => {
      const def = list.fields[name];
      if (def && REFERENCE.indexOf(def.type) >= 0) {
        const raw = it[`${name}Id`];
        const ids = Array.isArray(raw) ? raw : raw === null || raw === undefined ? [] : [raw];
        if (!ids.length) return '';
        return ids.map(id => `${String(id)} · ${lookupName(def, id)}`).join(', ');
      }
      return rawText(it[name]);
    };
    content = (
      <>
        <div style={css.toolbar}>
          <input style={css.input} type="search" placeholder={`Search ${list.title}`} value={query} onChange={e => setQuery(e.target.value)} aria-label="Search seed items" />
          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
            <input type="checkbox" checked={changedOnly} onChange={e => setChangedOnly(e.target.checked)} /> Only edited items ({list.changed.length})
          </label>
          <span style={{ flex: '1 1 auto' }} />
          <span style={{ color: '#605e5c' }}>{rows.length} of {list.items.length} items</span>
          {pages > 1 && (
            <>
              <button type="button" style={css.button} disabled={page === 0} onClick={() => setPage(page - 1)}>‹</button>
              <span>{page + 1} / {pages}</span>
              <button type="button" style={css.button} disabled={page >= pages - 1} onClick={() => setPage(page + 1)}>›</button>
            </>
          )}
        </div>
        <div style={css.tableWrap}>
          <table style={{ borderCollapse: 'collapse', width: '100%' }}>
            <thead>
              <tr>
                <th style={css.th}>Id</th>
                {fieldNames.map(n => <th key={n} style={css.th} title={list.fields[n].type}>{n}</th>)}
                <th style={css.th}>Comments</th>
                <th style={css.th}>Created</th>
                <th style={css.th}>Modified</th>
                <th style={css.th}>Author</th>
                <th style={css.th}>Editor</th>
              </tr>
            </thead>
            <tbody>
              {visible.map(it => {
                const edited = changed.has(Number(it.Id));
                return (
                  <tr
                    key={String(it.Id)}
                    tabIndex={0}
                    onClick={() => setInspect(it)}
                    onKeyDown={e => {
                      if (e.key === 'Enter') setInspect(it);
                    }}
                    title="Show all fields"
                    style={{ cursor: 'pointer', background: inspect === it ? '#deecf9' : edited ? '#fffbe6' : undefined }}
                  >
                    <td style={{ ...css.td, fontFamily: 'Consolas, monospace' }}>
                      {String(it.Id)}
                      {edited && <span style={css.badge}>edited</span>}
                    </td>
                    {fieldNames.map(n => {
                      const text = value(it, n);
                      return <td key={n} style={css.td} title={typeof text === 'string' ? text : undefined}><Cell value={text} /></td>;
                    })}
                    <td style={css.td}>{list.comments[String(it.Id)] || ''}</td>
                    <td style={css.td}>{rawText(it.Created)}</td>
                    <td style={css.td}>{rawText(it.Modified)}</td>
                    <td style={css.td}>{it.AuthorId === undefined ? '' : `${String(it.AuthorId)} · ${lookupName({ type: 'User' }, it.AuthorId)}`}</td>
                    <td style={css.td}>
                      <Cell value={it.EditorId === undefined || it.EditorId === null ? '' : `${String(it.EditorId)} · ${lookupName({ type: 'User' }, it.EditorId)}`} />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </>
    );
  }

  return (
    <div style={css.root}>
      <nav style={css.nav} aria-label="Seed lists">
        <div style={css.navTitle}>Lists</div>
        {seed.lists.map(l => (
          <button
            key={l.id}
            type="button"
            style={navButton(view.kind === 'list' && list !== undefined && list.title === l.title)}
            onClick={() => setView({ kind: 'list', title: l.title, tab: view.kind === 'list' ? view.tab : 'items' })}
          >
            <span style={{ flex: '1 1 auto' }}>{l.title}</span>
            {l.changed.length > 0 && <span style={css.badge}>{l.changed.length}</span>}
            <span style={{ color: '#605e5c' }}>{l.items.length}</span>
          </button>
        ))}
        <div style={{ ...css.navTitle, paddingTop: 16 }}>Directory</div>
        <button type="button" style={navButton(view.kind === 'users')} onClick={() => setView({ kind: 'users' })}>
          <span style={{ flex: '1 1 auto' }}>Users</span>
          <span style={{ color: '#605e5c' }}>{seed.users.length}</span>
        </button>
      </nav>
      <section style={css.main}>
        <div style={css.toolbar}>
          <strong style={{ fontSize: 16 }}>{view.kind === 'users' ? 'Users' : list ? list.title : ''}</strong>
          {list && view.kind === 'list' && (
            <span style={{ color: '#605e5c' }}>
              {list.baseTemplate === 101 ? 'Document library' : 'List'} · {list.id}
            </span>
          )}
          <span style={{ flex: '1 1 auto' }} />
          {list && view.kind === 'list' && (
            <>
              <button type="button" style={tabButton(view.tab === 'items')} onClick={() => setView({ ...view, tab: 'items' })}>Items</button>
              <button type="button" style={tabButton(view.tab === 'schema')} onClick={() => setView({ ...view, tab: 'schema' })}>Schema</button>
            </>
          )}
          <button type="button" style={css.button} onClick={() => setRefresh(r => r + 1)}>Refresh</button>
          <button type="button" style={css.button} onClick={reset} title="Reload the fixture files and drop in-memory edits and comments">Discard edits</button>
        </div>
        {content}
      </section>
      {inspect && list && view.kind === 'list' && (
        <aside aria-label="Item fields" style={{ flex: '0 0 360px', borderLeft: '1px solid #e1dfdd', background: '#fff', overflowY: 'auto', padding: 16 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12 }}>
            <strong style={{ fontSize: 15, flex: '1 1 auto' }}>{list.title} · {String(inspect.Id)}</strong>
            <button type="button" style={css.button} onClick={() => setInspect(undefined)}>Close</button>
          </div>
          <table style={{ borderCollapse: 'collapse', width: '100%' }}>
            <tbody>
              {Object.keys(inspect).map(k => (
                <tr key={k}>
                  <td style={{ ...css.td, fontFamily: 'Consolas, monospace', color: '#605e5c', whiteSpace: 'nowrap' }}>{k}</td>
                  <td style={{ ...css.td, whiteSpace: 'pre-wrap', maxWidth: 'none' }}><Cell value={rawText(inspect[k])} /></td>
                </tr>
              ))}
            </tbody>
          </table>
        </aside>
      )}
    </div>
  );
}
