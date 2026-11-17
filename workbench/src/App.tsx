import * as React from 'react';
import { loadTheme } from '@fluentui/react/lib/Styling';
import { MetricsWebPartView } from '../../src/webparts/metrics/MetricsWebPartView';
import { JsonConfigEditor } from '../../src/components/JsonConfigEditor';
import { parseConfig } from '../../src/core/validate';
import type { DataContext, IHttp } from '../../src/core/datasources';
import type { MetricsConfig } from '../../src/core/types';
import { SeedData } from './SeedData';
import { SECTION_BACKGROUNDS, SITE_THEMES, SectionBackground, WIDTHS, sectionTheme, siteTheme } from './spThemes';

const SAMPLE_FILES: Record<string, string> = import.meta.glob('../../src/samples/*.json', { eager: true, query: '?raw', import: 'default' });
const SAMPLES = Object.keys(SAMPLE_FILES)
  .map(path => ({ name: path.split('/').pop()?.replace(/\.json$/, '') || path, text: SAMPLE_FILES[path] }))
  .sort((a, b) => a.name.localeCompare(b.name));

const SITE = '/sites/dev';
const STORE = 'metrics-workbench:';

function storeGet(key: string): string | undefined {
  try {
    return localStorage.getItem(STORE + key) ?? undefined;
  } catch {
    return undefined;
  }
}

function storeSet(key: string, value: string | undefined): void {
  try {
    if (value === undefined) localStorage.removeItem(STORE + key);
    else localStorage.setItem(STORE + key, value);
  } catch {
    return;
  }
}

interface StoredEdit {
  base: string;
  text: string;
}

function readEdit(name: string): StoredEdit | undefined {
  const raw = storeGet('edit:' + name);
  if (!raw) return undefined;
  try {
    const parsed: unknown = JSON.parse(raw);
    const edit = parsed as Partial<StoredEdit>;
    return typeof edit.base === 'string' && typeof edit.text === 'string' ? { base: edit.base, text: edit.text } : undefined;
  } catch {
    return undefined;
  }
}

function writeEdit(name: string, base: string, text: string): void {
  storeSet('edit:' + name, text === base ? undefined : JSON.stringify({ base, text }));
  storeSet('config:' + name, undefined);
}

function initialText(name: string, original: string): string {
  const edit = readEdit(name);
  return edit ? edit.text : original;
}

function sampleText(name: string): string {
  const sample = SAMPLES.find(s => s.name === name) || SAMPLES[0];
  return sample.text;
}

interface RequestLog {
  id: number;
  url: string;
  status: number | string;
  ms: number;
  items?: number;
}

export type MockRole = 'owner' | 'member';

function useMockSharePoint(onRequest: (log: RequestLog) => void, role: MockRole): DataContext {
  const log = React.useRef(onRequest);
  log.current = onRequest;
  return React.useMemo<DataContext>(() => {
    let sequence = 0;
    const send = async (url: string, init: RequestInit): Promise<Response> => {
      const id = ++sequence;
      const started = performance.now();
      const record = (status: number | string, items?: number): void =>
        log.current({ id, url: `${init.method || 'GET'} ${url}`, status, ms: Math.round(performance.now() - started), items });
      let response: Response;
      try {
        const headers = new Headers(init.headers);
        if (new URL(url, location.href).origin === location.origin) headers.set('X-Mock-Role', role);
        response = await fetch(url, { ...init, headers });
      } catch (e) {
        record('failed');
        throw e;
      }
      if (!response.ok) {
        record(response.status);
        const body = await response.json().catch(() => undefined);
        const error = body && (body['odata.error'] || body.error);
        throw new Error(`HTTP ${response.status} – ${error && error.message ? error.message.value || error.message : response.statusText}`);
      }
      const length = response.headers.get('X-Item-Count');
      record(response.status, length === null ? undefined : Number(length));
      return response;
    };
    const request = async (url: string, init: RequestInit): Promise<unknown> => {
      const response = await send(url, init);
      const text = await response.text();
      return text ? JSON.parse(text) : undefined;
    };
    const http: IHttp = {
      getJson: (url, options) => request(url, { headers: options.headers }),
      postJson: (url, options) =>
        request(url, { method: 'POST', headers: options.headers, body: options.body === undefined ? undefined : JSON.stringify(options.body) }),
      getBinary: async (url, options) => (await send(url, { headers: options.headers })).arrayBuffer(),
    };
    return {
      http,
      webUrl: location.origin + SITE,
      user: { displayName: 'Dev User', email: 'dev.user@contoso.dev', loginName: 'i:0#.f|membership|dev.user@contoso.dev' },
    };
  }, [role]);
}

const bar: React.CSSProperties = { display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: '8px 12px', padding: '8px 16px', minHeight: 48, background: '#1b1a19', color: '#fff', fontSize: 13, flex: '0 0 auto' };

function viewTab(active: boolean): React.CSSProperties {
  return {
    fontFamily: 'inherit',
    fontSize: 13,
    padding: '4px 12px',
    border: 'none',
    cursor: 'pointer',
    whiteSpace: 'nowrap',
    background: active ? '#0078d4' : '#323130',
    color: '#fff',
    fontWeight: active ? 600 : 400,
  };
}
const select: React.CSSProperties = { fontFamily: 'inherit', fontSize: 13, background: '#323130', color: '#fff', border: '1px solid #484644', borderRadius: 4, padding: '4px 6px' };
const button: React.CSSProperties = { ...select, cursor: 'pointer', padding: '4px 10px', whiteSpace: 'nowrap' };
const lightButton: React.CSSProperties = { ...button, background: '#fff', color: '#323130', border: '1px solid #c8c6c4' };

export function App(): React.ReactElement {
  const initialSample = decodeURIComponent(location.hash.slice(1)) || storeGet('sample') || 'project-portfolio';
  const [sample, setSample] = React.useState(SAMPLES.some(s => s.name === initialSample) ? initialSample : SAMPLES[0].name);
  const original = sampleText(sample);
  const [text, setText] = React.useState(() => initialText(sample, original));
  const [outdated, setOutdated] = React.useState(() => {
    const edit = readEdit(sample);
    return !!edit && edit.base !== original;
  });
  const previousOriginal = React.useRef(original);

  React.useEffect(() => {
    if (previousOriginal.current === original) return;
    if (text === previousOriginal.current) setText(original);
    else setOutdated(true);
    previousOriginal.current = original;
  }, [original, text]);
  const [siteIndex, setSiteIndex] = React.useState(() => Number(storeGet('site')) || 0);
  const [background, setBackground] = React.useState<SectionBackground>(() => {
    const stored = storeGet('section');
    return SECTION_BACKGROUNDS.find(b => b === stored) || 'none';
  });
  const [widthIndex, setWidthIndex] = React.useState(() => Number(storeGet('width')) || 0);
  const [showEditor, setShowEditor] = React.useState(() => storeGet('editor') !== 'hidden');
  const [editMode, setEditMode] = React.useState(false);
  const [reloadToken, setReloadToken] = React.useState(0);
  const [requests, setRequests] = React.useState<RequestLog[]>([]);
  const [showLog, setShowLog] = React.useState(false);
  const [showSeed, setShowSeed] = React.useState(() => storeGet('seed') === 'shown');
  const [role, setRole] = React.useState<MockRole>(() => (storeGet('role') === 'member' ? 'member' : 'owner'));

  const [parsed, setParsed] = React.useState(() => parseConfig(text));
  const [lastGood, setLastGood] = React.useState<MetricsConfig | undefined>(parsed.config);
  React.useEffect(() => {
    const timer = setTimeout(() => {
      const result = parseConfig(text);
      setParsed(result);
      if (result.config) setLastGood(result.config);
      const edit = readEdit(sample);
      writeEdit(sample, outdated && edit ? edit.base : original, text);
    }, 250);
    return () => clearTimeout(timer);
  }, [text, sample, original, outdated]);

  React.useEffect(() => {
    storeSet('site', String(siteIndex));
    storeSet('section', background);
    storeSet('width', String(widthIndex));
    storeSet('editor', showEditor ? 'shown' : 'hidden');
    storeSet('role', role);
    storeSet('seed', showSeed ? 'shown' : undefined);
  }, [siteIndex, background, widthIndex, showEditor, role, showSeed]);

  React.useEffect(() => {
    const hot = import.meta.hot;
    if (!hot) return;
    const onUpdate = (): void => setReloadToken(t => t + 1);
    hot.on('sp-mock:update', onUpdate);
    return () => hot.off('sp-mock:update', onUpdate);
  }, []);

  const data = useMockSharePoint(entry => setRequests(r => [entry].concat(r).slice(0, 50)), role);

  const site = React.useMemo(() => siteTheme(SITE_THEMES[siteIndex] || SITE_THEMES[0]), [siteIndex]);
  const section = React.useMemo(() => sectionTheme(site, background), [site, background]);
  React.useEffect(() => {
    loadTheme(site);
  }, [site]);

  const chooseSample = (name: string): void => {
    setSample(name);
    storeSet('sample', name);
    location.hash = encodeURIComponent(name);
    const next = sampleText(name);
    const edit = readEdit(name);
    previousOriginal.current = next;
    setOutdated(!!edit && edit.base !== next);
    setText(initialText(name, next));
  };

  const loadLatestSample = (): void => {
    setOutdated(false);
    setText(original);
  };

  const width = WIDTHS[widthIndex] || WIDTHS[0];
  const errors = parsed.diagnostics.filter(d => d.severity === 'error');
  const warnings = parsed.diagnostics.filter(d => d.severity === 'warning');
  const dark = site.isInverted;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
      <header style={bar}>
        <strong style={{ fontSize: 14, whiteSpace: 'nowrap' }}>Metrics workbench</strong>
        <div role="tablist" aria-label="Workbench view" style={{ display: 'inline-flex', border: '1px solid #484644', borderRadius: 4, overflow: 'hidden', flex: '0 0 auto' }}>
          <button type="button" role="tab" aria-selected={!showSeed} style={viewTab(!showSeed)} onClick={() => setShowSeed(false)}>Preview</button>
          <button type="button" role="tab" aria-selected={showSeed} style={viewTab(showSeed)} onClick={() => setShowSeed(true)}>Seed data</button>
        </div>
        <span style={{ color: '#a19f9d', whiteSpace: 'nowrap' }}>mock SharePoint at {SITE}</span>
        <span style={{ flex: '1 1 auto' }} />
        <label style={{ whiteSpace: 'nowrap' }}>
          Sample{' '}
          <select style={select} value={sample} onChange={e => chooseSample(e.target.value)}>
            {SAMPLES.map(s => <option key={s.name} value={s.name}>{s.name}</option>)}
          </select>
        </label>
        <label style={{ whiteSpace: 'nowrap' }}>
          Site theme{' '}
          <select style={select} value={siteIndex} onChange={e => setSiteIndex(Number(e.target.value))}>
            {SITE_THEMES.map((t, i) => <option key={t.name} value={i}>{t.name}</option>)}
          </select>
        </label>
        <label style={{ whiteSpace: 'nowrap' }}>
          Section background{' '}
          <select style={select} value={background} onChange={e => setBackground(SECTION_BACKGROUNDS.find(b => b === e.target.value) || 'none')}>
            {SECTION_BACKGROUNDS.map(b => <option key={b} value={b}>{b}</option>)}
          </select>
        </label>
        <label style={{ whiteSpace: 'nowrap' }}>
          Width{' '}
          <select style={select} value={widthIndex} onChange={e => setWidthIndex(Number(e.target.value))}>
            {WIDTHS.map((w, i) => <option key={w.name} value={i}>{w.name}</option>)}
          </select>
        </label>
        <label style={{ whiteSpace: 'nowrap' }}>
          Signed in as{' '}
          <select style={select} value={role} onChange={e => setRole(e.target.value === 'member' ? 'member' : 'owner')}>
            <option value="owner">Site owner</option>
            <option value="member">Member</option>
          </select>
        </label>
        <button type="button" style={button} onClick={() => setEditMode(!editMode)} aria-pressed={editMode}>{editMode ? 'Edit mode' : 'Read mode'}</button>
        <button type="button" style={button} onClick={() => setReloadToken(t => t + 1)}>Reload data</button>
        <button type="button" style={button} onClick={() => setShowLog(!showLog)} aria-pressed={showLog}>Requests ({requests.length})</button>
        <button type="button" style={button} onClick={() => setShowEditor(!showEditor)} aria-pressed={showEditor}>{showEditor ? 'Hide' : 'Show'} JSON</button>
      </header>

      <div style={{ flex: '1 1 auto', display: 'flex', minHeight: 0 }}>
        {showEditor && (
          <aside style={{ flex: '0 0 min(44%, 640px)', display: 'flex', flexDirection: 'column', borderRight: '1px solid #e1dfdd', minWidth: 0 }}>
            <div style={{ flex: '1 1 auto', minHeight: 0 }}>
              <JsonConfigEditor value={text} onChange={setText} dark={dark} />
            </div>
            <div style={{ flex: '0 0 auto', display: 'flex', gap: 8, alignItems: 'center', padding: '6px 10px', fontSize: 12, borderTop: '1px solid #e1dfdd', background: '#faf9f8', color: '#323130' }}>
              <span style={{ color: errors.length ? '#a4262c' : '#107c10', fontWeight: 600 }}>{errors.length ? `${errors.length} error(s)` : 'Valid'}</span>
              {warnings.length > 0 && <span style={{ color: '#8a6100' }}>{warnings.length} warning(s)</span>}
              {errors.length > 0 && (
                <span style={{ color: '#605e5c', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={errors[0].message}>{errors[0].message}</span>
              )}
              <span style={{ flex: '1 1 auto' }} />
              <button type="button" style={lightButton} disabled={!parsed.config} onClick={() => parsed.config && setText(JSON.stringify(parsed.config, undefined, 2))}>Format</button>
              <button type="button" style={lightButton} onClick={() => navigator.clipboard.writeText(text)}>Copy</button>
              <button type="button" style={lightButton} disabled={text === original} onClick={loadLatestSample}>Reset</button>
            </div>
          </aside>
        )}

        {showSeed && (
          <main style={{ flex: '1 1 auto', minWidth: 0, overflow: 'hidden' }}>
            <SeedData reloadToken={reloadToken} onReset={() => setReloadToken(t => t + 1)} />
          </main>
        )}
        <main style={{ flex: '1 1 auto', overflow: 'auto', background: section.semanticColors.bodyBackground, position: 'relative', display: showSeed ? 'none' : undefined }}>
          {outdated && (
            <div role="status" style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', padding: '8px 16px', background: '#fff4ce', color: '#323130', fontSize: 13, borderBottom: '1px solid #f2d27a' }}>
              The {sample} sample file changed since you edited it here, so you are seeing your older edited copy.
              <button type="button" style={lightButton} onClick={loadLatestSample}>Load updated sample</button>
              <button type="button" style={lightButton} onClick={() => setOutdated(false)}>Keep my edits</button>
            </div>
          )}
          <div style={{ maxWidth: width.px, margin: width.px ? '24px auto' : 24, padding: width.px ? '0 16px' : 0 }}>
            <MetricsWebPartView
              key={sample}
              config={lastGood}
              error={!lastGood && errors.length ? errors.map(e => e.message).join('\n') : undefined}
              data={data}
              sectionTheme={section}
              locale="en-US"
              reloadToken={reloadToken}
              instanceKey="wb"
              isEditMode={editMode}
              editorOpen={false}
              editorText={text}
              onOpenEditor={() => setShowEditor(true)}
              onCloseEditor={() => setShowEditor(false)}
              onSaveConfig={setText}
              onUseStarter={() => chooseSample('starter')}
            />
          </div>
          {showLog && (
            <div style={{ position: 'fixed', right: 16, bottom: 16, width: 'min(560px, calc(100vw - 32px))', maxHeight: '40vh', overflow: 'auto', background: '#fff', border: '1px solid #c8c6c4', borderRadius: 8, boxShadow: '0 8px 24px rgba(0,0,0,0.18)', fontSize: 12, zIndex: 10, color: '#323130' }}>
              <div style={{ display: 'flex', padding: '8px 12px', borderBottom: '1px solid #edebe9', fontWeight: 600 }}>
                Mock SharePoint requests
                <span style={{ flex: 1 }} />
                <button type="button" style={{ border: 'none', background: 'none', cursor: 'pointer' }} onClick={() => setRequests([])}>Clear</button>
              </div>
              {requests.map(r => (
                <div key={`${r.id}-${r.url}`} style={{ display: 'flex', gap: 8, padding: '6px 12px', borderBottom: '1px solid #f3f2f1', fontFamily: 'Consolas, monospace' }}>
                  <span style={{ color: typeof r.status === 'number' && r.status < 300 ? '#107c10' : '#a4262c', flex: '0 0 auto' }}>{r.status}</span>
                  <span style={{ flex: '1 1 auto', wordBreak: 'break-all' }}>{decodeURIComponent(r.url.replace(location.origin, ''))}</span>
                  <span style={{ flex: '0 0 auto', color: '#605e5c' }}>{r.items !== undefined ? `${r.items} items · ` : ''}{r.ms} ms</span>
                </div>
              ))}
              {!requests.length && <div style={{ padding: 12, color: '#605e5c' }}>No requests yet.</div>}
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
