import { parse as parseJsonc, parseTree, findNodeAtLocation, ParseError, printParseErrorCode, Node as JsonNode } from 'jsonc-parser';
import { MetricsConfig, DatasetConfig, DatasetRef, WidgetConfig, ChartType } from './types';
import { parse as parseExpr, functionsUsed, ExpressionError } from './expr/parser';
import { isKnownFunction } from './expr/evaluate';

export interface Diagnostic {
  path: Array<string | number>;
  message: string;
  severity: 'error' | 'warning';
  from?: number;
  to?: number;
}

export interface ParseResult {
  config?: MetricsConfig;
  diagnostics: Diagnostic[];
}

const CHART_TYPES: ChartType[] = ['bar', 'column', 'line', 'area', 'pie', 'donut', 'scatter', 'gauge', 'funnel', 'heatmap', 'sankey', 'treemap', 'radar'];
const WIDGET_TYPES = ['kpi', 'chart', 'table', 'text'];
const SOURCE_TYPES = ['spList', 'inline', 'rest'];

export function parseConfig(text: string): ParseResult {
  const errors: ParseError[] = [];
  const value = parseJsonc(text, errors, { allowTrailingComma: true, disallowComments: false });
  const diagnostics: Diagnostic[] = errors.map(e => ({
    path: [],
    message: `JSON syntax: ${printParseErrorCode(e.error)}`,
    severity: 'error' as const,
    from: e.offset,
    to: e.offset + Math.max(1, e.length),
  }));
  if (errors.length) return { diagnostics };
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return { diagnostics: [{ path: [], message: 'Config must be a JSON object', severity: 'error', from: 0, to: text.length }] };
  }
  const config = value as MetricsConfig;
  const tree = parseTree(text, [], { allowTrailingComma: true });
  const semantic = validateConfig(config).map(d => {
    const node = tree ? nodeFor(tree, d.path) : undefined;
    return node ? { ...d, from: node.offset, to: node.offset + node.length } : { ...d, from: 0, to: 1 };
  });
  return { config, diagnostics: semantic };
}

function nodeFor(tree: JsonNode, path: Array<string | number>): JsonNode | undefined {
  for (let i = path.length; i >= 0; i--) {
    const n = findNodeAtLocation(tree, path.slice(0, i));
    if (n) {
      if (i < path.length && n.type === 'object') return n.parent && n.parent.type === 'property' ? (n.parent.children ? n.parent.children[0] : n) : n;
      return n;
    }
  }
  return undefined;
}

export function validateConfig(c: MetricsConfig): Diagnostic[] {
  const out: Diagnostic[] = [];
  const err = (path: Array<string | number>, message: string): void => { out.push({ path, message, severity: 'error' }); };
  const warn = (path: Array<string | number>, message: string): void => { out.push({ path, message, severity: 'warning' }); };

  const checkExpr = (path: Array<string | number>, expr: unknown): void => {
    if (typeof expr !== 'string') { err(path, 'Expected an expression string'); return; }
    try {
      const fns = functionsUsed(parseExpr(expr));
      for (const f of fns) if (!isKnownFunction(f)) err(path, `Unknown function "${f}()"`);
    } catch (e) {
      err(path, e instanceof ExpressionError ? e.message : String(e));
    }
  };

  const sources = c.dataSources && typeof c.dataSources === 'object' ? c.dataSources : undefined;
  const datasets = c.datasets && typeof c.datasets === 'object' ? c.datasets : {};
  if (!sources) err(['dataSources'], '"dataSources" is required (an object of named sources)');

  const known = (name: string): boolean => !!((sources && name in sources) || name in datasets);

  if (sources) {
    for (const name of Object.keys(sources)) {
      const s = sources[name] as unknown as Record<string, unknown>;
      const p = ['dataSources', name];
      if (!s || typeof s !== 'object') { err(p, 'Expected an object'); continue; }
      if (SOURCE_TYPES.indexOf(s.type as string) < 0) { err([...p, 'type'], `"type" must be one of ${SOURCE_TYPES.join(', ')}`); continue; }
      if (s.type === 'spList' && !s.list && !s.listId) err(p, 'spList needs "list" (title) or "listId"');
      if (s.type === 'spList' && s.select !== undefined && !Array.isArray(s.select)) err([...p, 'select'], '"select" must be an array of field names');
      if (s.type === 'inline' && !Array.isArray(s.rows)) err([...p, 'rows'], 'inline source needs a "rows" array');
      if (s.type === 'rest' && typeof s.url !== 'string') err([...p, 'url'], 'rest source needs a "url"');
      if (name in datasets) warn(p, `"${name}" is also a dataset name; the dataset wins`);
    }
  }

  const checkDataset = (p: Array<string | number>, ds: DatasetConfig): void => {
    if (!ds || typeof ds !== 'object') { err(p, 'Expected a dataset object'); return; }
    const froms = Array.isArray(ds.from) ? ds.from : [ds.from];
    if (!ds.from) err(p, '"from" is required');
    else froms.forEach((f, i) => { if (!known(f)) err(Array.isArray(ds.from) ? [...p, 'from', i] : [...p, 'from'], `Unknown dataset or data source "${f}"`); });
    (ds.join || []).forEach((j, i) => {
      const jp = [...p, 'join', i];
      if (!known(j.source)) err([...jp, 'source'], `Unknown dataset or data source "${j.source}"`);
      if (!Array.isArray(j.on) || j.on.length !== 2) err([...jp, 'on'], '"on" must be [leftField, rightField]');
    });
    if (ds.filter !== undefined) checkExpr([...p, 'filter'], ds.filter);
    if (ds.having !== undefined) checkExpr([...p, 'having'], ds.having);
    const compute = ds.compute || {};
    const aggregate = ds.aggregate || {};
    Object.keys(compute).forEach(k => checkExpr([...p, 'compute', k], compute[k]));
    Object.keys(aggregate).forEach(k => checkExpr([...p, 'aggregate', k], aggregate[k]));
    (ds.groupBy || []).forEach((g, i) => checkExpr([...p, 'groupBy', i], typeof g === 'string' ? g : g.field));
  };

  for (const name of Object.keys(datasets)) checkDataset(['datasets', name], datasets[name]);

  const visiting: Record<string, number> = {};
  const visit = (name: string): boolean => {
    if (visiting[name] === 1) return true;
    if (visiting[name] === 2 || !(name in datasets)) return false;
    visiting[name] = 1;
    const ds = datasets[name];
    const deps = (Array.isArray(ds.from) ? ds.from : [ds.from]).concat((ds.join || []).map(j => j.source));
    for (const d of deps) if (visit(d)) { err(['datasets', name], `Dataset "${name}" depends on itself`); visiting[name] = 2; return false; }
    visiting[name] = 2;
    return false;
  };
  Object.keys(datasets).forEach(visit);

  const checkRef = (p: Array<string | number>, ref: DatasetRef | undefined, required: boolean): void => {
    if (ref === undefined) { if (required) err(p, '"dataset" is required'); return; }
    if (typeof ref === 'string') { if (!known(ref)) err(p, `Unknown dataset or data source "${ref}"`); }
    else checkDataset(p, ref);
  };

  if (c.list !== undefined) validateList(c, sources, err, checkExpr);

  const dataView = c.dataView;
  if (dataView !== undefined && dataView !== false) {
    if (!dataView || typeof dataView !== 'object') err(['dataView'], '"dataView" must be an object or false');
    else if (dataView.sources !== undefined) {
      if (!Array.isArray(dataView.sources)) err(['dataView', 'sources'], '"sources" must be a list of data source or dataset names');
      else {
        dataView.sources.forEach((src, i) => {
          const name = typeof src === 'string' ? src : src && src.name;
          if (!name) err(['dataView', 'sources', i], 'Each entry needs a data source or dataset name');
          else if (!known(name)) err(['dataView', 'sources', i], `Unknown dataset or data source "${name}"`);
        });
      }
    }
  }

  if (!Array.isArray(c.widgets)) {
    err(['widgets'], '"widgets" must be an array');
    return out;
  }
  c.widgets.forEach((w: WidgetConfig, i) => {
    const p = ['widgets', i];
    if (!w || typeof w !== 'object') { err(p, 'Expected a widget object'); return; }
    if (WIDGET_TYPES.indexOf(w.type) < 0) { err([...p, 'type'], `"type" must be one of ${WIDGET_TYPES.join(', ')}`); return; }
    checkRef([...p, 'dataset'], w.dataset, w.type !== 'text');
    if (w.span !== undefined && (typeof w.span !== 'number' || w.span < 1)) err([...p, 'span'], '"span" must be a positive number');
    switch (w.type) {
      case 'kpi':
        if (!w.value) err(p, 'kpi needs "value", e.g. "sum(Hours)"');
        else checkExpr([...p, 'value'], w.value);
        if (w.compare) { checkExpr([...p, 'compare', 'value'], w.compare.value); if (w.compare.dataset) checkRef([...p, 'compare', 'dataset'], w.compare.dataset, false); }
        if (typeof w.target === 'string') checkExpr([...p, 'target'], w.target);
        (w.thresholds || []).forEach((t, ti) => checkExpr([...p, 'thresholds', ti, 'when'], t.when));
        if (w.sparkline && w.sparkline.dataset) checkRef([...p, 'sparkline', 'dataset'], w.sparkline.dataset, false);
        break;
      case 'chart':
        if (CHART_TYPES.indexOf(w.chart) < 0) { err([...p, 'chart'], `"chart" must be one of ${CHART_TYPES.join(', ')}`); break; }
        if (w.chart === 'sankey') { if (!w.from || !w.to) err(p, 'sankey needs "from", "to" and "value" fields'); }
        else if (w.chart === 'heatmap') { if (!w.x || !w.series || !w.value) err(p, 'heatmap needs "x", "series" and "value" fields'); }
        else if (w.chart === 'gauge') { if (!w.value && !w.y) err(p, 'gauge needs "value" (aggregate expression)'); }
        else if (!w.y) err(p, `${w.chart} chart needs "y" (value field or fields)`);
        if (w.chart !== 'gauge' && w.chart !== 'sankey' && w.chart !== 'heatmap' && w.chart !== 'scatter' && !w.x) warn(p, 'No "x" field: rows will be labelled by position');
        break;
      case 'table':
        if (w.detail) {
          checkRef([...p, 'detail', 'dataset'], w.detail.dataset, true);
          if (!Array.isArray(w.detail.on) || w.detail.on.length !== 2) err([...p, 'detail', 'on'], '"on" must be [masterField, detailField]');
        }
        break;
      case 'text':
        if (typeof w.text !== 'string') err(p, 'text widget needs "text"');
        else {
          const re = /\{\{([\s\S]*?)\}\}/g;
          let m: RegExpExecArray | null;
          while ((m = re.exec(w.text))) checkExpr([...p, 'text'], m[1].trim());
        }
        break;
    }
  });
  return out;
}

const LEVELS = ['good', 'warning', 'serious', 'critical', 'neutral', 'info'];

function validateList(
  c: MetricsConfig,
  sources: MetricsConfig['dataSources'] | undefined,
  err: (path: Array<string | number>, message: string) => void,
  checkExpr: (path: Array<string | number>, expr: unknown) => void,
): void {
  const list = c.list;
  const p = ['list'];
  if (!list || typeof list !== 'object') {
    err(p, '"list" must be an object');
    return;
  }
  const source = sources && typeof list.source === 'string' ? sources[list.source] : undefined;
  if (!list.source) err(p, 'list needs "source": the name of an spList data source');
  else if (!source) err([...p, 'source'], `Unknown data source "${list.source}"`);
  else if (source.type !== 'spList') err([...p, 'source'], `"${list.source}" must be an spList data source to show it as a list`);
  if (list.columns !== undefined && !Array.isArray(list.columns)) err([...p, 'columns'], '"columns" must be an array of field names or column objects');
  (list.columns || []).forEach((col, i) => {
    const field = typeof col === 'string' ? col : col && col.field;
    if (!field || typeof field !== 'string') err([...p, 'columns', i], 'Each column needs a field name');
  });
  if (list.quickFilters !== undefined && !Array.isArray(list.quickFilters)) err([...p, 'quickFilters'], '"quickFilters" must be an array of field names');
  if (list.filter !== undefined) checkExpr([...p, 'filter'], list.filter);
  if (list.pageSize !== undefined && (typeof list.pageSize !== 'number' || list.pageSize < 1)) err([...p, 'pageSize'], '"pageSize" must be a positive number');
  const colors = list.statusColors || {};
  Object.keys(colors).forEach(k => {
    if (LEVELS.indexOf(colors[k]) < 0) err([...p, 'statusColors', k], `Level must be one of ${LEVELS.join(', ')}`);
  });
  const autotag = list.autotag;
  if (autotag !== undefined) {
    const ap = [...p, 'autotag'];
    if (!autotag || typeof autotag !== 'object') err(ap, '"autotag" must be an object');
    else {
      if (!autotag.field) err(ap, 'autotag needs "field": the column that holds the tags');
      if (!autotag.api || typeof autotag.api !== 'object') err(ap, 'autotag needs "api": { "baseUrl", "embeddingModel" }');
      else {
        if (!autotag.api.baseUrl) err([...ap, 'api'], 'api needs "baseUrl", e.g. "https://ai.example.com/v1"');
        if (!autotag.api.embeddingModel) err([...ap, 'api'], 'api needs "embeddingModel"');
      }
      if (autotag.mode !== undefined && autotag.mode !== 'review' && autotag.mode !== 'auto') err([...ap, 'mode'], '"mode" must be "review" or "auto"');
      if (autotag.minScore !== undefined && (typeof autotag.minScore !== 'number' || autotag.minScore < -1 || autotag.minScore > 1)) err([...ap, 'minScore'], '"minScore" must be a number between -1 and 1');
      if (autotag.maxTags !== undefined && (typeof autotag.maxTags !== 'number' || autotag.maxTags < 1)) err([...ap, 'maxTags'], '"maxTags" must be at least 1');
      if (autotag.allow !== undefined && (!Array.isArray(autotag.allow) || autotag.allow.some(a => typeof a !== 'string'))) {
        err([...ap, 'allow'], '"allow" must be a list of "owners", "listOwners", emails or SharePoint group names');
      }
    }
  }
  const form = list.form;
  if (form !== undefined) {
    if (!form || typeof form !== 'object') err([...p, 'form'], '"form" must be an object');
    else {
      const fp = [...p, 'form'];
      if (form.fields !== undefined && !Array.isArray(form.fields)) err([...fp, 'fields'], '"fields" must be an array of field names');
      for (const column of ['main', 'side'] as const) {
        const entries = form[column];
        if (entries === undefined) continue;
        if (!Array.isArray(entries)) {
          err([...fp, column], `"${column}" must be an array of field names or sections like { "title": "Dates", "fields": ["Due"] }`);
          continue;
        }
        entries.forEach((entry, i) => {
          if (typeof entry === 'string') return;
          const ep = [...fp, column, i];
          if (!entry || typeof entry !== 'object') err(ep, 'Each entry must be a field name or a section object');
          else if (!Array.isArray(entry.fields) || entry.fields.some(f => typeof f !== 'string')) err(ep, 'A section needs "fields": an array of field names');
          else if (entry.columns !== undefined && entry.columns !== 1 && entry.columns !== 2) err([...ep, 'columns'], '"columns" must be 1 or 2');
        });
      }
      if (form.sideWidth !== undefined && (typeof form.sideWidth !== 'number' || form.sideWidth < 260 || form.sideWidth > 640)) {
        err([...fp, 'sideWidth'], '"sideWidth" must be a number of pixels between 260 and 640');
      }
    }
  }
}
