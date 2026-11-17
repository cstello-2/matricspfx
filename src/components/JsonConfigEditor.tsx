import * as React from 'react';
import { EditorView, keymap, lineNumbers, highlightActiveLine, drawSelection } from '@codemirror/view';
import { EditorState, Compartment, Extension } from '@codemirror/state';
import { defaultKeymap, history, historyKeymap, indentWithTab } from '@codemirror/commands';
import { json } from '@codemirror/lang-json';
import { linter, lintGutter, Diagnostic as CmDiagnostic } from '@codemirror/lint';
import { autocompletion, Completion, CompletionContext, CompletionResult, closeBrackets } from '@codemirror/autocomplete';
import { bracketMatching, foldGutter, indentOnInput, syntaxHighlighting, defaultHighlightStyle, HighlightStyle } from '@codemirror/language';
import { tags } from '@lezer/highlight';
import { parseConfig } from '../core/validate';
import { FUNCTION_NAMES, AGGREGATE_NAMES } from '../core/expr/evaluate';

export interface JsonConfigEditorProps {
  value: string;
  onChange: (text: string) => void;
  dark?: boolean;
  height?: number | string;
}

const KEYS = [
  'title', 'description', 'refreshSeconds', 'locale', 'dataSources', 'datasets', 'widgets', 'list', 'dataView', 'layout', 'theme',
  'sources', 'export', 'keyPrefix', 'quickFilters', 'statusColors', 'form', 'fields', 'main', 'side', 'sideWidth', 'readOnly', 'comments',
  'type', 'list', 'listId', 'webUrl', 'select', 'expand', 'filter', 'orderBy', 'top', 'pageSize', 'rename', 'rows', 'url', 'path', 'headers', 'sharepoint',
  'from', 'join', 'source', 'on', 'as', 'many', 'compute', 'groupBy', 'field', 'bucket', 'aggregate', 'having', 'sort', 'limit',
  'dataset', 'span', 'height', 'subtitle', 'value', 'label', 'format', 'compare', 'mode', 'goodWhen', 'target', 'thresholds', 'when', 'level', 'sparkline',
  'chart', 'x', 'y', 'series', 'stacked', 'to', 'min', 'max', 'showLegend', 'labels', 'smooth', 'colors', 'echarts', 'name', 'color', 'stack',
  'columns', 'search', 'detail', 'width', 'align', 'bar', 'link', 'status', 'text', 'gap', 'palette', 'cardStyle', 'decimals', 'currency', 'prefix', 'suffix',
];
const VALUES = [
  'spList', 'inline', 'rest', 'kpi', 'chart', 'table', 'text',
  'bar', 'column', 'line', 'area', 'pie', 'donut', 'scatter', 'gauge', 'funnel', 'heatmap', 'sankey', 'treemap', 'radar',
  'number', 'integer', 'decimal', 'percent', 'currency', 'compact', 'date', 'datetime',
  'day', 'week', 'month', 'quarter', 'year', 'left', 'inner', 'good', 'warning', 'serious', 'critical', 'neutral', 'info', 'outlined', 'elevated', 'flat',
];
const KEY_OPTIONS: Completion[] = KEYS.map(k => ({ label: k, type: 'property' }));
const VALUE_OPTIONS: Completion[] = VALUES.map((v): Completion => ({ label: v, type: 'enum' })).concat(
  FUNCTION_NAMES.concat(AGGREGATE_NAMES).map((f): Completion => ({ label: `${f}()`, type: 'function', apply: `${f}(` })),
);

function complete(ctx: CompletionContext): CompletionResult | null {
  const word = ctx.matchBefore(/[\w$]*/);
  if (!word || (word.from === word.to && !ctx.explicit)) return null;
  const line = ctx.state.doc.lineAt(ctx.pos);
  const before = line.text.slice(0, ctx.pos - line.from);
  const inString = (before.match(/"/g) || []).length % 2 === 1;
  if (!inString) return null;
  const isKey = /(^|[{,])\s*"[\w$]*$/.test(before);
  return { from: word.from, options: isKey ? KEY_OPTIONS : VALUE_OPTIONS, validFor: /^[\w$]*$/ };
}

const darkHighlight = HighlightStyle.define([
  { tag: tags.propertyName, color: '#9cdcfe' },
  { tag: tags.string, color: '#ce9178' },
  { tag: tags.number, color: '#b5cea8' },
  { tag: [tags.bool, tags.null], color: '#569cd6' },
  { tag: tags.comment, color: '#6a9955' },
]);

function themeExt(dark: boolean): Extension {
  return [
    EditorView.theme({
      '&': { height: '100%', fontSize: '13px', backgroundColor: dark ? '#1e1e1e' : '#ffffff', color: dark ? '#d4d4d4' : '#1f1f1f' },
      '.cm-scroller': { fontFamily: 'Consolas, "Cascadia Code", Menlo, monospace', lineHeight: '1.5' },
      '.cm-gutters': { backgroundColor: dark ? '#1e1e1e' : '#fafafa', color: dark ? '#858585' : '#999', border: 'none' },
      '.cm-activeLine': { backgroundColor: dark ? 'rgba(255,255,255,0.04)' : 'rgba(0,0,0,0.03)' },
      '.cm-tooltip': { fontFamily: 'Segoe UI, system-ui, sans-serif' },
      '&.cm-focused': { outline: 'none' },
    }, { dark }),
    syntaxHighlighting(dark ? darkHighlight : defaultHighlightStyle),
  ];
}

export function JsonConfigEditor(props: JsonConfigEditorProps): React.ReactElement {
  const host = React.useRef<HTMLDivElement>(null);
  const view = React.useRef<EditorView | undefined>(undefined);
  const themeComp = React.useRef(new Compartment());
  const onChange = React.useRef(props.onChange);
  onChange.current = props.onChange;

  React.useEffect(() => {
    const parent = host.current;
    if (!parent) return;
    const configLinter = linter(v => {
      const text = v.state.doc.toString();
      if (!text.trim()) return [];
      return parseConfig(text).diagnostics.map<CmDiagnostic>(d => ({
        from: Math.min(d.from || 0, text.length),
        to: Math.min(d.to || (d.from || 0) + 1, text.length),
        severity: d.severity,
        message: d.message,
      }));
    }, { delay: 300 });

    const editor = new EditorView({
      parent,
      state: EditorState.create({
        doc: props.value,
        extensions: [
          lineNumbers(),
          foldGutter(),
          history(),
          drawSelection(),
          indentOnInput(),
          bracketMatching(),
          closeBrackets(),
          highlightActiveLine(),
          json(),
          lintGutter(),
          configLinter,
          autocompletion({ override: [complete] }),
          keymap.of([...defaultKeymap, ...historyKeymap, indentWithTab]),
          themeComp.current.of(themeExt(!!props.dark)),
          EditorView.updateListener.of(u => { if (u.docChanged) onChange.current(u.state.doc.toString()); }),
        ],
      }),
    });
    view.current = editor;
    return () => {
      editor.destroy();
      view.current = undefined;
    };
  }, []);

  React.useEffect(() => {
    const v = view.current;
    if (v && v.state.doc.toString() !== props.value) {
      v.dispatch({ changes: { from: 0, to: v.state.doc.length, insert: props.value } });
    }
  }, [props.value]);

  React.useEffect(() => {
    if (view.current) view.current.dispatch({ effects: themeComp.current.reconfigure(themeExt(!!props.dark)) });
  }, [props.dark]);

  return <div ref={host} style={{ height: props.height || '100%', overflow: 'hidden' }} />;
}
