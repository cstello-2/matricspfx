import * as React from 'react';
import { Modal } from '@fluentui/react/lib/Modal';
import { ActionButton, DefaultButton, IconButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Link } from '@fluentui/react/lib/Link';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { TextField } from '@fluentui/react/lib/TextField';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { ThemeProvider } from '@fluentui/react/lib/Theme';
import { getTheme, mergeStyles } from '@fluentui/react/lib/Styling';
import { Row } from '../../core/types';
import { formatValue } from '../../core/format';
import { SpField, editValue, formValue, isEditable, isEmptyValue } from '../../core/sharepoint/listService';
import { AutoTagger, TagSuggestion, currentTags } from '../../core/ai/autotag';
import { MetricsTheme, themeFromSharePoint } from '../theme';
import { AppStyles, getAppStyles } from './appStyles';
import { FormSectionLayout, ListModel, formLayout, itemKey, itemPaths } from './useListModel';
import { FieldValue, Lozenge, people, plainText, statusLevel } from './fieldDisplay';
import { FormField } from './FormField';
import { Comments } from './Comments';
import { AutotagState } from './useAutotag';
import { SuggestionChips } from './TagChips';

export interface ItemViewProps {
  model: ListModel;
  id: number | undefined;
  order: number[];
  webUrl: string;
  currentUser?: string;
  locale?: string;
  autotag: AutotagState;
  onClose: () => void;
  onNavigate: (id: number) => void;
  onCreated: (id: number) => void;
  onChanged: () => void;
}

type Draft = Record<string, unknown>;

function personRefs(value: unknown): Array<{ id: number; name: string; email?: string }> {
  const list = Array.isArray(value) ? value : value ? [value] : [];
  return list
    .filter((v): v is Row => !!v && typeof v === 'object' && (v as Row).Id !== undefined)
    .map(v => ({ id: Number(v.Id), name: String(v.Title || ''), email: v.EMail ? String(v.EMail) : undefined }));
}

function absoluteUrl(serverRelative: string, webUrl: string): string {
  const m = /^(https?:\/\/[^/]+)/i.exec(webUrl);
  return (m ? m[1] : '') + serverRelative;
}

function same(a: unknown, b: unknown): boolean {
  return JSON.stringify(a === '' ? undefined : a) === JSON.stringify(b === '' ? undefined : b);
}

function readOnlyBy(model: ListModel, field: SpField): boolean {
  const ro = model.config.form && model.config.form.readOnly;
  if (!isEditable(field) || ro === true) return true;
  return Array.isArray(ro) && ro.indexOf(field.name) >= 0;
}

interface ResolvedSection extends FormSectionLayout {
  resolved: SpField[];
}

interface SuggesterProps {
  tagger: AutoTagger;
  row: Row;
  theme: MetricsTheme;
  styles: AppStyles;
  onAdd: (tag: string) => void;
}

function TagSuggester({ tagger, row, theme, styles, onAdd }: SuggesterProps): React.ReactElement {
  const [loading, setLoading] = React.useState(false);
  const [suggestions, setSuggestions] = React.useState<TagSuggestion[] | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);

  const suggest = (): void => {
    setLoading(true);
    setError(undefined);
    tagger.suggest([row]).then(
      ([result]) => {
        setSuggestions(result.suggested.map(s => ({ ...s, chosen: false })));
        setError(result.error);
        setLoading(false);
      },
      (e: Error) => {
        setError(e.message);
        setLoading(false);
      },
    );
  };

  return (
    <section className={styles.sideCard} aria-label="Tag suggestions">
      <h3 className={styles.sideCardTitle}>Suggested {tagger.field.title.toLowerCase()}</h3>
      {suggestions && (
        <SuggestionChips
          suggestions={suggestions}
          theme={theme}
          onToggle={tag => {
            onAdd(tag);
            setSuggestions(list => (list ? list.filter(s => s.tag !== tag) : list));
          }}
        />
      )}
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      <div>
        {loading ? (
          <Spinner size={SpinnerSize.small} label="Finding tags…" labelPosition="right" />
        ) : (
          <ActionButton iconProps={{ iconName: 'Lightbulb' }} text={suggestions ? 'Suggest again' : 'Suggest tags with AI'} onClick={suggest} styles={{ root: { marginLeft: -8 } }} />
        )}
      </div>
    </section>
  );
}

export function ItemView(props: ItemViewProps): React.ReactElement {
  const { model, id, order, locale } = props;
  const isNew = id === undefined;
  const siteTheme = getTheme();
  const theme = React.useMemo(() => themeFromSharePoint(siteTheme), [siteTheme]);
  const styles = getAppStyles(theme);
  const form = model.config.form || {};
  const service = model.service;
  const paths = React.useMemo(() => itemPaths(model), [model]);
  const pathKey = paths.join(',');

  const titleField = model.fieldMap.get(form.title || 'Title');
  const statusField = form.status ? model.fieldMap.get(form.status) : undefined;
  const sections = React.useMemo(() => {
    const layout = formLayout(model.config, model.columns, model.fields);
    const resolve = (list: FormSectionLayout[]): ResolvedSection[] =>
      list
        .map(sec => ({ ...sec, resolved: sec.fields.map(f => model.fieldMap.get(f)).filter((f): f is SpField => !!f) }))
        .filter(sec => sec.resolved.length > 0);
    return { main: resolve(layout.main), side: resolve(layout.side) };
  }, [model.config, model.columns, model.fields, model.fieldMap]);
  const formFieldList = React.useMemo(() => {
    const all = [titleField].concat(sections.main.concat(sections.side).reduce<SpField[]>((acc, sec) => acc.concat(sec.resolved), []));
    return all.filter((f, i): f is SpField => !!f && all.indexOf(f) === i);
  }, [titleField, sections]);

  const [row, setRow] = React.useState<Row | undefined>(isNew ? undefined : (model.rows || []).find(r => Number(r.Id) === id));
  const [loaded, setLoaded] = React.useState(isNew);
  const [loadError, setLoadError] = React.useState<string | undefined>(undefined);
  const [draft, setDraft] = React.useState<Draft>({});
  const [errors, setErrors] = React.useState<Record<string, string>>({});
  const [saving, setSaving] = React.useState(false);
  const [saveError, setSaveError] = React.useState<string | undefined>(undefined);
  const [savedAt, setSavedAt] = React.useState<number | undefined>(undefined);
  const [pending, setPending] = React.useState<(() => void) | undefined>(undefined);

  const initial = React.useMemo(() => {
    const values: Draft = {};
    for (const f of formFieldList) values[f.name] = formValue(f, row);
    return values;
  }, [formFieldList, row]);

  React.useEffect(() => {
    setDraft(initial);
    setErrors({});
  }, [initial]);

  React.useEffect(() => {
    if (id === undefined) return;
    let live = true;
    setLoadError(undefined);
    setSaveError(undefined);
    service.item(id, pathKey.split(',')).then(
      fresh => {
        if (!live) return;
        setRow(fresh);
        setLoaded(true);
      },
      (e: Error) => {
        if (live) setLoadError(e.message);
      },
    );
    return () => {
      live = false;
    };
  }, [service, id, pathKey]);

  React.useEffect(() => {
    if (savedAt === undefined) return;
    const timer = setTimeout(() => setSavedAt(undefined), 3000);
    return () => clearTimeout(timer);
  }, [savedAt]);

  const editableFields = formFieldList.filter(f => !readOnlyBy(model, f));
  const changed = editableFields.filter(f => !same(draft[f.name], initial[f.name]));
  const dirty = changed.length > 0;

  const setValue = (name: string, value: unknown): void => {
    setDraft(d => ({ ...d, [name]: value }));
    setErrors(e => {
      if (!e[name]) return e;
      const next = { ...e };
      delete next[name];
      return next;
    });
  };

  const validate = (): boolean => {
    const next: Record<string, string> = {};
    for (const f of editableFields) if (f.required && isEmptyValue(draft[f.name])) next[f.name] = `${f.title} is required.`;
    setErrors(next);
    return Object.keys(next).length === 0;
  };

  const save = (): void => {
    if (saving || (!dirty && !isNew) || !validate()) return;
    const fields = isNew ? editableFields.filter(f => !isEmptyValue(draft[f.name])) : changed;
    const payload = fields.reduce<Row>((acc, f) => ({ ...acc, ...editValue(f, draft[f.name]) }), {});
    setSaving(true);
    setSaveError(undefined);
    const write = id === undefined ? service.create(payload) : service.update(id, payload).then(() => id);
    write
      .then(savedId => {
        if (isNew) {
          props.onChanged();
          props.onCreated(savedId);
          return undefined;
        }
        return service.item(savedId, paths).then(fresh => {
          setRow(fresh);
          model.replaceRow(fresh);
          setSavedAt(Date.now());
          props.onChanged();
        });
      })
      .then(
        () => setSaving(false),
        (e: Error) => {
          setSaveError(e.message);
          setSaving(false);
        },
      );
  };

  const hasInput = isNew && editableFields.some(f => !isEmptyValue(draft[f.name]) && f.kind !== 'boolean');
  const guard = (action: () => void): void => {
    if (dirty || hasInput) setPending(() => action);
    else action();
  };

  const position = id === undefined ? -1 : order.indexOf(id);
  const previous = position > 0 ? order[position - 1] : undefined;
  const next = position >= 0 && position < order.length - 1 ? order[position + 1] : undefined;
  const listTitle = model.info ? model.info.title : model.service.source.list || 'List';
  const key = isNew ? 'New item' : row ? itemKey(model, row) : String(id);
  const displayForm = !isNew && model.info && model.info.displayFormUrl ? `${absoluteUrl(model.info.displayFormUrl, props.webUrl)}?ID=${id}` : undefined;
  const author = row ? people(row.Author)[0] : undefined;
  const editor = row ? people(row.Editor)[0] : undefined;

  const candidate = props.autotag.tagger;
  const tagger = candidate && props.autotag.access && props.autotag.access.allowed && model.fieldMap.get(candidate.field.name) === candidate.field ? candidate : undefined;
  const suggestionRow: Row = { ...(row || {}), ...draft };
  const addTag = (tag: string): void => {
    if (!tagger) return;
    const field = tagger.field;
    if (field.kind === 'multichoice') setValue(field.name, currentTags(field, suggestionRow).concat(tag).filter((t, i, all) => all.indexOf(t) === i));
    else if (field.kind === 'choice') setValue(field.name, tag);
    else if (field.kind === 'text' || field.kind === 'note') setValue(field.name, currentTags(field, suggestionRow).concat(tag).join('; '));
    else {
      tagger.vocabulary().then(
        vocab => {
          const match = vocab.find(v => v.name === tag);
          if (!match || match.id === undefined) return;
          const matchId = Number(match.id);
          if (field.kind === 'lookup') setValue(field.name, matchId);
          else setDraft(d => ({ ...d, [field.name]: (Array.isArray(d[field.name]) ? (d[field.name] as number[]) : []).filter(x => x !== matchId).concat(matchId) }));
        },
        (e: Error) => setSaveError(e.message),
      );
    }
  };

  const loadLookup = React.useCallback((f: SpField) => service.lookupOptions(f), [service]);
  const loadUsers = React.useCallback(() => service.siteUsers(), [service]);

  const renderReadOnly = (field: SpField): (() => React.ReactNode) => (): React.ReactNode =>
    row ? <FieldValue path={field.name} field={field} row={row} list={model.config} theme={theme} locale={locale} /> : <span style={{ color: theme.textMuted }}>—</span>;

  const renderStatus = (field: SpField): React.ReactNode => {
    const value = draft[field.name];
    const options: IDropdownOption[] = field.choices.map(c => ({ key: c, text: c }));
    return (
      <Dropdown
        key={field.name}
        label={field.title}
        options={options}
        selectedKey={value === undefined || value === '' ? null : String(value)}
        placeholder="Select a status"
        disabled={readOnlyBy(model, field) || saving}
        required={field.required}
        errorMessage={errors[field.name]}
        onChange={(_, o) => o && setValue(field.name, String(o.key))}
        onRenderTitle={opts => {
          const current = opts && opts[0] ? String(opts[0].key) : '';
          return current ? <Lozenge value={current} level={statusLevel(current, model.config.statusColors)} theme={theme} /> : <span>Select a status</span>;
        }}
        onRenderOption={option => (option ? <Lozenge value={option.text} level={statusLevel(option.text, model.config.statusColors)} theme={theme} /> : null)}
      />
    );
  };

  const renderLongText = (field: SpField): React.ReactNode => {
    if (readOnlyBy(model, field)) {
      return (
        <div key={field.name} className={styles.readOnlyField}>
          <span className={styles.readOnlyLabel}>{field.title}</span>
          <div className={styles.longText}>
            {row && row[field.name] ? plainText(String(row[field.name])) : <span style={{ color: theme.textMuted }}>—</span>}
          </div>
        </div>
      );
    }
    return (
      <TextField
        key={field.name}
        label={field.title}
        multiline
        autoAdjustHeight
        rows={6}
        placeholder={`Add ${field.title.toLowerCase()}…`}
        value={String(draft[field.name] ?? '')}
        onChange={(_, v) => setValue(field.name, v ?? '')}
        required={field.required}
        errorMessage={errors[field.name]}
        disabled={saving}
      />
    );
  };

  const renderField = (field: SpField, place: 'main' | 'side'): React.ReactNode => {
    if (field === statusField && field.kind === 'choice') return renderStatus(field);
    if (place === 'main' && field.kind === 'note') return renderLongText(field);
    return (
      <FormField
        key={field.name}
        field={readOnlyBy(model, field) ? { ...field, readOnly: true } : field}
        value={draft[field.name]}
        error={errors[field.name]}
        disabled={saving}
        loadLookup={loadLookup}
        loadUsers={loadUsers}
        initialPeople={row ? personRefs(row[field.name]) : []}
        onChange={v => setValue(field.name, v)}
        renderReadOnly={renderReadOnly(field)}
        readOnlyClassName={styles.readOnlyField}
        readOnlyLabelClassName={styles.readOnlyLabel}
      />
    );
  };

  const sideWidth = Math.max(260, Math.min(640, form.sideWidth || 380));
  const bodyClass = React.useMemo(
    () =>
      mergeStyles(styles.itemBody, {
        gridTemplateColumns: `minmax(0, 1fr) minmax(260px, ${sideWidth}px)`,
        selectors: { '@media (max-width: 900px)': { gridTemplateColumns: 'minmax(0, 1fr)' } },
      }),
    [styles.itemBody, sideWidth],
  );

  const onKeyDown = (e: React.KeyboardEvent): void => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      save();
    }
  };

  const showForm = isNew || (loaded && !!row);

  return (
    <Modal
      isOpen
      onDismiss={() => guard(props.onClose)}
      isBlocking={false}
      titleAriaId="metrics-item-title"
      styles={{
        main: { position: 'fixed', top: 0, left: 0, width: '100vw', height: '100vh', maxWidth: '100vw', maxHeight: '100vh', borderRadius: 0 },
        scrollableContent: { height: '100%', maxHeight: '100vh', overflow: 'hidden' },
      }}
    >
      <ThemeProvider theme={siteTheme} applyTo="none" style={{ height: '100%' }}>
        <div className={styles.item} onKeyDown={onKeyDown}>
          <div className={styles.itemTop}>
            <DefaultButton
              iconProps={{ iconName: 'Back' }}
              text={`Back to ${listTitle}`}
              onClick={() => guard(props.onClose)}
              styles={{ root: { flex: '0 0 auto' } }}
            />
            <nav className={styles.breadcrumb} aria-label="Breadcrumb">
              <Link onClick={() => guard(props.onClose)}>{listTitle}</Link>
              <span aria-hidden="true">/</span>
              <span className={styles.keyBadge}>{key}</span>
            </nav>
            <div className={styles.itemTopActions}>
              {row && typeof row.FileRef === 'string' && (
                <ActionButton iconProps={{ iconName: 'OpenFile' }} text="Open document" href={absoluteUrl(row.FileRef, props.webUrl)} target="_blank" rel="noopener noreferrer" />
              )}
              {displayForm && <ActionButton iconProps={{ iconName: 'OpenInNewWindow' }} text="Open in SharePoint" href={displayForm} target="_blank" rel="noopener noreferrer" />}
              {!isNew && (
                <>
                  <IconButton iconProps={{ iconName: 'ChevronUp' }} title="Previous item" ariaLabel="Previous item" disabled={previous === undefined} onClick={() => previous !== undefined && guard(() => props.onNavigate(previous))} />
                  <IconButton iconProps={{ iconName: 'ChevronDown' }} title="Next item" ariaLabel="Next item" disabled={next === undefined} onClick={() => next !== undefined && guard(() => props.onNavigate(next))} />
                </>
              )}
              <IconButton iconProps={{ iconName: 'Cancel' }} title="Close" ariaLabel="Close" onClick={() => guard(props.onClose)} />
            </div>
          </div>

          {!showForm && !loadError && <Spinner size={SpinnerSize.large} label="Loading item…" styles={{ root: { marginTop: 64 } }} />}
          {loadError && (
            <div style={{ padding: 24 }}>
              <MessageBar messageBarType={MessageBarType.error}>{loadError}</MessageBar>
            </div>
          )}

          {showForm && (
            <form
              className={bodyClass}
              onSubmit={e => {
                e.preventDefault();
                save();
              }}
              noValidate
            >
              <main className={styles.itemMain}>
                {savedAt !== undefined && (
                  <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setSavedAt(undefined)}>
                    Changes saved.
                  </MessageBar>
                )}
                {titleField && (
                  <TextField
                    id="metrics-item-title"
                    ariaLabel={titleField.title}
                    placeholder={isNew ? `${titleField.title} of the new item` : titleField.title}
                    value={String(draft[titleField.name] ?? '')}
                    onChange={(_, v) => setValue(titleField.name, v ?? '')}
                    readOnly={readOnlyBy(model, titleField)}
                    errorMessage={errors[titleField.name]}
                    autoFocus={isNew}
                    borderless
                    styles={{
                      fieldGroup: { height: 'auto', borderRadius: 6, background: 'transparent', selectors: { ':hover': { background: theme.hover } } },
                      field: { fontSize: 24, fontWeight: 600, lineHeight: '32px', padding: '6px 8px', color: theme.text },
                      root: { marginLeft: -8 },
                    }}
                  />
                )}

                {sections.main.map((sec, i) => (
                  <section key={`main-${i}`} className={styles.section} aria-label={sec.title}>
                    {sec.title && <h2 className={styles.sectionTitle}>{sec.title}</h2>}
                    <div className={sec.columns === 2 ? styles.fieldGridTwo : styles.fieldGrid}>
                      {sec.resolved.map(field => renderField(field, 'main'))}
                    </div>
                  </section>
                ))}

                {id !== undefined && form.comments !== false && (
                  <section className={styles.section}>
                    <h2 className={styles.sectionTitle}>Comments</h2>
                    <Comments service={service} itemId={id} currentUser={props.currentUser} styles={styles} locale={locale} />
                  </section>
                )}
              </main>

              <aside className={styles.itemSide}>
                {sections.side.map((sec, i) => (
                  <section key={`side-${i}`} className={styles.sideCard} aria-label={sec.title || 'Details'}>
                    {sec.title && <h3 className={styles.sideCardTitle}>{sec.title}</h3>}
                    <div className={sec.columns === 2 ? styles.fieldGridTwo : styles.fieldGrid}>
                      {sec.resolved.map(field => renderField(field, 'side'))}
                    </div>
                  </section>
                ))}

                {tagger && <TagSuggester tagger={tagger} row={suggestionRow} theme={theme} styles={styles} onAdd={addTag} />}

                {row && (
                  <div className={styles.meta}>
                    <div>Created {formatValue(row.Created, 'datetime', locale)}{author ? ` by ${author.name}` : ''}</div>
                    <div>Updated {formatValue(row.Modified, 'datetime', locale)}{editor ? ` by ${editor.name}` : ''}</div>
                  </div>
                )}
              </aside>
            </form>
          )}

          {showForm && (dirty || isNew || saveError) && (
            <div className={styles.saveBar} role="region" aria-label="Save changes">
              {saveError && (
                <MessageBar messageBarType={MessageBarType.error} styles={{ root: { flex: '1 1 100%' } }} onDismiss={() => setSaveError(undefined)}>
                  {saveError}
                </MessageBar>
              )}
              <span className={styles.saveBarText}>
                {isNew ? 'New item' : `${changed.length} unsaved change${changed.length === 1 ? '' : 's'}`} · Ctrl+S to save
              </span>
              <DefaultButton text={isNew ? 'Cancel' : 'Discard'} onClick={() => (isNew ? guard(props.onClose) : setDraft(initial))} disabled={saving} />
              <PrimaryButton text={saving ? 'Saving…' : isNew ? 'Create' : 'Save'} onClick={save} disabled={saving || (!dirty && !isNew)} />
            </div>
          )}

          <Dialog
            hidden={!pending}
            onDismiss={() => setPending(undefined)}
            dialogContentProps={{ type: DialogType.normal, title: 'Discard unsaved changes?', subText: 'Your changes to this item have not been saved.' }}
            modalProps={{ isBlocking: true }}
          >
            <DialogFooter>
              <PrimaryButton
                text="Discard changes"
                onClick={() => {
                  const action = pending;
                  setPending(undefined);
                  setDraft(initial);
                  if (action) action();
                }}
              />
              <DefaultButton text="Keep editing" onClick={() => setPending(undefined)} />
            </DialogFooter>
          </Dialog>
        </div>
      </ThemeProvider>
    </Modal>
  );
}
