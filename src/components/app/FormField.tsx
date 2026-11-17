import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { IconButton } from '@fluentui/react/lib/Button';
import { Label } from '@fluentui/react/lib/Label';
import { useTheme } from '@fluentui/react/lib/Theme';
import { toDate } from '../../core/dates';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { LookupOption, SiteUser, SpField, isEditable } from '../../core/sharepoint/listService';

export interface FormFieldProps {
  field: SpField;
  value: unknown;
  error?: string;
  disabled?: boolean;
  loadLookup: (field: SpField) => Promise<LookupOption[]>;
  loadUsers: () => Promise<SiteUser[]>;
  initialPeople?: Array<{ id: number; name: string; email?: string }>;
  onChange: (value: unknown) => void;
  renderReadOnly: () => React.ReactNode;
  readOnlyClassName: string;
  readOnlyLabelClassName: string;
}

const NONE = '__none__';

function noneOption(required: boolean): IDropdownOption[] {
  return required ? [] : [{ key: NONE, text: 'None' }];
}

function useLookupOptions(field: SpField, load: (field: SpField) => Promise<LookupOption[]>, enabled: boolean): IDropdownOption[] | undefined {
  const [options, setOptions] = React.useState<IDropdownOption[] | undefined>(undefined);
  const loader = React.useRef(load);
  loader.current = load;
  React.useEffect(() => {
    if (!enabled) return;
    let live = true;
    loader.current(field).then(
      items => {
        if (live) setOptions(items.map(o => ({ key: o.id, text: o.text })));
      },
      () => {
        if (live) setOptions([]);
      },
    );
    return () => {
      live = false;
    };
  }, [field, enabled]);
  return options;
}

function toPersona(u: { id: number; name: string; email?: string }): IPersonaProps {
  return { key: String(u.id), text: u.name, secondaryText: u.email };
}

interface PeopleFieldProps {
  field: SpField;
  value: unknown;
  error?: string;
  disabled?: boolean;
  loadUsers: () => Promise<SiteUser[]>;
  initialPeople: Array<{ id: number; name: string; email?: string }>;
  onChange: (value: unknown) => void;
}

function PeopleField({ field, value, error, disabled, loadUsers, initialPeople, onChange }: PeopleFieldProps): React.ReactElement {
  const theme = useTheme();
  const multi = field.kind === 'multiuser';
  const ids = (Array.isArray(value) ? value : value === undefined || value === null || value === '' ? [] : [value]).map(Number);
  const [known, setKnown] = React.useState<Array<{ id: number; name: string; email?: string }>>(initialPeople);
  const selected = ids.map(id => known.find(k => k.id === id) || { id, name: `User ${id}` }).map(toPersona);
  const search = (filter: string): Promise<IPersonaProps[]> =>
    loadUsers().then(users => {
      const q = filter.trim().toLowerCase();
      return users
        .filter(u => ids.indexOf(u.id) < 0 && (u.name.toLowerCase().indexOf(q) >= 0 || (u.email || '').toLowerCase().indexOf(q) >= 0))
        .slice(0, 10)
        .map(toPersona);
    });
  return (
    <div>
      <Label required={field.required} disabled={disabled}>{field.title}</Label>
      <NormalPeoplePicker
        onResolveSuggestions={search}
        onEmptyResolveSuggestions={() => search('')}
        selectedItems={selected}
        itemLimit={multi ? undefined : 1}
        disabled={disabled}
        inputProps={{ 'aria-label': field.title, placeholder: selected.length && !multi ? '' : 'Search people' }}
        pickerSuggestionsProps={{ suggestionsHeaderText: 'People', noResultsFoundText: 'No people found', loadingText: 'Searching…' }}
        onChange={items => {
          const people = (items || []).map(i => ({ id: Number(i.key), name: String(i.text || ''), email: i.secondaryText }));
          setKnown(k => k.concat(people.filter(p => !k.some(x => x.id === p.id))));
          const next = people.map(p => p.id);
          onChange(multi ? next : next.length ? next[0] : undefined);
        }}
      />
      {error && <div role="alert" style={{ color: theme.semanticColors.errorText, fontSize: 12, marginTop: 4 }}>{error}</div>}
    </div>
  );
}

function toggleKey<T>(list: T[], key: T, selected: boolean): T[] {
  return selected ? (list.indexOf(key) >= 0 ? list : list.concat(key)) : list.filter(k => k !== key);
}

export function FormField(props: FormFieldProps): React.ReactElement {
  const { field, value, error, disabled, onChange } = props;
  const theme = useTheme();
  const editable = isEditable(field);
  const isLookup = field.kind === 'lookup' || field.kind === 'multilookup';
  const lookupOptions = useLookupOptions(field, props.loadLookup, editable && isLookup);
  const label = field.title;

  if (!editable) {
    return (
      <div className={props.readOnlyClassName}>
        <span className={props.readOnlyLabelClassName}>{label}</span>
        <div>{props.renderReadOnly()}</div>
      </div>
    );
  }

  switch (field.kind) {
    case 'user':
    case 'multiuser':
      return (
        <PeopleField
          field={field}
          value={value}
          error={error}
          disabled={disabled}
          loadUsers={props.loadUsers}
          initialPeople={props.initialPeople || []}
          onChange={onChange}
        />
      );
    case 'note':
      return (
        <TextField
          label={label}
          multiline
          autoAdjustHeight
          rows={4}
          value={String(value ?? '')}
          onChange={(_, v) => onChange(v ?? '')}
          required={field.required}
          errorMessage={error}
          disabled={disabled}
          description={field.description}
        />
      );
    case 'number':
    case 'currency':
      return (
        <TextField
          label={label}
          type="number"
          inputMode="decimal"
          value={String(value ?? '')}
          onChange={(_, v) => onChange(v ?? '')}
          required={field.required}
          errorMessage={error}
          disabled={disabled}
          description={field.description}
        />
      );
    case 'boolean':
      return <Toggle label={label} checked={!!value} onChange={(_, checked) => onChange(!!checked)} disabled={disabled} onText="Yes" offText="No" />;
    case 'choice': {
      const options: IDropdownOption[] = noneOption(field.required).concat(field.choices.map(c => ({ key: c, text: c })));
      return (
        <Dropdown
          label={label}
          options={options}
          selectedKey={value === undefined || value === '' || value === null ? (field.required ? null : NONE) : String(value)}
          placeholder="Select an option"
          onChange={(_, o) => o && onChange(o.key === NONE ? '' : String(o.key))}
          required={field.required}
          errorMessage={error}
          disabled={disabled}
        />
      );
    }
    case 'multichoice': {
      const selected = Array.isArray(value) ? value.map(String) : [];
      return (
        <Dropdown
          label={label}
          multiSelect
          options={field.choices.map(c => ({ key: c, text: c }))}
          selectedKeys={selected}
          placeholder="Select options"
          onChange={(_, o) => o && onChange(toggleKey(selected, String(o.key), !!o.selected))}
          required={field.required}
          errorMessage={error}
          disabled={disabled}
        />
      );
    }
    case 'lookup': {
      const options: IDropdownOption[] = lookupOptions ? noneOption(field.required).concat(lookupOptions) : [];
      return (
        <Dropdown
          label={label}
          options={options}
          selectedKey={value === undefined || value === null ? (field.required ? null : NONE) : Number(value)}
          placeholder={lookupOptions ? 'Select an item' : 'Loading…'}
          onChange={(_, o) => o && onChange(o.key === NONE ? undefined : Number(o.key))}
          required={field.required}
          errorMessage={error}
          disabled={disabled || !lookupOptions}
        />
      );
    }
    case 'multilookup': {
      const selected = Array.isArray(value) ? value.map(Number) : [];
      return (
        <Dropdown
          label={label}
          multiSelect
          options={lookupOptions || []}
          selectedKeys={selected}
          placeholder={lookupOptions ? 'Select items' : 'Loading…'}
          onChange={(_, o) => o && onChange(toggleKey(selected, Number(o.key), !!o.selected))}
          required={field.required}
          errorMessage={error}
          disabled={disabled || !lookupOptions}
        />
      );
    }
    case 'datetime': {
      const date = toDate(value);
      return (
        <div>
          <Label required={field.required} disabled={disabled}>{label}</Label>
          <div style={{ display: 'flex', gap: 4, alignItems: 'flex-start' }}>
            <DatePicker
              value={date || undefined}
              onSelectDate={d => onChange(d ? d.toISOString() : undefined)}
              placeholder="Select a date"
              ariaLabel={label}
              disabled={disabled}
              allowTextInput
              styles={{ root: { flex: '1 1 auto' } }}
            />
            {date && !field.required && (
              <IconButton iconProps={{ iconName: 'Clear' }} title={`Clear ${label}`} ariaLabel={`Clear ${label}`} onClick={() => onChange(undefined)} disabled={disabled} />
            )}
          </div>
          {error && <div role="alert" style={{ color: theme.semanticColors.errorText, fontSize: 12, marginTop: 4 }}>{error}</div>}
        </div>
      );
    }
    default:
      return (
        <TextField
          label={label}
          value={String(value ?? '')}
          onChange={(_, v) => onChange(v ?? '')}
          required={field.required}
          errorMessage={error}
          disabled={disabled}
          description={field.description}
        />
      );
  }
}
