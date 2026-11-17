import { Row, SpListSourceConfig } from '../types';
import { DataContext, expandFor, listApiUrl, loadSpList, normalizeSpItem } from '../datasources';

export type SpFieldKind =
  | 'text' | 'note' | 'number' | 'currency' | 'boolean' | 'choice' | 'multichoice'
  | 'datetime' | 'lookup' | 'multilookup' | 'user' | 'multiuser' | 'url' | 'other';

export interface SpField {
  name: string;
  title: string;
  kind: SpFieldKind;
  readOnly: boolean;
  required: boolean;
  choices: string[];
  dateOnly: boolean;
  richText: boolean;
  lookupList?: string;
  lookupField?: string;
  description?: string;
}

export interface SpListInfo {
  id: string;
  title: string;
  itemCount: number;
  displayFormUrl?: string;
  isLibrary: boolean;
  entityType?: string;
}

export interface BasePermissions {
  High: number;
  Low: number;
}

export const PermissionKind = {
  EditListItems: 3,
  ManageLists: 12,
  ManageWeb: 31,
} as const;

export function hasPermission(perms: BasePermissions, kind: number): boolean {
  const bit = kind - 1;
  const word = bit < 32 ? perms.Low : perms.High;
  return Math.floor(word / Math.pow(2, bit % 32)) % 2 === 1;
}

function toPermissions(raw: Row): BasePermissions {
  const source = (raw.EffectiveBasePermissions as Row | undefined) || raw;
  return { High: Number(source.High) || 0, Low: Number(source.Low) || 0 };
}

export interface SpComment {
  id: string;
  text: string;
  createdDate: string;
  authorName: string;
  authorEmail?: string;
}

export interface LookupOption {
  id: number;
  text: string;
}

export interface SiteUser {
  id: number;
  name: string;
  email?: string;
}

const NOMETADATA = 'application/json;odata=nometadata';
const VERBOSE = 'application/json;odata=verbose';

const VERBOSE_HEADERS: Record<string, string> = {
  Accept: VERBOSE,
  'Content-Type': VERBOSE,
  'odata-version': '',
};

const KIND_BY_TYPE: Record<string, SpFieldKind> = {
  Text: 'text',
  Note: 'note',
  Number: 'number',
  Currency: 'currency',
  Integer: 'number',
  Counter: 'number',
  Boolean: 'boolean',
  Choice: 'choice',
  MultiChoice: 'multichoice',
  DateTime: 'datetime',
  Lookup: 'lookup',
  LookupMulti: 'multilookup',
  User: 'user',
  UserMulti: 'multiuser',
  URL: 'url',
};

export function toSpField(raw: Row): SpField {
  const type = String(raw.TypeAsString || raw.TypeDisplayName || '');
  const kind = KIND_BY_TYPE[type] || 'other';
  const name = String(raw.InternalName || raw.StaticName || '');
  const choices = Array.isArray(raw.Choices) ? raw.Choices.map(String) : [];
  return {
    name,
    title: String(raw.Title || name),
    kind,
    readOnly: !!raw.ReadOnlyField || type === 'Counter',
    required: !!raw.Required,
    choices,
    dateOnly: raw.DisplayFormat === 0 || raw.DisplayFormat === 'DateOnly',
    richText: !!raw.RichText,
    lookupList: raw.LookupList ? String(raw.LookupList).replace(/[{}]/g, '') : undefined,
    lookupField: raw.LookupField ? String(raw.LookupField) : undefined,
    description: raw.Description ? String(raw.Description) : undefined,
  };
}

export function rootField(path: string): string {
  return path.split('.')[0];
}

function valueOf(res: unknown): Row {
  if (res && typeof res === 'object') {
    const o = res as Row;
    if (o.d && typeof o.d === 'object') return o.d as Row;
    return o;
  }
  return {};
}

function listOf(res: unknown): Row[] {
  const v = valueOf(res);
  const items = v.value || v.results;
  return Array.isArray(items) ? (items as Row[]) : [];
}

export class SpListService {
  private readonly base: string;
  private fieldsPromise: Promise<SpField[]> | undefined;
  private usersPromise: Promise<SiteUser[]> | undefined;
  private infoPromise: Promise<SpListInfo> | undefined;

  public constructor(private readonly ctx: DataContext, public readonly source: SpListSourceConfig) {
    this.base = listApiUrl(source, ctx.webUrl);
  }

  public info(): Promise<SpListInfo> {
    if (!this.infoPromise) {
      this.infoPromise = this.ctx.http
        .getJson(`${this.base}?$select=Id,Title,ItemCount,DefaultDisplayFormUrl,BaseTemplate,ListItemEntityTypeFullName`, {
          sharepoint: true,
          headers: { Accept: NOMETADATA },
        })
        .then(res => {
          const v = valueOf(res);
          return {
            id: String(v.Id || ''),
            title: String(v.Title || this.source.list || ''),
            itemCount: Number(v.ItemCount) || 0,
            displayFormUrl: v.DefaultDisplayFormUrl ? String(v.DefaultDisplayFormUrl) : undefined,
            isLibrary: Number(v.BaseTemplate) === 101,
            entityType: v.ListItemEntityTypeFullName ? String(v.ListItemEntityTypeFullName) : undefined,
          };
        });
      this.infoPromise.catch(() => {
        this.infoPromise = undefined;
      });
    }
    return this.infoPromise;
  }

  public webApiUrl(): string {
    return this.base.slice(0, this.base.indexOf('/_api/') + '/_api/web'.length);
  }

  public async listPermissions(): Promise<BasePermissions> {
    const res = await this.ctx.http.getJson(`${this.base}/EffectiveBasePermissions`, { sharepoint: true, headers: { Accept: NOMETADATA } });
    return toPermissions(valueOf(res));
  }

  public async webPermissions(): Promise<BasePermissions> {
    const res = await this.ctx.http.getJson(`${this.webApiUrl()}/EffectiveBasePermissions`, { sharepoint: true, headers: { Accept: NOMETADATA } });
    return toPermissions(valueOf(res));
  }

  public async currentUserGroups(): Promise<string[]> {
    const res = await this.ctx.http.getJson(`${this.webApiUrl()}/currentuser/groups?$select=Title`, { sharepoint: true, headers: { Accept: NOMETADATA } });
    return listOf(res).map(g => String(g.Title));
  }

  public async fileContent(serverRelativeUrl: string): Promise<ArrayBuffer> {
    const path = encodeURIComponent(serverRelativeUrl.replace(/'/g, "''"));
    return this.ctx.http.getBinary(`${this.webApiUrl()}/GetFileByServerRelativePath(decodedurl='${path}')/$value`, { sharepoint: true });
  }

  public fields(): Promise<SpField[]> {
    if (!this.fieldsPromise) {
      this.fieldsPromise = this.ctx.http
        .getJson(`${this.base}/fields?$filter=Hidden eq false`, { sharepoint: true, headers: { Accept: NOMETADATA } })
        .then(res => listOf(res).map(toSpField));
      this.fieldsPromise.catch(() => {
        this.fieldsPromise = undefined;
      });
    }
    return this.fieldsPromise;
  }

  public async selectFor(paths: string[]): Promise<string[]> {
    const fields = await this.fields();
    const byName = new Map(fields.map((f): [string, SpField] => [f.name, f]));
    const select = new Set<string>(['Id', 'Title', 'Created', 'Modified', 'Author/Title', 'Author/EMail', 'Editor/Title']);
    for (const path of paths) {
      const parts = path.split('.');
      const field = byName.get(parts[0]);
      if (parts.length > 1) {
        select.add(parts.join('/'));
        if (field && (field.kind === 'lookup' || field.kind === 'multilookup' || field.kind === 'user' || field.kind === 'multiuser')) select.add(`${parts[0]}/Id`);
      } else if (field && (field.kind === 'user' || field.kind === 'multiuser')) {
        select.add(`${parts[0]}/Id`).add(`${parts[0]}/Title`).add(`${parts[0]}/EMail`);
      } else if (field && (field.kind === 'lookup' || field.kind === 'multilookup')) {
        select.add(`${parts[0]}/Id`).add(`${parts[0]}/${field.lookupField || 'Title'}`);
      } else {
        select.add(parts[0]);
      }
    }
    return Array.from(select);
  }

  public async items(paths: string[]): Promise<Row[]> {
    const select = await this.selectFor(paths);
    return loadSpList(this.source.list || this.source.listId || 'list', { ...this.source, select }, this.ctx);
  }

  public async item(id: number, paths: string[]): Promise<Row> {
    const select = await this.selectFor(paths);
    const expand = expandFor(select, this.source.expand);
    const query = [`$select=${encodeURIComponent(select.join(','))}`];
    if (expand.length) query.push(`$expand=${encodeURIComponent(expand.join(','))}`);
    const res = await this.ctx.http.getJson(`${this.base}/items(${id})?${query.join('&')}`, { sharepoint: true, headers: { Accept: NOMETADATA } });
    return normalizeSpItem(valueOf(res), this.source.rename);
  }

  public async writeBody(values: Row): Promise<Row> {
    const [info, fields] = await Promise.all([this.info(), this.fields()]);
    if (!info.entityType) throw new Error('SharePoint did not return the list item type (ListItemEntityTypeFullName)');
    return toVerbosePayload(values, info.entityType, fields);
  }

  public async update(id: number, values: Row): Promise<void> {
    await this.ctx.http.postJson(`${this.base}/items(${id})`, {
      sharepoint: true,
      headers: { ...VERBOSE_HEADERS, 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
      body: await this.writeBody(values),
    });
  }

  public async create(values: Row): Promise<number> {
    const res = await this.ctx.http.postJson(`${this.base}/items`, {
      sharepoint: true,
      headers: VERBOSE_HEADERS,
      body: await this.writeBody(values),
    });
    const id = Number(valueOf(res).Id);
    if (isNaN(id)) throw new Error('SharePoint did not return the new item id');
    return id;
  }

  public async comments(id: number): Promise<SpComment[]> {
    const res = await this.ctx.http.getJson(`${this.base}/items(${id})/Comments()`, { sharepoint: true, headers: { Accept: NOMETADATA } });
    return listOf(res).map(toComment);
  }

  public async addComment(id: number, text: string): Promise<SpComment> {
    const res = await this.ctx.http.postJson(`${this.base}/items(${id})/Comments()`, {
      sharepoint: true,
      headers: { Accept: NOMETADATA, 'Content-Type': NOMETADATA, 'odata-version': '' },
      body: { text },
    });
    return toComment(valueOf(res));
  }

  public siteUsers(): Promise<SiteUser[]> {
    if (!this.usersPromise) {
      this.usersPromise = this.ctx.http
        .getJson(`${this.webApiUrl()}/siteusers?$select=Id,Title,Email,PrincipalType&$filter=PrincipalType eq 1&$top=5000`, {
          sharepoint: true,
          headers: { Accept: NOMETADATA },
        })
        .then(res =>
          listOf(res)
            .map(u => ({ id: Number(u.Id), name: String(u.Title || ''), email: u.Email ? String(u.Email) : undefined }))
            .filter(u => u.name)
            .sort((a, b) => a.name.localeCompare(b.name)),
        );
    }
    return this.usersPromise;
  }

  public async lookupOptions(field: SpField): Promise<LookupOption[]> {
    if (!field.lookupList) return [];
    const display = field.lookupField || 'Title';
    const web = this.base.slice(0, this.base.indexOf('/_api/'));
    const url = `${web}/_api/web/lists(guid'${field.lookupList}')/items?$select=Id,${encodeURIComponent(display)}&$orderby=${encodeURIComponent(display)}&$top=5000`;
    const res = await this.ctx.http.getJson(url, { sharepoint: true, headers: { Accept: NOMETADATA } });
    return listOf(res).map(r => ({ id: Number(r.Id), text: String(r[display] === undefined || r[display] === null ? r.Id : r[display]) }));
  }
}

function toComment(raw: Row): SpComment {
  const author = (raw.author || {}) as Row;
  return {
    id: String(raw.id),
    text: String(raw.text || ''),
    createdDate: String(raw.createdDate || ''),
    authorName: String(author.name || author.Title || ''),
    authorEmail: author.email ? String(author.email) : undefined,
  };
}

const STRING_COLLECTION = 'Collection(Edm.String)';
const INT_COLLECTION = 'Collection(Edm.Int32)';

export function toVerbosePayload(values: Row, entityType: string, fields: SpField[]): Row {
  const byName = new Map(fields.map((f): [string, SpField] => [f.name, f]));
  const body: Row = { __metadata: { type: entityType } };
  for (const key of Object.keys(values)) {
    const value = values[key];
    if (!Array.isArray(value)) {
      body[key] = value;
      continue;
    }
    const field = byName.get(key) || (/Id$/.test(key) ? byName.get(key.slice(0, -2)) : undefined);
    const numeric = field ? field.kind === 'multilookup' || field.kind === 'multiuser' : value.every(v => typeof v === 'number');
    body[key] = { __metadata: { type: numeric ? INT_COLLECTION : STRING_COLLECTION }, results: value.map(v => (numeric ? Number(v) : String(v))) };
  }
  return body;
}

export function editValue(field: SpField, value: unknown): Row {
  switch (field.kind) {
    case 'lookup':
    case 'user':
      return { [`${field.name}Id`]: value === undefined || value === null || value === '' ? null : Number(value) };
    case 'multiuser':
    case 'multilookup':
      return { [`${field.name}Id`]: Array.isArray(value) ? value.map(Number) : [] };
    case 'number':
    case 'currency':
      return { [field.name]: value === '' || value === undefined || value === null ? null : Number(value) };
    case 'boolean':
      return { [field.name]: !!value };
    case 'multichoice':
      return { [field.name]: Array.isArray(value) ? value : [] };
    case 'choice':
      return { [field.name]: value === '' || value === undefined ? null : value };
    case 'datetime':
      return { [field.name]: value ? new Date(String(value)).toISOString() : null };
    default:
      return { [field.name]: value === undefined ? null : value };
  }
}

export function isEditable(field: SpField): boolean {
  return !field.readOnly && !field.richText && ['text', 'note', 'number', 'currency', 'boolean', 'choice', 'multichoice', 'datetime', 'lookup', 'multilookup', 'user', 'multiuser'].indexOf(field.kind) >= 0;
}

export function formValue(field: SpField, row: Row | undefined): unknown {
  const value = row ? row[field.name] : undefined;
  switch (field.kind) {
    case 'user':
    case 'lookup': {
      if (value && typeof value === 'object') return Number((value as Row).Id);
      const id = row ? row[`${field.name}Id`] : undefined;
      return id === null || id === undefined ? undefined : Number(id);
    }
    case 'multiuser':
    case 'multilookup': {
      const list = Array.isArray(value) ? value : [];
      return list.map(v => (v && typeof v === 'object' ? Number((v as Row).Id) : Number(v))).filter(n => !isNaN(n));
    }
    case 'multichoice':
      return Array.isArray(value) ? value.map(String) : [];
    case 'boolean':
      return !!value;
    case 'number':
    case 'currency':
      return value === null || value === undefined ? '' : String(value);
    case 'datetime':
      return value ? String(value) : undefined;
    default:
      return value === null || value === undefined ? '' : String(value);
  }
}

export function isEmptyValue(value: unknown): boolean {
  return value === undefined || value === null || value === '' || (Array.isArray(value) && value.length === 0);
}
