import type { Plugin, ViteDevServer } from 'vite';
import type { IncomingMessage, ServerResponse } from 'node:http';
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { join, normalize, sep } from 'node:path';
import { evaluate, compare } from '../../src/core/expr/evaluate';
import { odataFilterToExpr } from './odata';

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

interface CommentFixture {
  id: number;
  text: string;
  createdDate: string;
  authorId: number;
}

interface ListFixture {
  title: string;
  id: string;
  baseTemplate?: number;
  fields: Record<string, FieldDef>;
  items: Item[];
  comments?: Record<string, CommentFixture[]>;
}

interface Route {
  webPath: string;
  list: ListFixture;
  rest: string;
}

export interface SpMockOptions {
  fixturesDir: string;
  latency?: [number, number];
  currentUser: { id: number; name: string; email: string };
}

const SYSTEM_FIELDS: Record<string, FieldDef> = {
  ID: { type: 'Counter', title: 'ID', readOnly: true },
  Created: { type: 'DateTime', title: 'Created', readOnly: true },
  Modified: { type: 'DateTime', title: 'Modified', readOnly: true },
  Author: { type: 'User', title: 'Created By', readOnly: true },
  Editor: { type: 'User', title: 'Modified By', readOnly: true },
};

const EXPANDABLE = ['Lookup', 'LookupMulti', 'User', 'UserMulti'];

type Role = 'owner' | 'member';

function permissionMask(kinds: number[]): { High: string; Low: string } {
  let low = 0;
  let high = 0;
  for (const k of kinds) {
    if (k <= 32) low += Math.pow(2, k - 1);
    else high += Math.pow(2, k - 33);
  }
  return { High: String(high), Low: String(low) };
}

const FULL_CONTROL = { High: '2147483647', Low: '4294967295' };
const MEMBER = permissionMask([1, 2, 3, 4, 5, 6, 7, 10, 13, 17, 18, 20, 21, 22]);
const GROUPS: Record<Role, string[]> = { owner: ['Dev Site Owners'], member: ['Dev Site Members'] };

class HttpError extends Error {
  public constructor(public readonly status: number, message: string, public readonly code: string = '-1, Microsoft.SharePoint.Client.InvalidClientQueryException') {
    super(message);
  }
}

function send(res: ServerResponse, status: number, body?: unknown): void {
  res.statusCode = status;
  res.setHeader('Content-Type', 'application/json;odata=nometadata;charset=utf-8');
  const value = body && typeof body === 'object' ? (body as Item).value : undefined;
  if (Array.isArray(value)) res.setHeader('X-Item-Count', String(value.length));
  res.end(body === undefined ? '' : JSON.stringify(body));
}

function sendError(res: ServerResponse, e: unknown): void {
  const status = e instanceof HttpError ? e.status : 500;
  const code = e instanceof HttpError ? e.code : '-1, System.Exception';
  send(res, status, { 'odata.error': { code, message: { lang: 'en-US', value: e instanceof Error ? e.message : String(e) } } });
}

function readBody(req: IncomingMessage): Promise<Item> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    req.on('data', (c: Buffer) => chunks.push(c));
    req.on('end', () => {
      const text = Buffer.concat(chunks).toString('utf8');
      if (!text) {
        resolve({});
        return;
      }
      try {
        resolve(JSON.parse(text));
      } catch (e) {
        reject(new HttpError(400, `Invalid JSON body: ${(e as Error).message}`));
      }
    });
    req.on('error', reject);
  });
}

function isVerbose(header: string | string[] | undefined): boolean {
  return /odata=verbose/i.test(String(header || ''));
}

function entityTypeOf(list: { title: string; baseTemplate?: number }): string {
  const name = list.title.replace(/ /g, '_x0020_');
  return list.baseTemplate === 101 ? `SP.Data.${name}Item` : `SP.Data.${name}ListItem`;
}

function writePayload(list: { title: string; baseTemplate?: number }, req: IncomingMessage, body: Item): Item {
  if (!isVerbose(req.headers['content-type'])) return body;
  const meta = body.__metadata as Item | undefined;
  const expected = entityTypeOf(list);
  if (!meta || meta.type !== expected) {
    throw new HttpError(400, `A type named '${meta ? String(meta.type) : ''}' could not be resolved by the model. When a model is available, each type name must resolve to a valid type. Expected '${expected}'.`, '-1, Microsoft.SharePoint.Client.InvalidClientQueryException');
  }
  const out: Item = {};
  for (const key of Object.keys(body)) {
    if (key === '__metadata') continue;
    const value = body[key];
    if (Array.isArray(value)) {
      throw new HttpError(400, `An unexpected 'StartArray' node was found when reading from the JSON reader. A 'StartObject' node was expected. (property '${key}')`, '-1, Microsoft.OData.ODataException');
    }
    if (value && typeof value === 'object' && Array.isArray((value as Item).results)) {
      out[key] = (value as Item).results;
      continue;
    }
    out[key] = value;
  }
  return out;
}

function splitList(value: string | null): string[] {
  return value ? value.split(',').map(s => s.trim()).filter(Boolean) : [];
}

export function spMock(opts: SpMockOptions): Plugin {
  let lists = new Map<string, ListFixture>();
  let users: Item[] = [];
  let comments = new Map<string, CommentFixture[]>();
  let changed = new Set<string>();
  let notify: () => void = () => undefined;

  const load = (): void => {
    const next = new Map<string, ListFixture>();
    const nextComments = new Map<string, CommentFixture[]>();
    const dir = join(opts.fixturesDir, 'lists');
    if (existsSync(dir)) {
      for (const file of readdirSync(dir)) {
        if (!file.endsWith('.json')) continue;
        try {
          const list = JSON.parse(readFileSync(join(dir, file), 'utf8')) as ListFixture;
          list.fields = { ...SYSTEM_FIELDS, ...(list.fields || {}) };
          next.set(list.title.toLowerCase(), list);
          next.set(`guid:${list.id.toLowerCase()}`, list);
          for (const itemId of Object.keys(list.comments || {})) {
            nextComments.set(`${list.id}:${itemId}`, (list.comments || {})[itemId].slice());
          }
        } catch (e) {
          console.error(`[sp-mock] could not read ${file}: ${(e as Error).message}`);
        }
      }
    }
    const usersFile = join(opts.fixturesDir, 'users.json');
    const loaded: Item[] = existsSync(usersFile) ? JSON.parse(readFileSync(usersFile, 'utf8')) : [];
    const me = opts.currentUser;
    const self: Item = { Id: me.id, Title: me.name, EMail: me.email };
    users = loaded.some(u => u.Id === me.id) ? loaded : [self].concat(loaded);
    lists = next;
    comments = nextComments;
    changed = new Set<string>();
  };
  load();

  const findList = (key: string): ListFixture | undefined => lists.get(key.toLowerCase());
  const findUser = (id: unknown): Item | undefined => users.find(u => u.Id === id);

  const expandItem = (list: ListFixture, item: Item): Item => {
    const full: Item = { ...item, ID: item.Id, EditorId: item.EditorId === undefined ? item.AuthorId : item.EditorId };
    for (const name of Object.keys(list.fields)) {
      const def = list.fields[name];
      if (EXPANDABLE.indexOf(def.type) < 0) continue;
      const idValue = full[`${name}Id`];
      const resolve = (id: unknown): Item | null => {
        if (id === null || id === undefined) return null;
        if (def.type === 'User' || def.type === 'UserMulti') return findUser(id) || null;
        const target = def.list ? findList(def.list) : undefined;
        return (target && target.items.find(t => t.Id === id)) || null;
      };
      full[name] = Array.isArray(idValue) ? idValue.map(resolve).filter(Boolean) : resolve(idValue);
    }
    return full;
  };

  const project = (full: Item, select: string[], list: ListFixture, expand: string[]): Item => {
    if (!select.length || (select.length === 1 && select[0] === '*')) {
      const out: Item = {};
      for (const k of Object.keys(full)) {
        const def = list.fields[k];
        if (def && EXPANDABLE.indexOf(def.type) >= 0 && expand.indexOf(k) < 0) continue;
        out[k] = full[k];
      }
      return out;
    }
    const out: Item = {};
    for (const s of select) {
      const [head, sub] = s.split('/');
      if (!sub) {
        out[head] = full[head] === undefined ? null : full[head];
        continue;
      }
      if (expand.indexOf(head) < 0) {
        throw new HttpError(400, `The query to field '${head}' is not valid. The $select query string must specify the target fields and the $expand query string must contains ${head}.`);
      }
      const value = full[head];
      if (Array.isArray(value)) {
        const arr = (out[head] as Item[]) || value.map(() => ({}));
        value.forEach((x, i) => {
          arr[i][sub] = (x as Item)[sub];
        });
        out[head] = arr;
      } else if (value && typeof value === 'object') {
        const o = (out[head] as Item) || {};
        o[sub] = (value as Item)[sub];
        out[head] = o;
      } else {
        out[head] = null;
      }
    }
    return out;
  };

  const route = (pathname: string): Route | undefined => {
    const m = /^(.*?)\/_api\/web\/(?:lists\/getbytitle\('((?:[^']|'')+)'\)|lists\(guid'([0-9a-fA-F-]+)'\))(.*)$/i.exec(pathname);
    if (!m) return undefined;
    const list = m[2] ? findList(m[2].replace(/''/g, "'")) : findList(`guid:${m[3]}`);
    if (!list) {
      throw new HttpError(404, `List '${m[2] || m[3]}' does not exist at site with URL '${m[1]}'. Add workbench/fixtures/lists/${m[2] || 'Name'}.json to mock it.`, '-1, System.ArgumentException');
    }
    return { webPath: m[1], list, rest: m[4] };
  };

  const listInfo = (r: Route): Item => ({
    Id: r.list.id,
    Title: r.list.title,
    ItemCount: r.list.items.length,
    BaseTemplate: r.list.baseTemplate || 100,
    ListItemEntityTypeFullName: entityTypeOf(r.list),
    DefaultDisplayFormUrl: r.list.baseTemplate === 101 ? `${r.webPath}/${r.list.title}/Forms/DispForm.aspx` : `${r.webPath}/Lists/${r.list.title}/DispForm.aspx`,
  });

  const fileBytes = (serverRelativeUrl: string): Buffer => {
    for (const list of new Set(lists.values())) {
      const item = list.items.find(it => String(it.FileRef || '').toLowerCase() === serverRelativeUrl.toLowerCase());
      if (!item) continue;
      const root = join(opts.fixturesDir, 'files', list.title);
      const file = normalize(join(root, String(item.FileLeafRef)));
      if (file.startsWith(root + sep) && existsSync(file)) return readFileSync(file);
    }
    throw new HttpError(404, 'File Not Found.', '-2147024894, System.IO.FileNotFoundException');
  };

  const fieldsOf = (list: ListFixture): Item[] =>
    Object.keys(list.fields).map(name => {
      const def = list.fields[name];
      const target = def.list ? findList(def.list) : undefined;
      return {
        InternalName: name,
        Title: def.title || name,
        TypeAsString: def.type,
        ReadOnlyField: !!def.readOnly,
        Required: !!def.required,
        Choices: def.choices || [],
        DisplayFormat: def.type === 'DateTime' ? (def.dateOnly ? 0 : 1) : undefined,
        LookupList: target ? `{${target.id}}` : undefined,
        LookupField: def.type === 'Lookup' || def.type === 'LookupMulti' ? def.field || 'Title' : undefined,
        Description: def.description || '',
        Hidden: false,
      };
    });

  const queryItems = (r: Route, q: URLSearchParams, host: string | undefined, pathname: string): Item => {
    const select = splitList(q.get('$select'));
    const expand = splitList(q.get('$expand'));
    const top = Math.min(5000, parseInt(q.get('$top') || '100', 10) || 100);
    const skip = parseInt((q.get('$skiptoken') || '').replace(/^.*p_ID=/, ''), 10) || 0;
    let rows = r.list.items.map(it => expandItem(r.list, it));
    const filter = q.get('$filter');
    if (filter) {
      const expr = odataFilterToExpr(filter);
      try {
        rows = rows.filter(row => !!evaluate(expr, { row }));
      } catch (e) {
        throw new HttpError(400, `Invalid $filter "${filter}": ${(e as Error).message}`);
      }
    }
    const orderBy = q.get('$orderby');
    const keys = orderBy
      ? orderBy.split(',').map(k => {
          const [field, dir] = k.trim().split(/\s+/);
          return { path: field.split('/'), desc: (dir || '').toLowerCase() === 'desc' };
        })
      : [{ path: ['Id'], desc: false }];
    rows.sort((a, b) => {
      for (const k of keys) {
        const get = (row: Item): unknown => k.path.reduce<unknown>((o, p) => (o && typeof o === 'object' ? (o as Item)[p] : undefined), row);
        const c = compare(get(a), get(b));
        if (c) return k.desc ? -c : c;
      }
      return 0;
    });
    const body: Item = { value: rows.slice(skip, skip + top).map(row => project(row, select, r.list, expand)) };
    if (rows.length > skip + top) {
      const next = new URLSearchParams(q);
      next.set('$skiptoken', `Paged=TRUE&p_ID=${skip + top}`);
      body['odata.nextLink'] = `${host ? `http://${host}` : ''}${pathname}?${next.toString()}`;
    }
    return body;
  };

  const findItem = (r: Route, id: number): Item => {
    const item = r.list.items.find(it => it.Id === id);
    if (!item) throw new HttpError(404, 'Item does not exist. It may have been deleted by another user.', '-2130575338, System.ArgumentException');
    return item;
  };

  const updateItem = (r: Route, id: number, body: Item): void => {
    const item = findItem(r, id);
    for (const key of Object.keys(body)) {
      const lookupName = key.replace(/Id$/, '');
      const def = r.list.fields[key] || (key !== lookupName ? r.list.fields[lookupName] : undefined);
      if (!def) throw new HttpError(400, `The property '${key}' does not exist on type 'SP.Data.${r.list.title}ListItem'. Make sure to only use property names that are defined by the type.`);
      if (def.readOnly) throw new HttpError(400, `Field '${key}' is read-only.`);
      if (EXPANDABLE.indexOf(def.type) >= 0 && key === lookupName) throw new HttpError(400, `Set lookup and person fields through '${key}Id'.`);
      if (def.type === 'Choice' && body[key] !== null && def.choices && def.choices.indexOf(String(body[key])) < 0) {
        throw new HttpError(400, `'${String(body[key])}' is not a valid choice for ${def.title || key}.`);
      }
      item[key] = body[key];
    }
    item.Modified = new Date().toISOString();
    item.EditorId = opts.currentUser.id;
    changed.add(`${r.list.id}:${id}`);
  };

  const commentJson = (c: CommentFixture): Item => {
    const user = findUser(c.authorId);
    return {
      id: String(c.id),
      text: c.text,
      createdDate: c.createdDate,
      author: { id: c.authorId, name: user ? user.Title : 'Unknown user', email: user ? user.EMail : '' },
    };
  };

  const seed = (): Item => {
    const unique = Array.from(new Set(lists.values()));
    return {
      lists: unique.map(list => ({
        title: list.title,
        id: list.id,
        baseTemplate: list.baseTemplate || 100,
        fields: list.fields,
        items: list.items,
        comments: Object.fromEntries(list.items.map(it => [String(it.Id), (comments.get(`${list.id}:${it.Id}`) || []).length])),
        changed: list.items.filter(it => changed.has(`${list.id}:${it.Id}`)).map(it => it.Id),
      })),
      users,
    };
  };

  const handleDev = (req: IncomingMessage, res: ServerResponse, pathname: string): boolean => {
    if (pathname === '/__sp-mock/seed' && req.method === 'GET') {
      send(res, 200, seed());
      return true;
    }
    if (pathname === '/__sp-mock/reset' && req.method === 'POST') {
      load();
      notify();
      send(res, 204);
      return true;
    }
    return false;
  };

  const handle = async (req: IncomingMessage, res: ServerResponse): Promise<boolean> => {
    const url = new URL(req.url || '/', 'http://localhost');
    const pathname = decodeURIComponent(url.pathname);
    if (pathname.indexOf('/__sp-mock/') === 0) return handleDev(req, res, pathname);
    if (!/\/_api\//i.test(pathname)) return false;
    const role: Role = String(req.headers['x-mock-role'] || 'owner') === 'member' ? 'member' : 'owner';
    if (/\/_api\/web\/EffectiveBasePermissions$/i.test(pathname)) return reply(res, 200, role === 'owner' ? FULL_CONTROL : MEMBER);
    if (/\/_api\/web\/currentuser\/groups$/i.test(pathname)) return reply(res, 200, { value: GROUPS[role].map(Title => ({ Title })) });
    if (/\/_api\/web\/siteusers$/i.test(pathname)) {
      return reply(res, 200, { value: users.map(u => ({ Id: u.Id, Title: u.Title, Email: u.EMail, PrincipalType: 1 })) });
    }
    const file =
      /\/_api\/web\/GetFileByServerRelativePath\(decodedurl='((?:[^']|'')+)'\)\/\$value$/i.exec(pathname) ||
      /\/_api\/web\/GetFileByServerRelativeUrl\('((?:[^']|'')+)'\)\/\$value$/i.exec(pathname);
    if (file) {
      const bytes = fileBytes(file[1].replace(/''/g, "'"));
      res.statusCode = 200;
      res.setHeader('Content-Type', 'application/octet-stream');
      res.end(bytes);
      return true;
    }
    const r = route(pathname);
    if (!r) throw new HttpError(404, `The workbench SharePoint mock does not implement ${pathname}.`);
    const method = String(req.headers['x-http-method'] || req.method || 'GET').toUpperCase();
    const rest = r.rest.replace(/\/$/, '');

    if (rest === '' && method === 'GET') return reply(res, 200, listInfo(r));
    if (rest === '/EffectiveBasePermissions' && method === 'GET') return reply(res, 200, role === 'owner' ? FULL_CONTROL : MEMBER);
    if (rest === '/fields' && method === 'GET') return reply(res, 200, { value: fieldsOf(r.list) });
    if (rest === '/items' && method === 'GET') return reply(res, 200, queryItems(r, url.searchParams, req.headers.host, url.pathname));
    if (rest === '/items' && method === 'POST' && r.list.baseTemplate === 101) {
      throw new HttpError(400, 'To add an item to a document library, use SPFileCollection.Add()', '-2130575245, Microsoft.SharePoint.SPException');
    }
    if (rest === '/items' && method === 'POST') {
      const body = writePayload(r.list, req, await readBody(req));
      const now = new Date().toISOString();
      const id = r.list.items.reduce((m, it) => Math.max(m, Number(it.Id) || 0), 0) + 1;
      const item: Item = { Id: id, ID: id, Created: now, Modified: now, AuthorId: opts.currentUser.id, EditorId: opts.currentUser.id };
      r.list.items.push(item);
      try {
        updateItem(r, id, body);
      } catch (e) {
        r.list.items.splice(r.list.items.indexOf(item), 1);
        throw e;
      }
      const created = project(expandItem(r.list, item), [], r.list, []);
      return reply(res, 201, isVerbose(req.headers.accept) ? { d: created } : created);
    }

    const itemMatch = /^\/items\((\d+)\)(\/Comments(?:\(\))?)?$/i.exec(rest);
    if (itemMatch) {
      const id = Number(itemMatch[1]);
      const key = `${r.list.id}:${id}`;
      if (itemMatch[2]) {
        findItem(r, id);
        if (method === 'GET') {
          const list = (comments.get(key) || []).slice().sort((a, b) => b.createdDate.localeCompare(a.createdDate));
          return reply(res, 200, { value: list.map(commentJson) });
        }
        if (method === 'POST') {
          const body = await readBody(req);
          const text = String(body.text || '').trim();
          if (!text) throw new HttpError(400, 'Comment text cannot be empty.');
          const list = comments.get(key) || [];
          const all = Array.from(comments.values()).reduce<CommentFixture[]>((a, b) => a.concat(b), []);
          const comment: CommentFixture = { id: all.reduce((m, c) => Math.max(m, c.id), 0) + 1, text, createdDate: new Date().toISOString(), authorId: opts.currentUser.id };
          comments.set(key, list.concat(comment));
          changed.add(key);
          return reply(res, 201, commentJson(comment));
        }
      } else {
        if (method === 'GET') {
          const select = splitList(url.searchParams.get('$select'));
          const expand = splitList(url.searchParams.get('$expand'));
          return reply(res, 200, project(expandItem(r.list, findItem(r, id)), select, r.list, expand));
        }
        if (method === 'MERGE' || method === 'PATCH') {
          updateItem(r, id, writePayload(r.list, req, await readBody(req)));
          return reply(res, 204);
        }
      }
    }
    throw new HttpError(405, `The workbench SharePoint mock does not implement ${method} ${pathname}.`);
  };

  const reply = (res: ServerResponse, status: number, body?: unknown): Promise<boolean> => {
    const [lo, hi] = opts.latency || [120, 450];
    return new Promise(resolve => {
      setTimeout(() => {
        send(res, status, body);
        resolve(true);
      }, lo + Math.random() * (hi - lo));
    });
  };

  return {
    name: 'sp-mock',
    configureServer(server: ViteDevServer) {
      notify = () => server.ws.send({ type: 'custom', event: 'sp-mock:update', data: { file: 'reset' } });
      server.watcher.add(opts.fixturesDir);
      server.watcher.on('change', file => {
        if (file.replace(/\\/g, '/').indexOf(opts.fixturesDir.replace(/\\/g, '/')) === 0) {
          load();
          server.ws.send({ type: 'custom', event: 'sp-mock:update', data: { file } });
          server.config.logger.info(`[sp-mock] reloaded fixtures (${file.split(/[\\/]/).pop()})`, { timestamp: true });
        }
      });
      server.middlewares.use((req, res, next) => {
        handle(req, res).then(
          handled => {
            if (!handled) next();
          },
          e => sendError(res, e),
        );
      });
    },
  };
}
