import { AutotagConfig, Row, TagDefinition } from '../types';
import { getPath, isNil } from '../expr/evaluate';
import { formatValue } from '../format';
import { SpField, SpListService, BasePermissions, PermissionKind, hasPermission } from '../sharepoint/listService';
import { canExtract, extractText, htmlToText } from '../text/extract';
import { OpenAiCompatibleClient, RankedTag, cosine } from './openai';

export interface VocabularyTag extends TagDefinition {
  id?: number;
}

export interface TagSuggestion {
  tag: string;
  score: number;
  chosen: boolean;
}

export interface ItemSuggestion {
  itemId: number;
  title: string;
  current: string[];
  suggested: TagSuggestion[];
  usedFileText: boolean;
  error?: string;
}

export const TAGGABLE_KINDS = ['choice', 'multichoice', 'lookup', 'multilookup', 'text', 'note'];
const DEFAULT_SOURCE_FIELDS = ['Title', 'Description'];
const SEPARATOR = /\s*[;,]\s*/;

export function currentTags(field: SpField, row: Row): string[] {
  const value = row[field.name];
  if (isNil(value) || value === '') return [];
  const list = Array.isArray(value) ? value : [value];
  if (field.kind === 'lookup' || field.kind === 'multilookup') {
    const display = field.lookupField || 'Title';
    return list.map(v => (v && typeof v === 'object' ? String((v as Row)[display] ?? '') : String(v))).filter(Boolean);
  }
  if (field.kind === 'text' || field.kind === 'note') return String(value).split(SEPARATOR).filter(Boolean);
  return list.map(String);
}

export function tagPayload(field: SpField, tags: string[], vocabulary: VocabularyTag[]): Row {
  switch (field.kind) {
    case 'multichoice':
      return { [field.name]: tags };
    case 'choice':
      return { [field.name]: tags.length ? tags[0] : null };
    case 'multilookup':
    case 'lookup': {
      const ids = tags
        .map(t => vocabulary.find(v => v.name === t))
        .filter((v): v is VocabularyTag => !!v && v.id !== undefined)
        .map(v => Number(v.id));
      return field.kind === 'lookup' ? { [`${field.name}Id`]: ids.length ? ids[0] : null } : { [`${field.name}Id`]: ids };
    }
    default:
      return { [field.name]: tags.join('; ') };
  }
}

export function mergeTags(field: SpField, current: string[], chosen: string[], merge: boolean): string[] {
  if (field.kind === 'choice' || field.kind === 'lookup') return chosen.slice(0, 1);
  if (!merge) return chosen.slice();
  const out = current.slice();
  for (const t of chosen) if (out.indexOf(t) < 0) out.push(t);
  return out;
}

export function rankTags(itemVector: number[], tagVectors: number[][], vocabulary: VocabularyTag[]): RankedTag[] {
  return vocabulary
    .map((t, i) => ({ name: t.name, description: t.description, score: cosine(itemVector, tagVectors[i]) }))
    .sort((a, b) => b.score - a.score);
}

function fieldText(row: Row, path: string): string {
  const value = getPath(row, path);
  if (isNil(value) || value === '') return '';
  const text = formatValue(value, 'text');
  return /<[a-z][\s\S]*>/i.test(text) ? htmlToText(text) : text;
}

interface ItemText {
  text: string;
  usedFileText: boolean;
  error?: string;
}

export interface AutotagAccess {
  allowed: boolean;
  reason: string;
}

export async function checkAutotagAccess(
  service: SpListService,
  config: AutotagConfig,
  user: { email?: string; loginName?: string } | undefined,
): Promise<AutotagAccess> {
  const allow = config.allow && config.allow.length ? config.allow : ['owners', 'listOwners'];
  const lower = allow.map(a => a.toLowerCase());
  let webPerms: BasePermissions | undefined;
  let listPerms: BasePermissions | undefined;
  if (lower.indexOf('owners') >= 0) {
    webPerms = await service.webPermissions();
    if (hasPermission(webPerms, PermissionKind.ManageWeb)) return { allowed: true, reason: 'site owner' };
  }
  if (lower.indexOf('listowners') >= 0) {
    listPerms = await service.listPermissions();
    if (hasPermission(listPerms, PermissionKind.ManageLists)) return { allowed: true, reason: 'list owner' };
  }
  const identities = [user && user.email, user && user.loginName].filter((x): x is string => !!x).map(x => x.toLowerCase());
  if (lower.some(a => identities.indexOf(a) >= 0)) return { allowed: true, reason: 'listed in autotag.allow' };
  const named = allow.filter(a => ['owners', 'listowners'].indexOf(a.toLowerCase()) < 0 && a.indexOf('@') < 0 && a.indexOf('|') < 0);
  if (named.length) {
    const groups = (await service.currentUserGroups()).map(g => g.toLowerCase());
    if (named.some(n => groups.indexOf(n.toLowerCase()) >= 0)) return { allowed: true, reason: 'member of an allowed group' };
  }
  return { allowed: false, reason: 'Only site owners, list owners and people or groups in autotag.allow can auto-tag.' };
}

export class AutoTagger {
  private vocabularyPromise: Promise<VocabularyTag[]> | undefined;
  private vectorsPromise: Promise<number[][]> | undefined;

  public constructor(
    private readonly service: SpListService,
    private readonly ai: OpenAiCompatibleClient,
    public readonly config: AutotagConfig,
    public readonly field: SpField,
  ) {}

  public get maxTags(): number {
    return this.config.maxTags || 3;
  }

  public get minScore(): number {
    return this.config.minScore === undefined ? 0.25 : this.config.minScore;
  }

  public vocabulary(): Promise<VocabularyTag[]> {
    if (!this.vocabularyPromise) this.vocabularyPromise = this.loadVocabulary();
    return this.vocabularyPromise;
  }

  private async loadVocabulary(): Promise<VocabularyTag[]> {
    const descriptions = this.config.tagDescriptions || {};
    const describe = (name: string, description?: string): VocabularyTag => ({ name, description: description || descriptions[name] });
    if (this.config.tags && this.config.tags.length) {
      const configured: VocabularyTag[] = this.config.tags.map(t => describe(t.name, t.description));
      if (this.field.kind !== 'lookup' && this.field.kind !== 'multilookup') return configured;
      const options = await this.service.lookupOptions(this.field);
      return configured.map(t => {
        const match = options.find(o => o.text === t.name);
        return match ? { ...t, id: match.id } : t;
      });
    }
    if (this.field.kind === 'choice' || this.field.kind === 'multichoice') return this.field.choices.map(c => describe(c));
    if (this.field.kind === 'lookup' || this.field.kind === 'multilookup') {
      return (await this.service.lookupOptions(this.field)).map(o => ({ ...describe(o.text), id: o.id }));
    }
    throw new Error(`Add "tags" to the autotag config: ${this.field.title} is a ${this.field.kind} column without a fixed list of values.`);
  }

  private tagVectors(): Promise<number[][]> {
    if (!this.vectorsPromise) {
      this.vectorsPromise = this.vocabulary().then(vocab => this.ai.embed(vocab.map(t => (t.description ? `${t.name}: ${t.description}` : t.name))));
    }
    return this.vectorsPromise;
  }

  public sourcePaths(): string[] {
    return this.config.sourceFields && this.config.sourceFields.length ? this.config.sourceFields : DEFAULT_SOURCE_FIELDS;
  }

  public async itemText(row: Row): Promise<ItemText> {
    const parts = this.sourcePaths().map(p => fieldText(row, p)).filter(Boolean);
    let usedFileText = false;
    const fileRef = typeof row.FileRef === 'string' ? row.FileRef : undefined;
    const fileName = typeof row.FileLeafRef === 'string' ? row.FileLeafRef : fileRef;
    if (this.config.includeFileText !== false && fileRef && fileName && canExtract(fileName)) {
      const content = await extractText(fileName, await this.service.fileContent(fileRef));
      if (content && content.trim()) {
        parts.push(content);
        usedFileText = true;
      }
    }
    const maxChars = this.config.maxChars || 6000;
    return { text: parts.join('\n\n').slice(0, maxChars), usedFileText };
  }

  public async suggest(rows: Row[], onProgress?: (done: number, total: number) => void): Promise<ItemSuggestion[]> {
    const vocab = await this.vocabulary();
    const vectors = await this.tagVectors();
    const results: ItemSuggestion[] = [];
    const batch = 16;
    for (let i = 0; i < rows.length; i += batch) {
      const slice = rows.slice(i, i + batch);
      const texts = await Promise.all(
        slice.map(r => this.itemText(r).catch((e: Error): ItemText => ({ text: '', usedFileText: false, error: `Could not read the file: ${e.message}` }))),
      );
      const embeddable = texts.map(t => t.text || ' ');
      const itemVectors = await this.ai.embed(embeddable);
      for (let j = 0; j < slice.length; j++) {
        const row = slice[j];
        const current = currentTags(this.field, row);
        const textResult = texts[j];
        const base: ItemSuggestion = {
          itemId: Number(row.Id),
          title: String(row.Title || row.FileLeafRef || row.Id),
          current,
          suggested: [],
          usedFileText: textResult.usedFileText,
          error: textResult.error,
        };
        if (!textResult.text.trim()) {
          results.push({ ...base, error: base.error || 'No text to tag' });
          continue;
        }
        const ranked = rankTags(itemVectors[j], vectors, vocab).filter(t => t.score >= this.minScore);
        const candidates = ranked.slice(0, Math.max(this.maxTags * 3, 8));
        let picked = candidates.slice(0, this.maxTags).map(c => c.name);
        if (this.ai.canRerank && candidates.length) {
          picked = await this.ai.chooseTags(textResult.text, candidates, this.maxTags);
        }
        const suggested = candidates
          .filter(c => picked.indexOf(c.name) >= 0 && current.indexOf(c.name) < 0)
          .map(c => ({ tag: c.name, score: c.score, chosen: true }));
        results.push({ ...base, suggested });
      }
      if (onProgress) onProgress(Math.min(rows.length, i + batch), rows.length);
    }
    return results;
  }

  public async apply(row: Row, chosen: string[]): Promise<void> {
    const vocab = await this.vocabulary();
    const current = currentTags(this.field, row);
    const tags = mergeTags(this.field, current, chosen, this.config.merge !== false);
    await this.service.update(Number(row.Id), tagPayload(this.field, tags, vocab));
  }
}
