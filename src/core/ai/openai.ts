import { AiApiConfig } from '../types';
import { IHttp } from '../datasources';

export interface RankedTag {
  name: string;
  description?: string;
  score: number;
}

interface EmbeddingResponse {
  data?: Array<{ index?: number; embedding?: number[] }>;
}

interface ChatResponse {
  choices?: Array<{ message?: { content?: string } }>;
}

export function resolveApiUrl(baseUrl: string, webUrl: string): string {
  const base = baseUrl.replace(/\/$/, '');
  if (/^https?:\/\//i.test(base)) return base;
  const origin = /^(https?:\/\/[^/]+)/i.exec(webUrl);
  return (origin ? origin[1] : '') + (base.charAt(0) === '/' ? base : `/${base}`);
}

export function normalize(v: number[]): number[] {
  let sum = 0;
  for (const x of v) sum += x * x;
  const len = Math.sqrt(sum);
  return len > 0 ? v.map(x => x / len) : v.slice();
}

export function cosine(a: number[], b: number[]): number {
  let dot = 0;
  let na = 0;
  let nb = 0;
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) {
    dot += a[i] * b[i];
    na += a[i] * a[i];
    nb += b[i] * b[i];
  }
  return na && nb ? dot / Math.sqrt(na * nb) : 0;
}

export function parseTagReply(content: string, allowed: string[]): string[] {
  const start = content.indexOf('{');
  const end = content.lastIndexOf('}');
  if (start < 0 || end <= start) return [];
  const parsed: unknown = JSON.parse(content.slice(start, end + 1));
  const tags = parsed && typeof parsed === 'object' ? (parsed as { tags?: unknown }).tags : undefined;
  if (!Array.isArray(tags)) return [];
  const byLower = new Map(allowed.map((t): [string, string] => [t.toLowerCase(), t]));
  const out: string[] = [];
  for (const t of tags) {
    const match = byLower.get(String(t).toLowerCase());
    if (match && out.indexOf(match) < 0) out.push(match);
  }
  return out;
}

export class OpenAiCompatibleClient {
  private readonly base: string;

  public constructor(private readonly http: IHttp, private readonly config: AiApiConfig, webUrl: string) {
    this.base = resolveApiUrl(config.baseUrl, webUrl);
  }

  public get canRerank(): boolean {
    return !!this.config.chatModel;
  }

  public async embed(texts: string[]): Promise<number[][]> {
    const size = Math.max(1, this.config.batchSize || 32);
    const out: number[][] = [];
    for (let i = 0; i < texts.length; i += size) {
      const batch = texts.slice(i, i + size);
      const res = (await this.http.postJson(`${this.base}/embeddings`, {
        sharepoint: false,
        headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
        body: { model: this.config.embeddingModel, input: batch },
      })) as EmbeddingResponse;
      const data = res && Array.isArray(res.data) ? res.data.slice() : [];
      if (data.length !== batch.length) throw new Error(`Embeddings API returned ${data.length} vectors for ${batch.length} inputs`);
      data.sort((a, b) => (a.index || 0) - (b.index || 0));
      for (const d of data) {
        if (!Array.isArray(d.embedding)) throw new Error('Embeddings API returned an item without an embedding');
        out.push(normalize(d.embedding));
      }
    }
    return out;
  }

  public async chooseTags(text: string, candidates: RankedTag[], maxTags: number): Promise<string[]> {
    if (!this.config.chatModel || !candidates.length) return candidates.slice(0, maxTags).map(c => c.name);
    const list = candidates.map(c => `- ${c.name}${c.description ? `: ${c.description}` : ''}`).join('\n');
    const res = (await this.http.postJson(`${this.base}/chat/completions`, {
      sharepoint: false,
      headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
      body: {
        model: this.config.chatModel,
        temperature: 0,
        messages: [
          {
            role: 'system',
            content: 'You assign tags to documents. Only use tags from the provided list. Reply with JSON only, in the form {"tags": ["tag"]}.',
          },
          {
            role: 'user',
            content: `Tags:\n${list}\n\nDocument:\n${text}\n\nPick at most ${maxTags} tags that clearly apply. Return an empty list if none do.`,
          },
        ],
      },
    })) as ChatResponse;
    const content = res && res.choices && res.choices[0] && res.choices[0].message ? res.choices[0].message.content || '' : '';
    return parseTagReply(content, candidates.map(c => c.name)).slice(0, maxTags);
  }
}
