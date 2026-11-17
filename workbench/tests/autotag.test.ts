import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { OpenAiCompatibleClient, cosine, parseTagReply, resolveApiUrl } from '../../src/core/ai/openai';
import { AutoTagger, checkAutotagAccess, currentTags, mergeTags, tagPayload } from '../../src/core/ai/autotag';
import { SpListService, hasPermission, toSpField } from '../../src/core/sharepoint/listService';
import { docxXmlToText, extractText, canExtract } from '../../src/core/text/extract';
import { IHttp, HttpPostOptions } from '../../src/core/datasources';
import { embed, chooseTags } from '../ai-mock/plugin';
import { AutotagConfig, Row } from '../../src/core/types';

const fixtures = join(__dirname, '..', 'fixtures');

function aiHttp(files: Record<string, ArrayBuffer> = {}, perms = { web: { High: 0, Low: 0 }, list: { High: 0, Low: 0 }, groups: [] as string[] }): { http: IHttp; posts: Array<{ url: string; body: unknown }> } {
  const posts: Array<{ url: string; body: unknown }> = [];
  const http: IHttp = {
    getJson: url => {
      if (/\/_api\/web\/EffectiveBasePermissions$/.test(url)) return Promise.resolve({ High: String(perms.web.High), Low: String(perms.web.Low) });
      if (/EffectiveBasePermissions$/.test(url)) return Promise.resolve({ High: String(perms.list.High), Low: String(perms.list.Low) });
      if (/currentuser\/groups/.test(url)) return Promise.resolve({ value: perms.groups.map(Title => ({ Title })) });
      return Promise.resolve({ value: [] });
    },
    postJson: (url, options: HttpPostOptions) => {
      posts.push({ url, body: options.body });
      const body = options.body as { input?: string[]; messages?: Array<{ role: string; content: string }> };
      if (url.endsWith('/embeddings')) return Promise.resolve({ data: (body.input || []).map((t, index) => ({ index, embedding: embed(t) })).reverse() });
      if (url.endsWith('/chat/completions')) {
        const user = (body.messages || []).filter(m => m.role === 'user').map(m => m.content).join('\n');
        return Promise.resolve({ choices: [{ message: { content: `Sure:\n\`\`\`json\n${JSON.stringify({ tags: chooseTags(user) })}\n\`\`\`` } }] });
      }
      return Promise.resolve(undefined);
    },
    getBinary: url => {
      const m = /GetFileByServerRelativePath\(decodedurl='(.+)'\)/.exec(url);
      const key = m ? decodeURIComponent(m[1]) : '';
      return files[key] ? Promise.resolve(files[key]) : Promise.reject(new Error('File Not Found.'));
    },
  };
  return { http, posts };
}

const topics = toSpField({ InternalName: 'Topics', Title: 'Topics', TypeAsString: 'MultiChoice', Choices: ['HR', 'Finance', 'Travel', 'Security', 'Remote work'] });

const config: AutotagConfig = {
  field: 'Topics',
  api: { baseUrl: '/mock-ai/v1', embeddingModel: 'mock' },
  tagDescriptions: {
    HR: 'people, employees, onboarding, benefits, leave',
    Finance: 'money, budget, expenses, receipts, reimbursement',
    Travel: 'flights, hotels, trips, taxis, booking travel',
    Security: 'passwords, multi-factor authentication, phishing, VPN',
    'Remote work': 'working from home, home office, remote',
  },
  sourceFields: ['Title', 'Summary'],
  maxTags: 2,
  minScore: 0.05,
};

function toArrayBuffer(b: Buffer): ArrayBuffer {
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
}

describe('OpenAI-compatible client', () => {
  it('resolves relative base URLs against the site origin', () => {
    expect(resolveApiUrl('/mock-ai/v1/', 'https://contoso.sharepoint.com/sites/a')).toBe('https://contoso.sharepoint.com/mock-ai/v1');
    expect(resolveApiUrl('https://ai.example.com/v1', 'https://x')).toBe('https://ai.example.com/v1');
  });

  it('batches embeddings and restores input order', async () => {
    const { http, posts } = aiHttp();
    const client = new OpenAiCompatibleClient(http, { baseUrl: 'https://ai/v1', embeddingModel: 'm', batchSize: 2 }, 'https://x');
    const vectors = await client.embed(['apples', 'bananas', 'apples']);
    expect(posts.map(p => p.url)).toEqual(['https://ai/v1/embeddings', 'https://ai/v1/embeddings']);
    expect(cosine(vectors[0], vectors[2])).toBeCloseTo(1);
    expect(cosine(vectors[0], vectors[1])).toBeLessThan(0.5);
  });

  it('parses tag replies wrapped in prose or code fences, keeping only allowed tags', () => {
    expect(parseTagReply('```json\n{"tags": ["hr", "Unknown", "Finance"]}\n```', ['HR', 'Finance'])).toEqual(['HR', 'Finance']);
    expect(parseTagReply('no json here', ['HR'])).toEqual([]);
  });
});

describe('text extraction', () => {
  it('reads Word documents', async () => {
    const buffer = toArrayBuffer(readFileSync(join(fixtures, 'files', 'Policies', 'Remote work policy.docx')));
    const text = await extractText('Remote work policy.docx', buffer);
    expect(text).toContain('Employees may work remotely up to three days per week');
    expect(text).toContain('company VPN');
  });

  it('reads text formats and skips binaries', async () => {
    expect(await extractText('notes.md', toArrayBuffer(Buffer.from('# Title\nBody')))).toBe('# Title\nBody');
    expect(await extractText('car.pdf', toArrayBuffer(Buffer.from('%PDF')))).toBeUndefined();
    expect(canExtract('a.DOCX')).toBe(true);
    expect(canExtract('a.pdf')).toBe(false);
  });

  it('keeps paragraphs, tabs and entities from document XML', () => {
    expect(docxXmlToText('<w:body><w:p><w:r><w:t>A &amp; B</w:t><w:tab/><w:t xml:space="preserve">C</w:t></w:r></w:p><w:p><w:r><w:t>D</w:t></w:r></w:p></w:body>')).toBe('A & B\tC\nD');
  });
});

describe('auto-tagger', () => {
  const rows: Row[] = [
    { Id: 1, Title: 'Remote work policy', Summary: 'Working from home rules', FileRef: '/sites/dev/Policies/Remote work policy.docx', FileLeafRef: 'Remote work policy.docx', Topics: [] },
    { Id: 2, Title: 'Travel and expenses', Summary: 'Booking flights and hotels, claiming receipts', Topics: ['Finance'] },
    { Id: 3, Title: 'Password standard', Summary: 'Multi-factor authentication and phishing', Topics: [] },
    { Id: 4, Title: '', Summary: '', Topics: [] },
  ];
  const files = { '/sites/dev/Policies/Remote work policy.docx': toArrayBuffer(readFileSync(join(fixtures, 'files', 'Policies', 'Remote work policy.docx'))) };

  it('suggests the closest tags, skips tags already applied, and reads file text', async () => {
    const { http } = aiHttp(files);
    const ctx = { http, webUrl: 'https://contoso.sharepoint.com/sites/dev' };
    const service = new SpListService(ctx, { type: 'spList', list: 'Policies' });
    const tagger = new AutoTagger(service, new OpenAiCompatibleClient(http, config.api, ctx.webUrl), config, topics);
    const progress: number[] = [];
    const result = await tagger.suggest(rows, done => progress.push(done));
    expect(result[0].suggested.map(s => s.tag)).toContain('Remote work');
    expect(result[0].usedFileText).toBe(true);
    expect(result[1].suggested.map(s => s.tag)).toContain('Travel');
    expect(result[1].suggested.map(s => s.tag)).not.toContain('Finance');
    expect(result[2].suggested[0].tag).toBe('Security');
    expect(result[3].error).toBe('No text to tag');
    expect(progress).toEqual([4]);
  });

  it('lets a chat model pick among the closest tags', async () => {
    const { http, posts } = aiHttp();
    const ctx = { http, webUrl: 'https://x/sites/dev' };
    const service = new SpListService(ctx, { type: 'spList', list: 'Policies' });
    const withChat = { ...config, api: { ...config.api, chatModel: 'chat' } };
    const tagger = new AutoTagger(service, new OpenAiCompatibleClient(http, withChat.api, ctx.webUrl), withChat, topics);
    const [result] = await tagger.suggest([rows[2]]);
    expect(posts.some(p => p.url.endsWith('/chat/completions'))).toBe(true);
    expect(result.suggested.map(s => s.tag)).toEqual(['Security']);
  });

  it('merges and converts tags for each column type', () => {
    const lookup = toSpField({ InternalName: 'Topic', TypeAsString: 'LookupMulti', LookupList: '{x}' });
    const text = toSpField({ InternalName: 'Keywords', TypeAsString: 'Text' });
    expect(mergeTags(topics, ['HR'], ['Finance', 'HR'], true)).toEqual(['HR', 'Finance']);
    expect(mergeTags(topics, ['HR'], ['Finance'], false)).toEqual(['Finance']);
    expect(tagPayload(lookup, ['HR', 'Finance'], [{ name: 'HR', id: 4 }, { name: 'Finance', id: 7 }])).toEqual({ TopicId: [4, 7] });
    expect(tagPayload(text, ['a', 'b'], [])).toEqual({ Keywords: 'a; b' });
    expect(currentTags(text, { Keywords: 'a; b,c' })).toEqual(['a', 'b', 'c']);
  });
});

describe('auto-tag permissions', () => {
  const FULL = { High: 2147483647, Low: 4294967295 };
  const MEMBER = { High: 0, Low: 0b1111 };
  const service = (perms: Parameters<typeof aiHttp>[1]): SpListService => new SpListService({ http: aiHttp({}, perms).http, webUrl: 'https://x/sites/a' }, { type: 'spList', list: 'L' });

  it('reads permission bits', () => {
    expect(hasPermission(FULL, 31)).toBe(true);
    expect(hasPermission(MEMBER, 3)).toBe(true);
    expect(hasPermission(MEMBER, 12)).toBe(false);
  });

  it('allows site owners and list owners by default', async () => {
    expect((await checkAutotagAccess(service({ web: FULL, list: FULL, groups: [] }), config, undefined)).reason).toBe('site owner');
    expect((await checkAutotagAccess(service({ web: MEMBER, list: FULL, groups: [] }), config, undefined)).reason).toBe('list owner');
    expect((await checkAutotagAccess(service({ web: MEMBER, list: MEMBER, groups: [] }), config, undefined)).allowed).toBe(false);
  });

  it('allows configured people and groups', async () => {
    const members = { web: MEMBER, list: MEMBER, groups: ['Policy Editors'] };
    expect((await checkAutotagAccess(service(members), { ...config, allow: ['Policy Editors'] }, undefined)).allowed).toBe(true);
    expect((await checkAutotagAccess(service(members), { ...config, allow: ['ann@contoso.com'] }, { email: 'Ann@contoso.com' })).allowed).toBe(true);
    expect((await checkAutotagAccess(service(members), { ...config, allow: ['owners'] }, { email: 'ann@contoso.com' })).allowed).toBe(false);
  });
});
