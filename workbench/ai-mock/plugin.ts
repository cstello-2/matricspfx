import type { Plugin, ViteDevServer } from 'vite';
import type { IncomingMessage, ServerResponse } from 'node:http';

const DIMENSIONS = 512;
const STOPWORDS = new Set(
  'a an and are as at be by for from has have in is it its of on or that the this to was were will with we you your our not no can should must may into about after before than then them they their there these those which who what when where how all any each other some such only also more most very'.split(' '),
);

function stem(word: string): string {
  for (const suffix of ['ations', 'ation', 'ings', 'ing', 'ies', 'ied', 'ed', 'es', 's']) {
    if (word.length > suffix.length + 3 && word.endsWith(suffix)) return word.slice(0, -suffix.length);
  }
  return word;
}

export function tokens(text: string): string[] {
  return (text.toLowerCase().match(/[a-z0-9]+/g) || []).filter(w => w.length > 1 && !STOPWORDS.has(w)).map(stem);
}

function hash(s: string): number {
  let h = 2166136261;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return h >>> 0;
}

export function embed(text: string): number[] {
  const v = new Array<number>(DIMENSIONS).fill(0);
  const counts = new Map<string, number>();
  for (const t of tokens(text)) counts.set(t, (counts.get(t) || 0) + 1);
  counts.forEach((count, token) => {
    const weight = 1 + Math.log(count);
    const h = hash(token);
    v[h % DIMENSIONS] += (h & 0x80000000 ? -1 : 1) * weight;
    if (token.length >= 5) {
      for (let i = 0; i + 4 <= token.length; i++) {
        const g = hash(`#${token.slice(i, i + 4)}`);
        v[g % DIMENSIONS] += (g & 0x80000000 ? -0.25 : 0.25) * weight;
      }
    }
  });
  const len = Math.sqrt(v.reduce((a, b) => a + b * b, 0)) || 1;
  return v.map(x => x / len);
}

interface ChatMessage {
  role: string;
  content: string;
}

export function chooseTags(prompt: string): string[] {
  const tagSection = /Tags:\n([\s\S]*?)\n\nDocument:\n([\s\S]*?)\n\nPick at most (\d+)/.exec(prompt);
  if (!tagSection) return [];
  const documentTokens = new Set(tokens(tagSection[2]));
  const max = Number(tagSection[3]);
  return tagSection[1]
    .split('\n')
    .map(line => /^- ([^:]+)(?::\s*(.*))?$/.exec(line))
    .filter((m): m is RegExpExecArray => !!m)
    .map(m => ({ name: m[1].trim(), overlap: tokens(`${m[1]} ${m[2] || ''}`).filter(t => documentTokens.has(t)).length }))
    .filter(c => c.overlap > 0)
    .sort((a, b) => b.overlap - a.overlap)
    .slice(0, max)
    .map(c => c.name);
}

function readJson(req: IncomingMessage): Promise<Record<string, unknown>> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    req.on('data', (c: Buffer) => chunks.push(c));
    req.on('end', () => {
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}'));
      } catch (e) {
        reject(e);
      }
    });
    req.on('error', reject);
  });
}

function send(res: ServerResponse, status: number, body: unknown): void {
  res.statusCode = status;
  res.setHeader('Content-Type', 'application/json');
  res.end(JSON.stringify(body));
}

export function aiMock(basePath: string): Plugin {
  return {
    name: 'ai-mock',
    configureServer(server: ViteDevServer) {
      server.middlewares.use((req, res, next) => {
        const path = (req.url || '').split('?')[0];
        if (req.method !== 'POST' || path.indexOf(basePath) !== 0) {
          next();
          return;
        }
        const route = path.slice(basePath.length);
        readJson(req).then(
          body => {
            if (route === '/embeddings') {
              const input = Array.isArray(body.input) ? body.input.map(String) : [String(body.input || '')];
              send(res, 200, {
                object: 'list',
                model: body.model,
                data: input.map((text, index) => ({ object: 'embedding', index, embedding: embed(text) })),
                usage: { prompt_tokens: 0, total_tokens: 0 },
              });
              return;
            }
            if (route === '/chat/completions') {
              const messages = Array.isArray(body.messages) ? (body.messages as ChatMessage[]) : [];
              const user = messages.filter(m => m.role === 'user').map(m => m.content).join('\n');
              send(res, 200, {
                object: 'chat.completion',
                model: body.model,
                choices: [{ index: 0, finish_reason: 'stop', message: { role: 'assistant', content: JSON.stringify({ tags: chooseTags(user) }) } }],
              });
              return;
            }
            send(res, 404, { error: { message: `The workbench AI mock does not implement ${route}` } });
          },
          (e: Error) => send(res, 400, { error: { message: `Invalid JSON: ${e.message}` } }),
        );
      });
    },
  };
}
