const TEXT_EXTENSIONS = ['txt', 'md', 'markdown', 'csv', 'tsv', 'json', 'log', 'xml', 'yaml', 'yml'];
const HTML_EXTENSIONS = ['html', 'htm', 'aspx'];

export function extensionOf(fileName: string): string {
  const m = /\.([A-Za-z0-9]+)$/.exec(fileName);
  return m ? m[1].toLowerCase() : '';
}

export function canExtract(fileName: string): boolean {
  const ext = extensionOf(fileName);
  return TEXT_EXTENSIONS.indexOf(ext) >= 0 || HTML_EXTENSIONS.indexOf(ext) >= 0 || ext === 'docx';
}

const ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };

export function decodeEntities(text: string): string {
  return text.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (whole, code: string) => {
    if (code.charAt(0) === '#') {
      const n = code.charAt(1).toLowerCase() === 'x' ? parseInt(code.slice(2), 16) : parseInt(code.slice(1), 10);
      return isNaN(n) ? whole : String.fromCodePoint(n);
    }
    const named = ENTITIES[code.toLowerCase()];
    return named === undefined ? whole : named;
  });
}

export function htmlToText(html: string): string {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  doc.querySelectorAll('script, style, noscript').forEach(el => el.remove());
  return (doc.body.textContent || '').replace(/\s+\n/g, '\n').trim();
}

export function docxXmlToText(xml: string): string {
  const out: string[] = [];
  const token = /<w:p[\s>]|<\/w:p>|<w:tab\/>|<w:br\/>|<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
  let m: RegExpExecArray | null;
  while ((m = token.exec(xml))) {
    const t = m[0];
    if (t === '</w:p>' || t === '<w:br/>') out.push('\n');
    else if (t === '<w:tab/>') out.push('\t');
    else if (m[1] !== undefined) out.push(decodeEntities(m[1]));
  }
  return out.join('').replace(/\n{3,}/g, '\n\n').trim();
}

interface ZipEntry {
  name: string;
  method: number;
  compressedSize: number;
  localOffset: number;
}

function readEntries(view: DataView): ZipEntry[] {
  let eocd = -1;
  for (let i = view.byteLength - 22; i >= Math.max(0, view.byteLength - 65557); i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error('Not a zip file');
  const count = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);
  const decoder = new TextDecoder();
  const entries: ZipEntry[] = [];
  for (let i = 0; i < count; i++) {
    if (view.getUint32(offset, true) !== 0x02014b50) throw new Error('Corrupt zip central directory');
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localOffset = view.getUint32(offset + 42, true);
    const name = decoder.decode(new Uint8Array(view.buffer, view.byteOffset + offset + 46, nameLength));
    entries.push({ name, method, compressedSize, localOffset });
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

async function inflateRaw(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

export async function readZipEntry(buffer: ArrayBuffer, name: string): Promise<Uint8Array | undefined> {
  const view = new DataView(buffer);
  const entry = readEntries(view).find(e => e.name === name);
  if (!entry) return undefined;
  const local = entry.localOffset;
  if (view.getUint32(local, true) !== 0x04034b50) throw new Error('Corrupt zip local header');
  const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
  const data = new Uint8Array(buffer, start, entry.compressedSize);
  if (entry.method === 0) return data.slice();
  if (entry.method === 8) return inflateRaw(data);
  throw new Error(`Unsupported zip compression method ${entry.method}`);
}

export async function extractText(fileName: string, buffer: ArrayBuffer): Promise<string | undefined> {
  const ext = extensionOf(fileName);
  if (TEXT_EXTENSIONS.indexOf(ext) >= 0) return new TextDecoder().decode(buffer);
  if (HTML_EXTENSIONS.indexOf(ext) >= 0) return htmlToText(new TextDecoder().decode(buffer));
  if (ext === 'docx') {
    const xml = await readZipEntry(buffer, 'word/document.xml');
    return xml ? docxXmlToText(new TextDecoder().decode(xml)) : undefined;
  }
  return undefined;
}
