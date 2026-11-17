
export type Node =
  | { k: 'lit'; v: unknown }
  | { k: 'id'; name: string }
  | { k: 'arr'; items: Node[] }
  | { k: 'get'; obj: Node; key: Node }
  | { k: 'un'; op: string; a: Node }
  | { k: 'bin'; op: string; a: Node; b: Node }
  | { k: 'cond'; c: Node; a: Node; b: Node }
  | { k: 'call'; name: string; args: Node[] };

export class ExpressionError extends Error {
  public constructor(message: string, public readonly expression: string, public readonly position?: number) {
    super(position !== undefined ? `${message} (at ${position + 1} in "${expression}")` : `${message} in "${expression}"`);
    this.name = 'ExpressionError';
  }
}

type TokType = 'num' | 'str' | 'id' | 'op' | 'eof';
interface Tok { t: TokType; v: string; pos: number }

const WORD_OPS: Record<string, string> = { and: '&&', or: '||', not: '!', in: 'in' };
const OPS3 = ['===', '!=='];
const OPS2 = ['==', '!=', '<=', '>=', '&&', '||', '??', '<>'];
const OPS1 = '+-*/%<>!?:()[],.=';
const NUMBER = /^[0-9]+(\.[0-9]+)?([eE][+-]?[0-9]+)?/;
const IDENTIFIER = /^[A-Za-z_$@][A-Za-z0-9_$@]*/;

function tokenize(src: string): Tok[] {
  const toks: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const c = src[i];
    if (/\s/.test(c)) { i++; continue; }
    const start = i;
    const num = NUMBER.exec(src.slice(i));
    if (num) {
      toks.push({ t: 'num', v: num[0], pos: start });
      i += num[0].length;
      continue;
    }
    if (c === '"' || c === "'") {
      let s = '';
      i++;
      while (i < src.length && src[i] !== c) {
        if (src[i] === '\\' && i + 1 < src.length) {
          const n = src[i + 1];
          s += n === 'n' ? '\n' : n === 't' ? '\t' : n;
          i += 2;
        } else {
          s += src[i++];
        }
      }
      if (i >= src.length) throw new ExpressionError('Unterminated string', src, start);
      i++;
      toks.push({ t: 'str', v: s, pos: start });
      continue;
    }
    if (c === '`') {
      const end = src.indexOf('`', i + 1);
      if (end < 0) throw new ExpressionError('Unterminated `quoted name`', src, start);
      toks.push({ t: 'id', v: src.slice(i + 1, end), pos: start });
      i = end + 1;
      continue;
    }
    const ident = IDENTIFIER.exec(src.slice(i));
    if (ident) {
      const w = ident[0];
      const lw = w.toLowerCase();
      if (WORD_OPS[lw]) toks.push({ t: 'op', v: WORD_OPS[lw], pos: start });
      else toks.push({ t: 'id', v: w, pos: start });
      i += w.length;
      continue;
    }
    const three = src.substr(i, 3);
    if (OPS3.indexOf(three) >= 0) { toks.push({ t: 'op', v: three.slice(0, 2), pos: start }); i += 3; continue; }
    const two = src.substr(i, 2);
    if (OPS2.indexOf(two) >= 0) { toks.push({ t: 'op', v: two === '<>' ? '!=' : two, pos: start }); i += 2; continue; }
    if (OPS1.indexOf(c) >= 0) { toks.push({ t: 'op', v: c === '=' ? '==' : c, pos: start }); i++; continue; }
    throw new ExpressionError(`Unexpected character '${c}'`, src, i);
  }
  toks.push({ t: 'eof', v: '', pos: src.length });
  return toks;
}

class Parser {
  private i = 0;
  public constructor(private readonly toks: Tok[], private readonly src: string) {}

  public parseAll(): Node {
    if (this.peek().t === 'eof') throw new ExpressionError('Empty expression', this.src, 0);
    const n = this.ternary();
    const t = this.peek();
    if (t.t !== 'eof') throw new ExpressionError(`Unexpected '${t.v}'`, this.src, t.pos);
    return n;
  }

  private peek(): Tok { return this.toks[this.i]; }
  private next(): Tok { return this.toks[this.i++]; }
  private isOp(v: string): boolean { const t = this.peek(); return t.t === 'op' && t.v === v; }
  private eat(v: string): boolean { if (this.isOp(v)) { this.i++; return true; } return false; }
  private expect(v: string): void {
    const t = this.peek();
    if (!this.eat(v)) throw new ExpressionError(`Expected '${v}' but found '${t.v || 'end'}'`, this.src, t.pos);
  }

  private ternary(): Node {
    const c = this.nullish();
    if (this.eat('?')) {
      const a = this.ternary();
      this.expect(':');
      const b = this.ternary();
      return { k: 'cond', c, a, b };
    }
    return c;
  }

  private binLevel(ops: string[], nextLevel: () => Node): Node {
    let a = nextLevel();
    for (;;) {
      const t = this.peek();
      if (t.t === 'op' && ops.indexOf(t.v) >= 0) {
        this.i++;
        a = { k: 'bin', op: t.v, a, b: nextLevel() };
      } else {
        return a;
      }
    }
  }

  private nullish(): Node { return this.binLevel(['??'], () => this.or()); }
  private or(): Node { return this.binLevel(['||'], () => this.and()); }
  private and(): Node { return this.binLevel(['&&'], () => this.equality()); }
  private equality(): Node { return this.binLevel(['==', '!='], () => this.comparison()); }
  private comparison(): Node { return this.binLevel(['<', '<=', '>', '>=', 'in'], () => this.additive()); }
  private additive(): Node { return this.binLevel(['+', '-'], () => this.multiplicative()); }
  private multiplicative(): Node { return this.binLevel(['*', '/', '%'], () => this.unary()); }

  private unary(): Node {
    if (this.eat('!')) return { k: 'un', op: '!', a: this.unary() };
    if (this.eat('-')) return { k: 'un', op: '-', a: this.unary() };
    if (this.eat('+')) return this.unary();
    return this.postfix();
  }

  private postfix(): Node {
    let n = this.primary();
    for (;;) {
      if (this.eat('.')) {
        const t = this.next();
        if (t.t !== 'id' && t.t !== 'num') throw new ExpressionError('Expected a field name after "."', this.src, t.pos);
        n = { k: 'get', obj: n, key: { k: 'lit', v: t.v } };
      } else if (this.isOp('[') ) {
        this.i++;
        const key = this.ternary();
        this.expect(']');
        n = { k: 'get', obj: n, key };
      } else {
        return n;
      }
    }
  }

  private primary(): Node {
    const t = this.next();
    switch (t.t) {
      case 'num': return { k: 'lit', v: parseFloat(t.v) };
      case 'str': return { k: 'lit', v: t.v };
      case 'id': {
        const lw = t.v.toLowerCase();
        if (lw === 'true') return { k: 'lit', v: true };
        if (lw === 'false') return { k: 'lit', v: false };
        if (lw === 'null') return { k: 'lit', v: null };
        if (this.isOp('(')) {
          this.i++;
          const args: Node[] = [];
          if (!this.isOp(')')) {
            do { args.push(this.ternary()); } while (this.eat(','));
          }
          this.expect(')');
          return { k: 'call', name: t.v, args };
        }
        return { k: 'id', name: t.v };
      }
      case 'op':
        if (t.v === '(') {
          const n = this.ternary();
          this.expect(')');
          return n;
        }
        if (t.v === '[') {
          const items: Node[] = [];
          if (!this.isOp(']')) {
            do { items.push(this.ternary()); } while (this.eat(','));
          }
          this.expect(']');
          return { k: 'arr', items };
        }
        throw new ExpressionError(`Unexpected '${t.v}'`, this.src, t.pos);
      default:
        throw new ExpressionError('Unexpected end of expression', this.src, t.pos);
    }
  }
}

const cache = new Map<string, Node>();

export function parse(expression: string): Node {
  let n = cache.get(expression);
  if (!n) {
    n = new Parser(tokenize(expression), expression).parseAll();
    if (cache.size > 2000) cache.clear();
    cache.set(expression, n);
  }
  return n;
}

export function functionsUsed(n: Node, out: string[] = []): string[] {
  switch (n.k) {
    case 'call': out.push(n.name); n.args.forEach(a => functionsUsed(a, out)); break;
    case 'arr': n.items.forEach(a => functionsUsed(a, out)); break;
    case 'get': functionsUsed(n.obj, out); functionsUsed(n.key, out); break;
    case 'un': functionsUsed(n.a, out); break;
    case 'bin': functionsUsed(n.a, out); functionsUsed(n.b, out); break;
    case 'cond': functionsUsed(n.c, out); functionsUsed(n.a, out); functionsUsed(n.b, out); break;
    default: break;
  }
  return out;
}
