type Token =
  | { kind: 'string'; value: string }
  | { kind: 'datetime'; value: string }
  | { kind: 'word'; value: string }
  | { kind: 'punct'; value: '(' | ')' | ',' };

const OPERATORS: Record<string, string> = { eq: '==', ne: '!=', gt: '>', ge: '>=', lt: '<', le: '<=', and: '&&', or: '||', not: '!' };
const LITERALS = ['true', 'false', 'null'];

function tokenize(filter: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  while (i < filter.length) {
    const c = filter[i];
    if (/\s/.test(c)) {
      i++;
      continue;
    }
    if (c === '(' || c === ')' || c === ',') {
      tokens.push({ kind: 'punct', value: c });
      i++;
      continue;
    }
    const datetime = /^datetime'([^']*)'/i.exec(filter.slice(i));
    if (datetime) {
      tokens.push({ kind: 'datetime', value: datetime[1] });
      i += datetime[0].length;
      continue;
    }
    if (c === "'") {
      let value = '';
      let j = i + 1;
      for (;;) {
        if (j >= filter.length) throw new Error(`Unterminated string in $filter at ${i}`);
        if (filter[j] === "'" && filter[j + 1] === "'") {
          value += "'";
          j += 2;
        } else if (filter[j] === "'") {
          break;
        } else {
          value += filter[j++];
        }
      }
      tokens.push({ kind: 'string', value });
      i = j + 1;
      continue;
    }
    const word = /^[\w.-]+(\/[A-Za-z_]\w*)*/.exec(filter.slice(i));
    if (!word) throw new Error(`Unexpected '${c}' in $filter at ${i}`);
    tokens.push({ kind: 'word', value: word[0] });
    i += word[0].length;
  }
  return tokens;
}

class Translator {
  private pos = 0;

  public constructor(private readonly tokens: Token[]) {}

  public translate(): string {
    const parts: string[] = [];
    while (this.pos < this.tokens.length) parts.push(this.term());
    return parts.join(' ');
  }

  private term(): string {
    const t = this.tokens[this.pos++];
    switch (t.kind) {
      case 'string':
        return JSON.stringify(t.value);
      case 'datetime':
        return `date(${JSON.stringify(t.value)})`;
      case 'punct':
        if (t.value === '(') return `(${this.until(')').join(' ')})`;
        throw new Error(`Unexpected '${t.value}' in $filter`);
      case 'word':
        return this.word(t.value);
    }
  }

  private word(value: string): string {
    const lower = value.toLowerCase();
    if (OPERATORS[lower]) return OPERATORS[lower];
    if (LITERALS.indexOf(lower) >= 0) return lower;
    if (/^-?\d/.test(value)) return value;
    const next = this.tokens[this.pos];
    if (next && next.kind === 'punct' && next.value === '(') {
      this.pos++;
      return this.call(lower, this.arguments());
    }
    return value.split('/').map(p => '`' + p + '`').join('.');
  }

  private call(name: string, args: string[]): string {
    switch (name) {
      case 'substringof':
        return `contains(${args[1]}, ${args[0]})`;
      case 'startswith':
        return `startsWith(${args[0]}, ${args[1]})`;
      case 'endswith':
        return `endsWith(${args[0]}, ${args[1]})`;
      default:
        throw new Error(`The workbench mock does not support the OData function ${name}()`);
    }
  }

  private arguments(): string[] {
    const args: string[] = [];
    let current: string[] = [];
    for (;;) {
      const t = this.tokens[this.pos];
      if (!t) throw new Error('Unclosed ( in $filter');
      if (t.kind === 'punct' && t.value === ')') {
        this.pos++;
        args.push(current.join(' '));
        return args;
      }
      if (t.kind === 'punct' && t.value === ',') {
        this.pos++;
        args.push(current.join(' '));
        current = [];
        continue;
      }
      current.push(this.term());
    }
  }

  private until(close: ')'): string[] {
    const parts: string[] = [];
    for (;;) {
      const t = this.tokens[this.pos];
      if (!t) throw new Error('Unclosed ( in $filter');
      if (t.kind === 'punct' && t.value === close) {
        this.pos++;
        return parts;
      }
      parts.push(this.term());
    }
  }
}

export function odataFilterToExpr(filter: string): string {
  return new Translator(tokenize(filter)).translate();
}
