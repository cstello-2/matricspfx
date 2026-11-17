import { describe, it, expect } from 'vitest';
import { evaluate, evaluateAggregate } from '../../src/core/expr/evaluate';
import { parse } from '../../src/core/expr/parser';

const row = { Hours: 4, Rate: 85, Status: 'Open', Done: false, Tags: ['a', 'b'], project: { Title: 'Apollo', Id: 3 }, 'Due Date': '2026-03-15' };
const ev = (e: string, r: Record<string, unknown> = row, vars?: Record<string, unknown>) => evaluate(e, { row: r, vars });

describe('expressions', () => {
  it('does arithmetic with precedence', () => {
    expect(ev('Hours * Rate + 10')).toBe(350);
    expect(ev('(Hours + 1) * 2')).toBe(10);
    expect(ev('-Hours % 3')).toBe(-1);
    expect(ev('Hours / 0')).toBeNull();
  });

  it('compares and combines logic', () => {
    expect(ev("Status == 'Open' && !Done")).toBe(true);
    expect(ev("Status = 'open'")).toBe(false);
    expect(ev("Status in ['Open', 'Blocked']")).toBe(true);
    expect(ev("'a' in Tags")).toBe(true);
    expect(ev('Hours > 3 and Hours <= 4')).toBe(true);
    expect(ev("Hours > 5 ? 'big' : 'small'")).toBe('small');
    expect(ev('Missing ?? 7')).toBe(7);
  });

  it('reads nested paths, quoted names and is case-insensitive as fallback', () => {
    expect(ev('project.Title')).toBe('Apollo');
    expect(ev('PROJECT.title')).toBe('Apollo');
    expect(ev('`Due Date`')).toBe('2026-03-15');
    expect(ev('Tags[1]')).toBe('b');
  });

  it('has string and date functions', () => {
    expect(ev("upper(left(Status, 2))")).toBe('OP');
    expect(ev("contains(project.Title, 'pol')")).toBe(true);
    expect(ev('month(`Due Date`)')).toBe(3);
    expect(ev("bucket(`Due Date`, 'quarter')")).toBe('2026-Q1');
    expect(ev("daysBetween('2026-03-01', `Due Date`)")).toBe(14);
    expect(ev("formatDate(`Due Date`, 'dd/MM/yyyy')")).toBe('15/03/2026');
  });

  it('reads variables', () => {
    expect(ev('$user.email', row, { user: { email: 'me@x.com' } })).toBe('me@x.com');
  });

  it('aggregates over rows', () => {
    const rows = [{ h: 1, s: 'a' }, { h: 2, s: 'b' }, { h: 3, s: 'a' }, { h: null, s: 'c' }];
    expect(evaluateAggregate('sum(h)', rows)).toBe(6);
    expect(evaluateAggregate('avg(h)', rows)).toBe(2);
    expect(evaluateAggregate('count()', rows)).toBe(4);
    expect(evaluateAggregate('count(h)', rows)).toBe(3);
    expect(evaluateAggregate("countIf(s == 'a')", rows)).toBe(2);
    expect(evaluateAggregate('distinct(s)', rows)).toBe(3);
    expect(evaluateAggregate('round(sum(h) / count(), 2)', rows)).toBe(1.5);
    expect(evaluateAggregate('max(h)', rows)).toBe(3);
    expect(evaluateAggregate('median(h)', rows)).toBe(2);
    expect(evaluateAggregate('max(1, 5)', rows)).toBe(5);
  });

  it('aggregates arrays in row context (one-to-many joins)', () => {
    const r = { tasks: [{ Hours: 2 }, { Hours: 5 }] };
    expect(ev('sum(tasks.Hours)', r)).toBe(7);
    expect(ev('count(tasks)', r)).toBe(2);
  });

  it('reports syntax errors with position', () => {
    expect(() => parse('Hours * ')).toThrow(/Unexpected end/);
    expect(() => parse("Status == 'Open")).toThrow(/Unterminated/);
    expect(() => ev('nope(1)')).toThrow(/Unknown function/);
  });
});
