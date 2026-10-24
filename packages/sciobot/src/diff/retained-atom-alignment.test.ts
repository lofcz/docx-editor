import { expect, test } from 'bun:test';
import { alignRetainedAtoms } from './retained-atom-alignment';
import type { CompiledInlineTarget } from './inline-target';
import type { TargetAstNode } from '../projection/parse-target';

type Symbol = 'a' | 'b' | 'x' | 'y' | 'i' | 'j';
const image = (value: Symbol) => value === 'i' || value === 'j';
const atom = (value: Symbol) => value === 'x' || value === 'y' || image(value);
const ast = (value: Symbol): TargetAstNode =>
  image(value) ? { type: 'image', url: 'asset:' + value } : { type: 'inlineMath', value };
function compiled(values: readonly Symbol[]): CompiledInlineTarget {
  return {
    ast: { type: 'paragraph', children: [] },
    hyperlinks: [],
    plainText: null,
    tokens: values.map((value, index) => ({
      path: [index],
      marks: [],
      hyperlink: null,
      ...(atom(value)
        ? {
            kind: 'atom' as const,
            ast: ast(value),
          }
        : { kind: 'text' as const, value }),
    })),
  };
}
function distance(a: readonly Symbol[], b: readonly Symbol[]): number {
  let row = b.map((_, i) => i + 1);
  row.unshift(0);
  for (let i = 0; i < a.length; i++) {
    const next = [i + 1];
    for (let j = 0; j < b.length; j++)
      next.push(a[i] === b[j] ? row[j]! : Math.min(row[j + 1]!, next[j]!) + 1);
    row = next;
  }
  return row[b.length]!;
}
function cost(
  a: readonly Symbol[],
  b: readonly Symbol[],
  matches: readonly number[],
  replacements = false
): number {
  const old = a.flatMap((value, index) => (atom(value) ? [index] : []));
  const next = b.flatMap((value, index) => (atom(value) ? [index] : []));
  let total = next.length - old.length,
    from = 0,
    to = 0;
  for (let i = 0; i <= old.length; i++) {
    const end = old[i] ?? a.length,
      target = i === old.length ? b.length : next[matches[i]!]!;
    total += distance(
      a.slice(from, end),
      b.slice(to, target).filter((value) => !atom(value))
    );
    from = end + 1;
    to = target + 1;
  }
  if (replacements)
    for (let i = 0; i < old.length; i++) if (a[old[i]!] !== b[next[matches[i]!]!]) total++;
  return total;
}
/** Enumerate every legal identity placement, then compare ordinary text gaps independently. */
function oracle(a: readonly Symbol[], b: readonly Symbol[], replacements = false) {
  const old = a.filter(atom),
    next = b.filter(atom);
  let best = Infinity;
  const visit = (matches: number[], start: number) => {
    if (matches.length === old.length) {
      best = Math.min(best, cost(a, b, matches, replacements));
      return;
    }
    for (let i = start; i < next.length; i++)
      if (
        next[i] === old[matches.length] ||
        (replacements && image(next[i]!) && image(old[matches.length]!))
      )
        visit([...matches, i], i + 1);
  };
  visit([], 0);
  return best;
}
test('repeated atom identities stay beside their original text, including a middle duplicate', () => {
  expect(
    alignRetainedAtoms(
      compiled(['a', 'x', 'b', 'x', 'a']),
      compiled(['x', 'a', 'x', 'b', 'x', 'x', 'a'])
    )
  ).toEqual([1, 2]);
  expect(
    alignRetainedAtoms(
      compiled(['a', 'x', 'b', 'x', 'a']),
      compiled(['a', 'x', 'x', 'b', 'x', 'a'])
    )
  ).toEqual([0, 2]);
  expect(
    alignRetainedAtoms(compiled(['x', 'a', 'b', 'y']), compiled(['a', 'b', 'x', 'x', 'y']))
  ).toEqual([0, 2]);
});
test('mandatory identities cannot disappear or reorder when text offers a cheaper lexical match', () => {
  expect(alignRetainedAtoms(compiled(['x', 'a', 'y']), compiled(['y', 'a', 'x', 'x']))).toBeNull();
  expect(alignRetainedAtoms(compiled(['x', 'x']), compiled(['x']))).toBeNull();
});
test('bounded small streams choose a minimum text-edit placement among every legal retained identity mapping', () => {
  const values: Symbol[][] = [[]];
  let level: Symbol[][] = [[]];
  for (let i = 0; i < 4; i++) {
    level = level.flatMap((prefix) =>
      (['a', 'b', 'x'] as const).map((value) => [...prefix, value])
    );
    values.push(...level);
  }
  let cases = 0;
  for (const a of values)
    for (const b of values) {
      const old = a.filter(atom),
        next = b.filter(atom);
      if (!old.length || next.length <= old.length) continue;
      const found = alignRetainedAtoms(compiled(a), compiled(b));
      expect(found).not.toBeNull();
      expect(found).toHaveLength(old.length);
      expect(cost(a, b, found!)).toBe(oracle(a, b));
      cases++;
    }
  expect(cases).toBeGreaterThan(1000);
});

test('image replacements and fresh atoms choose minimum edits without replacing other atom kinds', () => {
  expect(
    alignRetainedAtoms(compiled(['a', 'i', 'b', 'x']), compiled(['j', 'a', 'j', 'b', 'x']), true)
  ).toEqual([1, 2]);
  expect(alignRetainedAtoms(compiled(['i', 'x']), compiled(['j', 'y', 'j']), true)).toBeNull();
  const values: Symbol[][] = [[]];
  let level: Symbol[][] = [[]];
  for (let i = 0; i < 3; i++) {
    level = level.flatMap((prefix) =>
      (['a', 'i', 'j', 'x'] as const).map((value) => [...prefix, value])
    );
    values.push(...level);
  }
  let cases = 0;
  for (const a of values)
    for (const b of values) {
      const old = a.filter(atom),
        next = b.filter(atom);
      if (!old.some(image) || next.length <= old.length) continue;
      const minimum = oracle(a, b, true),
        found = alignRetainedAtoms(compiled(a), compiled(b), true);
      if (minimum === Infinity) expect(found).toBeNull();
      else {
        expect(found).toHaveLength(old.length);
        expect(cost(a, b, found!, true)).toBe(minimum);
      }
      cases++;
    }
  expect(cases).toBeGreaterThan(500);
});
