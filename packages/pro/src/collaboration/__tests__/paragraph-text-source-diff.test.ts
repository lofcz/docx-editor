/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { diffSourceTokens } from '../document/paragraph-text-source-diff';
import type { Token } from '../document/paragraph-text-diff';
function text(value: string, sources: readonly (string | null)[]): Token[] {
  return value.split('').map((insert, index) => ({
    insert,
    key: `c${insert}`,
    strong: `c${insert}`,
    attributes: {},
    source: sources[index] ?? null,
  }));
}
test('new labels cannot borrow identically spelled original source characters', () => {
  const before = text('and', ['a', 'n', 'd']);
  const after = text('and', [null, null, null]);
  const steps = diffSourceTokens(before, after);
  expect(steps.filter((step) => step.op === 'eq')).toEqual([]);
  expect(steps.filter((step) => step.op === 'del')).toHaveLength(3);
  expect(steps.filter((step) => step.op === 'ins')).toHaveLength(3);
});
test('explicit middle-fragment sources retain their exact offsets between new labels', () => {
  const before = text('and', ['a', 'n', 'd']);
  const after = text('xandPaper', [null, 'a', 'n', 'd', null, null, null, null, null]);
  expect(diffSourceTokens(before, after).filter((step) => step.op === 'eq')).toEqual([
    { op: 'eq', before: 0, after: 1 },
    { op: 'eq', before: 1, after: 2 },
    { op: 'eq', before: 2, after: 3 },
  ]);
});
test('whole supplementary code points retain both original code units', () => {
  const before = text('A😀B', ['a', 'high', 'low', 'b']);
  const after = text('A😀😀B', ['a', null, null, 'high', 'low', 'b']);
  const steps = diffSourceTokens(before, after);
  expect(steps.filter((step) => step.op === 'eq')).toEqual([
    { op: 'eq', before: 0, after: 0 },
    { op: 'eq', before: 1, after: 3 },
    { op: 'eq', before: 2, after: 4 },
    { op: 'eq', before: 3, after: 5 },
  ]);
  expect(steps.filter((step) => step.op === 'ins')).toEqual([
    { op: 'ins', after: 1 },
    { op: 'ins', after: 2 },
  ]);
});
test('equal identities cannot retain a different visible character', () => {
  const before = text('and', ['a', 'n', 'd']);
  const after = text('end', ['a', 'n', 'd']);
  expect(diffSourceTokens(before, after).filter((step) => step.op === 'eq')).toEqual([
    { op: 'eq', before: 1, after: 1 },
    { op: 'eq', before: 2, after: 2 },
  ]);
});
