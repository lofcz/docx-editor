/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { writeInlineText } from '../document/paragraph-text-view';
import { isParagraphBoundary } from '../document/paragraph-text';
import { tokensOfText, type Token } from '../document/paragraph-text-diff';
import { paragraphInsertionPosition } from '../document/paragraph-control-replacement';

const plain: Token = { key: 'cx', strong: 'cx', insert: 'x', attributes: {} };
const control: Token = {
  ...plain,
  attributes: { w: '[{"i":"control","k":"contentControl","l":"sdt"}]' },
};
function boundaryText() {
  const doc = new Y.Doc();
  return { doc, text: writeInlineText(doc.getMap('record'), []) };
}
test('paragraph boundaries carry no visible token and preserve positions after them', () => {
  const { doc, text } = boundaryText();
  try {
    expect(text.toDelta()).toEqual([{ insert: { p: 1 } }]);
    expect(tokensOfText(text)).toEqual([]);
    text.insert(1, 'x');
    expect(tokensOfText(text).map((token) => token.position)).toEqual([1]);
    text.delete(1, 1);
    expect(text.toDelta()).toEqual([{ insert: { p: 1 } }]);
  } finally {
    doc.destroy();
  }
});
test('unobserved prefixes and text use opposite sides of an empty paragraph boundary', () => {
  const { doc, text } = boundaryText();
  try {
    expect(paragraphInsertionPosition(text, [], [plain], 0)).toBe(1);
    expect(paragraphInsertionPosition(text, [], [control], 0)).toBe(0);
    text.insert(1, 'hidden relocated text');
    expect(paragraphInsertionPosition(text, [], [plain], 0)).toBe(1);
  } finally {
    doc.destroy();
  }
});
test('typing before an observed control retains its explicit position', () => {
  const { doc, text } = boundaryText();
  try {
    text.insert(0, 'x', control.attributes);
    const before = tokensOfText(text);
    expect(paragraphInsertionPosition(text, before, [plain], 0)).toBe(0);
    expect(paragraphInsertionPosition(text, before, [plain], 2)).toBe(2);
  } finally {
    doc.destroy();
  }
});
test('only the exact boundary shape is hidden from token diffs', () => {
  for (const value of [null, [], { p: 2 }, { p: 1, n: 'node' }, { p: 1, a: 1 }])
    expect(isParagraphBoundary(value)).toBe(false);
  expect(isParagraphBoundary({ p: 1 })).toBe(true);
  const { doc, text } = boundaryText();
  try {
    text.insertEmbed(1, { p: 1, n: 'node' });
    expect(tokensOfText(text).map((token) => token.key)).toEqual(['enode']);
  } finally {
    doc.destroy();
  }
});
