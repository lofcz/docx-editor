import { expect, test } from 'bun:test';
import { registerUndoHistoryPosition, undoHistoryPosition } from '../store/undo-history-position';
const symbol = Symbol.for('@docx-editor.dev/undo-history-authority');
test('transparent wrappers reuse native authority through frozen and nested owners', () => {
  const native = Object.freeze({}),
    key = {},
    position = { current: () => key, split: () => {} };
  registerUndoHistoryPosition(native, position);
  const first = Object.freeze({ [symbol]: native }),
    second = Object.freeze({ [symbol]: first });
  expect(undoHistoryPosition(second)).toBe(position);
  expect(undoHistoryPosition(first)!.current()).toBe(key);
  expect(undoHistoryPosition(native)).toBe(position);
});
test('invalid, cyclic and throwing wrapper capabilities cannot hang a local history lookup', () => {
  const a: { [symbol]?: object } = {},
    b: { [symbol]?: object } = {};
  a[symbol] = b;
  b[symbol] = a;
  expect(undoHistoryPosition(a)).toBeUndefined();
  expect(undoHistoryPosition({ [symbol]: 3 })).toBeUndefined();
  expect(
    undoHistoryPosition(
      new Proxy(
        {},
        {
          get() {
            throw new Error('host failure');
          },
        }
      )
    )
  ).toBeUndefined();
});
