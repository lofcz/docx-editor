import { expect, test } from 'bun:test';
import reference from '../__tests__/fixtures/find-replace.json';
import { findAndReplace } from './find-replace';

for (const [index, item] of reference.cases.entries()) {
  test(`${item.sourceTest} [${index + 1}]`, () => {
    const { content, oldString, newString, replaceAll, expectedReplacements } = item.input;
    expect<unknown>(
      findAndReplace(content, oldString, newString, replaceAll, expectedReplacements)
    ).toEqual(item.expected);
  });
}

test('failed matching never returns a partial replacement', () => {
  const source = 'keep x and x';
  expect(findAndReplace(source, 'x', 'changed', false)).toMatchObject({ success: false });
  expect(findAndReplace(source, 'x', 'changed', true, 1)).toMatchObject({ success: false });
  expect(source).toBe('keep x and x');
});
