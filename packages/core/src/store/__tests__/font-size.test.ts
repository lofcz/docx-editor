import { expect, test } from 'bun:test';
import { fontSizePoints, fontSizeAttribute } from '../package/font-size.ts';
test('Word integer half-points and exact positive universal measures share one read', () => {
  for (const [raw, expected] of [
    ['22', 11],
    ['6.75pt', 6.75],
    ['0.5in', 36],
    ['2.54cm', 72],
    ['25.4mm', 72],
    ['1pc', 12],
    ['1pi', 12],
  ] as const)
    expect(fontSizePoints(raw)).toBeCloseTo(expected, 12);
  for (const raw of ['13.5', '-2pt', '0', '2em', '1e2pt', 'bad'])
    expect(fontSizePoints(raw)).toBeNull();
  expect(fontSizeAttribute(6.75)).toBe('6.75pt');
  expect(fontSizeAttribute(11)).toBe('22');
});
