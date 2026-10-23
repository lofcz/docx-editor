import { expect, test } from 'bun:test';
import {
  createMathPayload,
  decodeMathPayload,
  mathWordFallback,
  withMathSource,
} from './math-codec';
import { encodeBlockPayload } from './payload-schema';

test('inline and multiline LaTeX retain exact whitespace, line endings, escapes, and Unicode', () => {
  for (const kind of ['inline', 'block'] as const) {
    for (const source of [
      '',
      '  x^2 + α  ',
      '\\frac{a}{b} & <x> "quoted"',
      '\\begin{aligned}\r\n x &= y \\\\\r z &= 𝛼\n\\end{aligned}\n',
    ]) {
      const payload = createMathPayload('math-1', kind, source);
      const read = decodeMathPayload(encodeBlockPayload(payload));
      expect(read).toEqual({ status: 'supported', kind, source, payload });
    }
  }
});

test('source updates preserve private fields and nested future data without mutating input', () => {
  const original = {
    ...createMathPayload('math-1', 'inline', 'old'),
    privateField: { v: 2 },
    attributes: { texExpression: 'old', vendor: ['retained', 7] },
  };
  const next = withMathSource(original, '  new\r\nsource  ');
  expect(next).toEqual({
    ...original,
    attributes: { ...original.attributes, texExpression: '  new\r\nsource  ' },
  });
  expect(original.attributes.texExpression).toBe('old');
});

test('Word fallback is readable while the authoritative source stays exact', () => {
  const payload = createMathPayload('math-1', 'block', '\\frac{x}{y}\r\n  α + 𝛼\r\n');
  expect(mathWordFallback(payload)).toEqual({
    label: '\\frac{x}{y}',
    inlineText: '\\frac{x}{y}   α + 𝛼 ',
    paragraphs: ['\\frac{x}{y}', '  α + 𝛼', ''],
  });
  expect(payload.attributes.texExpression).toBe('\\frac{x}{y}\r\n  α + 𝛼\r\n');
  expect(mathWordFallback(createMathPayload('empty', 'inline', '')).inlineText).toBe('□');
  expect(mathWordFallback(createMathPayload('unicode', 'block', '𝛼'.repeat(130))).label).toBe(
    '𝛼'.repeat(128)
  );
});

test('unknown versions remain opaque; malformed sources and unrelated payloads refuse', () => {
  const future = '{"version":9,"futureMath":true}';
  expect(decodeMathPayload(future)).toMatchObject({ status: 'unsupported-version', raw: future });
  for (const attributes of [
    {},
    { texExpression: 12 },
    { texExpression: '\u0000' },
    { texExpression: '\ud800' },
  ]) {
    expect(
      decodeMathPayload(JSON.stringify({ id: 'm', type: 'equation', version: 1, attributes }))
        .status
    ).toBe('invalid');
  }
  expect(
    decodeMathPayload(
      encodeBlockPayload({ id: 'm', type: 'other', version: 1, attributes: { texExpression: 'x' } })
    ).status
  ).toBe('invalid');
  expect(() => createMathPayload('bad id', 'inline', 'x')).toThrow();
  expect(() => createMathPayload('valid', 'inline', '\udfff')).toThrow();
});
