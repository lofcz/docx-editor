import { expect, test } from 'bun:test';
import { parseTargetMarkdown } from './parse-target';
import { serializeCodeBlock } from './code-block';

test('code block serialization preserves language, blank lines, and backtick runs', () => {
  const payload = {
    version: 1 as const,
    id: 'code-1',
    type: 'code_block',
    attributes: { language: 'ts', value: 'const tick = ````;\n\nreturn tick;' },
  };
  const serialized = serializeCodeBlock(payload);
  expect(serialized.text).toBe('`````ts\nconst tick = ````;\n\nreturn tick;\n`````');
  const parsed = parseTargetMarkdown(serialized.text);
  expect(parsed.success).toBe(true);
  if (!parsed.success) return;
  expect(parsed.target.ast.children[0]).toMatchObject({
    type: 'code',
    lang: 'ts',
    value: payload.attributes.value,
  });
});
