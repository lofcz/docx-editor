import { expect, test } from 'bun:test';
import type { BlockPayload } from '../blocks/payload-schema';
import { projectedCodeMetadata } from './code-review-payload';

const before: BlockPayload = {
  id: 'code',
  type: 'code_block',
  version: 2,
  attributes: { value: 'original\n\tsource', language: 'python', opaque: { imported: true } },
  extension: { retained: 'opaque' },
};

test('source-only code review retains peer language and never writes a source cache', () => {
  const after = { ...before, attributes: { ...before.attributes, value: 'proposed\nsource' } };
  const current: BlockPayload = {
    ...before,
    attributes: {
      ...before.attributes,
      value: 'original\n\tsource plus peer',
      language: 'typescript',
    },
  };
  for (const view of ['committed', 'working'] as const)
    expect(projectedCodeMetadata(current, [before, after], view)).toBe(current);
});

test('code review changes declared metadata while preserving source and opaque peer fields', () => {
  const after = {
    ...before,
    attributes: { ...before.attributes, value: 'proposed', language: 'javascript' },
  };
  const current: BlockPayload = {
    ...before,
    attributes: {
      ...before.attributes,
      value: 'native peer source',
      opaque: { peer: 'preserved' },
    },
    extension: { peer: 'retained' },
  };
  expect(projectedCodeMetadata(current, [before, after], 'working')).toEqual({
    ...current,
    attributes: { ...current.attributes, language: 'javascript' },
  });
  expect(projectedCodeMetadata(current, [before, after], 'committed')).toEqual(current);
  expect(() =>
    projectedCodeMetadata(
      { ...current, attributes: { ...current.attributes, language: 'ruby' } },
      [before, after],
      'working'
    )
  ).toThrow('outside its retained history');
  expect(current.attributes.language).toBe('python');
  expect(current.attributes.value).toBe('native peer source');
});
