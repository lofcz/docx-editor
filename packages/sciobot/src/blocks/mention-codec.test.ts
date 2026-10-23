import { expect, test } from 'bun:test';
import { createMentionPayload, decodeMentionPayload, mentionWordFallback } from './mention-codec';
import { encodeBlockPayload } from './payload-schema';
import { serializeMention } from '../projection/mention';
import { parseTargetMarkdown } from '../projection/parse-target';

test('person/document keys, display strings, and private fields roundtrip independently', () => {
  const attributes = { icon: '📕', coverImage: 'asset:cover', private: { future: [true, 7] } };
  for (const key of ['user-7', '/docs/lesson-1', 'literal:%2F (#é𝛼)']) {
    for (const value of [
      'Ada Lovelace',
      '',
      ' \tMath\r\nnotes\n ',
      '@ada and @bob',
      '𝛼 [x] <entity> {danger()} &copy; \\ * _ ` $ ! | ~',
    ]) {
      const payload = createMentionPayload('block-id', key, value, attributes);
      expect(decodeMentionPayload(encodeBlockPayload(payload))).toEqual({
        status: 'supported',
        key,
        value,
        payload,
      });
      expect(payload.attributes.private).toEqual(attributes.private);
      const serialized = serializeMention(payload);
      const parsed = parseTargetMarkdown(serialized.text);
      if (!parsed.success) throw Error(parsed.error);
      expect(parsed.target.normalizedAst.children).toHaveLength(1);
      const paragraph = parsed.target.normalizedAst.children![0]!;
      if (!Array.isArray(paragraph.children)) throw Error('Expected mention children');
      expect(paragraph.children).toHaveLength(1);
      expect(paragraph.children![0]).toMatchObject({
        type: 'mention',
        username: key,
        displayText: value,
        children: [{ type: 'text', value }],
      });
      const rawParagraph = parsed.target.ast.children[0]!;
      if (!Array.isArray(rawParagraph.children)) throw Error('Expected source mention children');
      expect(rawParagraph.children[0]).toMatchObject({
        position: { start: { offset: 0 }, end: { offset: serialized.text.length } },
      });
      expect(
        serialized.properties.map((property) => ({
          path: property.path,
          valueStart: property.valueStart,
          valueEnd: property.valueEnd,
        }))
      ).toEqual([
        { path: ['key'], valueStart: null, valueEnd: null },
        { path: ['value'], valueStart: null, valueEnd: null },
      ]);
    }
  }
  expect(attributes).toEqual({
    icon: '📕',
    coverImage: 'asset:cover',
    private: { future: [true, 7] },
  });
});

test('Word labels normalize display line boundaries without changing stored source or Unicode', () => {
  const payload = createMentionPayload('mention', '/docs/empty', '  A\r\nB\t𝛼\n');
  expect(mentionWordFallback(payload)).toEqual({ text: '  A B 𝛼 ', label: '  A B 𝛼 ' });
  expect(payload.attributes.value).toBe('  A\r\nB\t𝛼\n');
  expect(mentionWordFallback(createMentionPayload('empty', '/docs/empty', ''))).toEqual({
    text: '/docs/empty',
    label: '/docs/empty',
  });
  expect(mentionWordFallback(createMentionPayload('long', 'user', '𝛼'.repeat(130))).label).toBe(
    '𝛼'.repeat(128)
  );
});

test('unknown versions stay opaque; invalid sources and executable attributes refuse', () => {
  const raw = '{"version":9,"futureMention":true}';
  expect(decodeMentionPayload(raw)).toMatchObject({ status: 'unsupported-version', raw });
  for (const [key, value] of [
    ['', 'Name'],
    ['person', '\u0000'],
    ['\ud800', 'Name'],
  ])
    expect(() => createMentionPayload('m', key!, value!)).toThrow();
  expect(
    decodeMentionPayload(
      encodeBlockPayload({
        id: 'm',
        type: 'other',
        version: 1,
        attributes: { key: 'u', value: 'Name' },
      })
    ).status
  ).toBe('invalid');
  let called = false;
  expect(() =>
    createMentionPayload('m', 'u', 'Name', {
      get private() {
        called = true;
        return 'executed';
      },
    })
  ).toThrow();
  expect(called).toBe(false);
});
