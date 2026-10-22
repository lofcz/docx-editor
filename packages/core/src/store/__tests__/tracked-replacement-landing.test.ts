import { expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint, findNode, trackedReplacementLanding } from '../index.ts';
import { ADA, paragraphId, part } from './tracked-edit-fixture.ts';

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const pending = (kind: 'ins' | 'del', author: string, text: string, date = ADA.date) =>
  `<w:${kind} w:id="1" w:author="${author}" w:date="${date}"><w:r><w:${kind === 'del' ? 'delText' : 't'}>${text}</w:${kind === 'del' ? 'delText' : 't'}></w:r></w:${kind}>`;

test.each([
  { name: 'plain text remains struck', body: run('abcd'), start: 1, end: 3, expected: 3 },
  {
    name: 'own pending insertion retracts',
    body: run('a') + pending('ins', ADA.author, 'bc') + run('d'),
    start: 1,
    end: 3,
    expected: 1,
  },
  {
    name: 'foreign pending insertion remains',
    body: run('a') + pending('ins', 'Another author', 'bc') + run('d'),
    start: 1,
    end: 3,
    expected: 3,
  },
  {
    name: 'an adjacent current strike joins',
    body: run('a') + pending('del', ADA.author, 'bc') + run('d'),
    start: 0,
    end: 1,
    expected: 3,
  },
  {
    name: 'an old strike remains independent',
    body: run('a') + pending('del', ADA.author, 'bc', '2020-01-01T00:00:00Z') + run('d'),
    start: 0,
    end: 1,
    expected: 1,
  },
  {
    name: 'insertion inside a strike lands past it',
    body: run('a') + pending('del', ADA.author, 'bc') + run('d'),
    start: 2,
    end: 2,
    expected: 3,
  },
])(
  'the public canonical landing helper preserves native semantics: $name',
  ({ body, start, end, expected }) => {
    const source = part(`<w:p>${body}</w:p>`);
    const paragraph = findNode(source, paragraphId(source));
    if (paragraph?.kind !== 'paragraph') throw new Error('Missing paragraph');
    const before = canonicalOoxmlFingerprint(source);
    expect(trackedReplacementLanding(paragraph, start, end, ADA.author, ADA.date)).toBe(expected);
    expect(canonicalOoxmlFingerprint(source)).toBe(before);
  }
);
