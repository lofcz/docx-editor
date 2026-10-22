import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  contentControlContentNodeOf,
  contentControlPropertiesOf,
  contentControlsIn,
  findNode,
  moveNode,
  parentNodeOf,
  readOoxmlPart,
  serializeOoxmlPart,
  type OoxmlElement,
  type OoxmlPart,
} from '../index.ts';
import {
  observeCanonicalPrimitiveJournal,
  packageTransactionPublished,
  runObservedStoreTransaction,
} from '../package/canonical-primitive-capture.ts';
import type { CanonicalPrimitiveEffect } from '../package/canonical-primitive-journal.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const control = (name: string, text: string, extra = '') =>
  `<w:sdt ${extra}><w:sdtPr><w:alias w:val="${name}"/></w:sdtPr><w:sdtContent>${text}</w:sdtContent></w:sdt>`;
function load(body: string): OoxmlPart {
  const read = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body>${body}</w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  return read.part;
}
function named(part: OoxmlPart, name: string): OoxmlElement {
  return contentControlsIn(part.root).find(
    ({ node }) => contentControlPropertiesOf(node).alias === name
  )!.node;
}
const content = (part: OoxmlPart, name: string) => contentControlContentNodeOf(named(part, name))!;
const fixture = () =>
  load(
    control('source', p('Before') + control('move', p('Text')) + p('After')) +
      control('target', p('Target'))
  );

test('native move preserves all identities, reuses descendants, and journals exactly one move', () => {
  const part = fixture(),
    source = named(part, 'move'),
    target = content(part, 'target');
  const original = canonicalOoxmlFingerprint(part),
    host = {},
    effects: CanonicalPrimitiveEffect[] = [];
  const stop = observeCanonicalPrimitiveJournal(host, (journal) =>
    effects.push(...journal.effects)
  );
  const result = runObservedStoreTransaction(
    host,
    () => {
      const moved = moveNode(part, source.id, target.id, 1);
      return moved.ok
        ? { ok: true as const, change: moved.part }
        : { ok: false as const, change: null };
    },
    packageTransactionPublished
  );
  stop();
  if (!result.ok) throw new Error('Native move failed');
  expect(findNode(result.change, source.id)).toBe(source);
  expect(parentNodeOf(result.change, source.id)!.id).toBe(target.id);
  expect(effects).toEqual([
    {
      kind: 'moveNode',
      logicalId: source.id,
      destinationParentLogicalId: target.id,
      destinationIndex: 1,
    },
  ]);
  expect(canonicalOoxmlFingerprint(part)).toBe(original);
  expect(canonicalOoxmlFingerprint(result.change)).toBe(
    canonicalOoxmlFingerprint(
      load(
        control('source', p('Before') + p('After')) +
          control('target', p('Target') + control('move', p('Text')))
      )
    )
  );
});

test('same-parent ordering uses post-removal indices and exact no-ops preserve the part', () => {
  const part = fixture(),
    node = named(part, 'move'),
    parent = content(part, 'source');
  expect(moveNode(part, node.id, parent.id, 1)).toEqual({ ok: true, part });
  const moved = moveNode(part, node.id, parent.id, 2);
  if (!moved.ok) throw new Error('Reorder failed');
  expect(content(moved.part, 'source').children.map((child) => child.id)).toEqual([
    parent.children[0]!.id,
    parent.children[2]!.id,
    node.id,
  ]);
});

test('cycles, root moves, missing nodes, invalid indices, and invalid Word structure refuse atomically', () => {
  const part = fixture(),
    node = named(part, 'move'),
    parent = content(part, 'source'),
    original = canonicalOoxmlFingerprint(part);
  for (const [id, target, index] of [
    [node.id, node.id, 0],
    [named(part, 'source').id, content(part, 'move').id, 0],
    [part.root.id, parent.id, 0],
    ['missing', parent.id, 0],
    [node.id, 'missing', 0],
    [node.id, parent.id, -1],
    [node.id, parent.id, 9],
    [node.id, parent.id, 0.5],
    [node.id, parent.children[0]!.id, 0],
  ] as const) {
    expect(moveNode(part, id, target, index).ok).toBe(false);
    expect(canonicalOoxmlFingerprint(part)).toBe(original);
  }
});

test('opaque QName values retain inherited namespace meaning across conflicts and serialization', () => {
  const part = load(
    control('source', control('move', p('Text'), 'mc:Ignorable="q"'), 'xmlns:q="urn:source"') +
      control('target', p('Target'), 'xmlns:q="urn:target"')
  );
  const node = named(part, 'move'),
    target = content(part, 'target');
  const moved = moveNode(part, node.id, target.id, 1);
  if (!moved.ok) throw new Error('Namespace move failed');
  expect(named(moved.part, 'move').namespaceBindings).toContainEqual({
    prefix: 'q',
    namespaceUri: 'urn:source',
  });
  const xml = serializeOoxmlPart(moved.part);
  const reopened = readOoxmlPart(xml, { name: part.name, contentType: part.contentType });
  if (!reopened.ok) throw new Error(reopened.reason);
  expect(canonicalOoxmlFingerprint(reopened.part)).toBe(canonicalOoxmlFingerprint(moved.part));
  expect(named(reopened.part, 'move').namespaceBindings).toContainEqual({
    prefix: 'q',
    namespaceUri: 'urn:source',
  });
});
