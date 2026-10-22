import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { attachOwnedTextAttributes } from './owned-text';
import { OWNED_TEXT_PREFIX, nativeAttributes } from './ownership-codec';

function fixture(clients: readonly [number, number]) {
  const docs = [new Y.Doc(), new Y.Doc()];
  docs.forEach((doc, index) => (doc.clientID = clients[index]!));
  const ownership = docs.map((doc) => attachOwnedTextAttributes(doc));
  const left = docs[0]!.getText('paragraph');
  left.insert(0, 'Two', { shell: 'original' });
  Y.applyUpdate(docs[1]!, Y.encodeStateAsUpdate(docs[0]!), 'relay');
  return {
    docs,
    left,
    right: docs[1]!.getText('paragraph'),
    ownership,
    destroy() {
      ownership.forEach((item) => item.dispose());
      docs.forEach((doc) => doc.destroy());
    },
  };
}
const read = (text: Y.Text) =>
  (text.toDelta() as { insert: string; attributes?: Record<string, unknown> }[]).map((entry) => ({
    insert: entry.insert,
    attributes: nativeAttributes(entry.attributes ?? {}),
  }));
function sync(docs: Y.Doc[], first: 0 | 1) {
  const other = 1 - first;
  Y.applyUpdate(docs[first]!, Y.encodeStateAsUpdate(docs[other]!), 'relay');
  Y.applyUpdate(docs[other]!, Y.encodeStateAsUpdate(docs[first]!), 'relay');
}
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    test(`owned formatting excludes unseen suffix and survives history; clients ${clients.join('/')}, receiver ${first}`, () => {
      const pair = fixture(clients),
        undo = new Y.UndoManager(pair.left, { trackedOrigins: new Set(['alice']) });
      try {
        pair.docs[0]!.transact(
          () => pair.left.format(0, 3, { deletion: 'alice', bold: true }),
          'alice'
        );
        pair.right.insert(3, ' Bob', { shell: 'original' });
        sync(pair.docs, first);
        const pending = [
          { insert: 'Two', attributes: { shell: 'original', deletion: 'alice', bold: true } },
          { insert: ' Bob', attributes: { shell: 'original' } },
        ];
        expect(read(pair.left)).toEqual(pending);
        expect(read(pair.right)).toEqual(pending);
        undo.undo();
        sync(pair.docs, first);
        expect(
          read(pair.left)
            .map((item) => item.insert)
            .join('')
        ).toBe('Two Bob');
        expect(
          read(pair.left).every((item) => !item.attributes.deletion && !item.attributes.bold)
        ).toBe(true);
        undo.redo();
        sync(pair.docs, first);
        expect(read(pair.left)).toEqual(pending);
        expect(read(pair.right)).toEqual(pending);
        pair.ownership.forEach((item) => item.assertHealthy());
      } finally {
        undo.destroy();
        pair.destroy();
      }
    });
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    test(`a leading tracked replacement and peer suffix retain text and attributes through redo; clients ${clients.join('/')}, receiver ${first}`, () => {
      const pair = fixture(clients),
        undo = new Y.UndoManager(pair.left, { trackedOrigins: new Set(['alice']) });
      try {
        pair.docs[0]!.transact(() => {
          pair.left.insert(0, 'Changed', { shell: 'replacement', insertion: 'alice', bold: true });
          pair.left.format(7, 3, { shell: 'deleted', deletion: 'alice', bold: true });
        }, 'alice');
        pair.right.insert(3, ' Bob', { shell: 'original' });
        sync(pair.docs, first);
        const proposed = () =>
          read(pair.left)
            .filter((span) => !span.attributes.deletion)
            .map((span) => span.insert)
            .join('');
        expect(proposed()).toBe('Changed Bob');
        undo.undo();
        sync(pair.docs, first);
        expect(
          read(pair.left)
            .map((span) => span.insert)
            .join('')
        ).toBe('Two Bob');
        undo.redo();
        sync(pair.docs, first);
        expect(proposed()).toBe('Changed Bob');
        expect(read(pair.left)).toEqual(read(pair.right));
        pair.ownership.forEach((item) => item.assertHealthy());
      } finally {
        undo.destroy();
        pair.destroy();
      }
    });

test('repair groups UTF-16 changes without splitting a supplementary character', () => {
  const pair = fixture([73, 41]);
  try {
    pair.left.format(0, 3, { deletion: 'alice' });
    pair.right.insert(3, ' 😀', { shell: 'original' });
    sync(pair.docs, 0);
    expect(pair.left.toString()).toBe('Two 😀');
    expect(
      read(pair.left)
        .filter((span) => !span.attributes.deletion)
        .map((span) => span.insert)
        .join('')
    ).toBe(' 😀');
    expect(read(pair.left)).toEqual(read(pair.right));
  } finally {
    pair.destroy();
  }
});

test('unregistered and disposed documents retain ordinary Yjs behavior', () => {
  const doc = new Y.Doc(),
    text = doc.getText('plain');
  text.insert(0, 'Before');
  expect(Object.keys(text.toDelta()[0]!.attributes ?? {})).toEqual([]);
  const registered = new Y.Doc();
  const handle = attachOwnedTextAttributes(registered);
  handle.dispose();
  handle.dispose();
  registered.getText('plain').insert(0, 'After disposal');
  expect(registered.getMap('sciobot-owned-text-operations-v1').size).toBe(0);
  text.insert(6, 'After');
  expect(text.toString()).toBe('BeforeAfter');
  expect(
    Object.keys(text.toDelta()[0]!.attributes ?? {}).some((key) =>
      key.startsWith(OWNED_TEXT_PREFIX)
    )
  ).toBe(false);
  doc.destroy();
  registered.destroy();
});
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    test(`a peer repair delivered before its intents does not loop; clients ${clients.join('/')}, receiver ${first}`, () => {
      const pair = fixture(clients);
      let connected = false,
        relayed = 0;
      pair.docs.forEach((doc, index) =>
        doc.on('update', (update: Uint8Array, origin: unknown) => {
          if (!connected || origin === 'relay') return;
          if (++relayed > 64) throw new Error('Owned repairs did not converge.');
          Y.applyUpdate(pair.docs[1 - index]!, update, 'relay');
        })
      );
      try {
        pair.left.format(0, 3, { deletion: 'alice', run: 'alice' });
        pair.right.insert(3, ' Bob', { shell: 'original' });
        pair.right.format(0, 7, { run: 'bob' });
        pair.right.format(4, 3, { italic: true });
        connected = true;
        // The receiver's repair relays at once, ahead of the other peer's offline intents.
        sync(pair.docs, first);
        const expected = [
          { insert: 'Two', attributes: { shell: 'original', deletion: 'alice', run: 'alice' } },
          { insert: ' ', attributes: { shell: 'original', run: 'bob' } },
          { insert: 'Bob', attributes: { shell: 'original', run: 'bob', italic: true } },
        ];
        const [left, right] = [read(pair.left), read(pair.right)];
        expect(left).toEqual(right);
        expect(left.map((item) => item.insert).join('')).toBe('Two Bob');
        expect(left.at(-1)).toEqual(expected.at(-1));
        expect(pair.docs[0]!.store.pendingStructs).toBeNull();
        expect(pair.docs[1]!.store.pendingStructs).toBeNull();
        pair.ownership.forEach((item) => item.assertHealthy());
      } finally {
        connected = false;
        pair.destroy();
      }
    });
