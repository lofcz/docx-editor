import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { attachOwnedTextAttributes, OwnedTextError } from './owned-text';
import {
  decodeIntent,
  decodeLineage,
  detachedAttributes,
  nativeAttributes,
  OWNED_TEXT_OPERATIONS,
} from './ownership-codec';

const read = (text: Y.Text) =>
  (text.toDelta() as { insert: unknown; attributes?: Record<string, unknown> }[]).map((entry) => ({
    insert: entry.insert,
    attributes: nativeAttributes(entry.attributes ?? {}),
  }));
const relay = (docs: Y.Doc[]) => {
  Y.applyUpdate(docs[1]!, Y.encodeStateAsUpdate(docs[0]!));
  Y.applyUpdate(docs[0]!, Y.encodeStateAsUpdate(docs[1]!));
};

test('partial formatting, field removal, tabs, and embeds retain their insertion baselines', () => {
  const docs = [new Y.Doc(), new Y.Doc()];
  docs[0]!.clientID = 73;
  docs[1]!.clientID = 41;
  const handles = docs.map((doc) => attachOwnedTextAttributes(doc));
  const left = docs[0]!.getText('p'),
    right = docs[1]!.getText('p');
  try {
    left.insert(0, 'Two\t', { color: 'red', italic: true });
    relay(docs);
    left.format(1, 3, { italic: null, bold: true });
    right.insertEmbed(4, { type: 'image', id: 'peer' }, { italic: true });
    relay(docs);
    const expected = [
      { insert: 'T', attributes: { color: 'red', italic: true } },
      { insert: 'wo\t', attributes: { color: 'red', bold: true } },
      { insert: { type: 'image', id: 'peer' }, attributes: { italic: true } },
    ];
    expect(read(left)).toEqual(expected);
    expect(read(right)).toEqual(expected);
    handles.forEach((handle) => handle.assertHealthy());
  } finally {
    handles.forEach((handle) => handle.dispose());
    docs.forEach((doc) => doc.destroy());
  }
});

test('concurrent fields combine and the same field uses deterministic causal order', () => {
  const docs = [new Y.Doc(), new Y.Doc()];
  docs[0]!.clientID = 73;
  docs[1]!.clientID = 41;
  const handles = docs.map((doc) => attachOwnedTextAttributes(doc));
  const left = docs[0]!.getText('p'),
    right = docs[1]!.getText('p');
  try {
    left.insert(0, 'Two', {});
    relay(docs);
    left.format(0, 3, { bold: true, color: 'red' });
    right.format(0, 3, { italic: true, color: 'blue' });
    relay(docs);
    const expected = [{ insert: 'Two', attributes: { bold: true, italic: true, color: 'red' } }];
    expect(read(left)).toEqual(expected);
    expect(read(right)).toEqual(expected);
    right.format(0, 3, { color: 'green' });
    relay(docs);
    expect(read(left)[0]!.attributes).toEqual({ bold: true, italic: true, color: 'green' });
  } finally {
    handles.forEach((handle) => handle.dispose());
    docs.forEach((doc) => doc.destroy());
  }
});

test('repeated restored text and parent lineage survives raw-room reopening', () => {
  const docs = [new Y.Doc(), new Y.Doc()];
  const handles = docs.map((doc) => attachOwnedTextAttributes(doc));
  const root = docs[0]!.getMap<Y.Text>('paragraphs');
  const undo = new Y.UndoManager(root, { trackedOrigins: new Set(['author']) });
  let reopened: Y.Doc | undefined;
  let reopenedHandle: ReturnType<typeof attachOwnedTextAttributes> | undefined;
  try {
    docs[0]!.transact(() => root.set('p', new Y.Text('Two')), 'author');
    relay(docs);
    undo.undo();
    relay(docs);
    expect(docs[1]!.getMap('paragraphs').has('p')).toBe(false);
    undo.redo();
    relay(docs);
    root.get('p')!.format(0, 3, { bold: true });
    relay(docs);
    undo.stopCapturing();
    docs[0]!.transact(() => root.delete('p'), 'author');
    relay(docs);
    for (let i = 0; i < 3; i++) {
      undo.undo();
      relay(docs);
      expect(read(docs[1]!.getMap<Y.Text>('paragraphs').get('p')!)).toEqual([
        { insert: 'Two', attributes: { bold: true } },
      ]);
      if (i < 2) {
        undo.redo();
        relay(docs);
      }
    }
    reopened = new Y.Doc();
    reopenedHandle = attachOwnedTextAttributes(reopened);
    Y.applyUpdate(reopened, Y.encodeStateAsUpdate(docs[0]!));
    const restored = reopened.getMap<Y.Text>('paragraphs').get('p')!;
    expect(read(restored)).toEqual([{ insert: 'Two', attributes: { bold: true } }]);
    restored.insert(3, ' Peer', { italic: true });
    Y.applyUpdate(docs[0]!, Y.encodeStateAsUpdate(reopened));
    relay(docs);
    expect(read(root.get('p')!)).toEqual([
      { insert: 'Two', attributes: { bold: true } },
      { insert: ' Peer', attributes: { italic: true } },
    ]);
    reopenedHandle.assertHealthy();
    handles.forEach((handle) => handle.assertHealthy());
  } finally {
    reopenedHandle?.dispose();
    reopened?.destroy();
    undo.destroy();
    handles.forEach((handle) => handle.dispose());
    docs.forEach((doc) => doc.destroy());
  }
});

test('attachment rejects legacy rooms and unknown versions without changing their wire content', () => {
  const doc = new Y.Doc();
  try {
    doc.getText('legacy').insert(0, 'Legacy');
    const before = Y.encodeStateAsUpdate(doc);
    expect(() => attachOwnedTextAttributes(doc)).toThrow(OwnedTextError);
    expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
    doc.getMap(OWNED_TEXT_OPERATIONS).set('protocol', '2');
    expect(() => attachOwnedTextAttributes(doc)).toThrow('Unsupported');
  } finally {
    doc.destroy();
  }
});

test('invalid data never invokes getters or changes local text and history', () => {
  const doc = new Y.Doc(),
    handle = attachOwnedTextAttributes(doc),
    text = doc.getText('p');
  try {
    text.insert(0, 'Two', {});
    let reads = 0;
    const attributes = {
      get bold() {
        reads++;
        return true;
      },
    };
    const before = Y.encodeStateAsUpdate(doc);
    expect(() => text.insert(3, 'bad', attributes)).toThrow(OwnedTextError);
    expect(() => text.format(0, 3, JSON.parse('{"__proto__":{"polluted":true}}'))).toThrow(
      OwnedTextError
    );
    expect(() =>
      text.applyDelta([{ insert: 'bad' }, { retain: 100, attributes: { bold: true } }])
    ).toThrow(OwnedTextError);
    expect(() => text.applyDelta([{ insert: 'bad', attributes: { bold: undefined } }])).toThrow(
      OwnedTextError
    );
    expect(() => text.insert(3, 'bad', { large: 'x'.repeat(8 * 1024 * 1024) })).toThrow(
      OwnedTextError
    );
    expect(reads).toBe(0);
    expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
    handle.assertHealthy();
  } finally {
    handle.dispose();
    doc.destroy();
  }
});

test('bounded codecs reject malformed records, unsafe keys, overflowing ranges, and excess depth', () => {
  const valid = {
    kind: 'insert',
    stamp: [1, 41, 0],
    ranges: [[41, 0, 3]],
    values: { bold: true },
  } as const;
  expect(decodeIntent(JSON.stringify(valid))).toEqual(valid);
  for (const value of [
    '',
    '{}',
    JSON.stringify({ ...valid, ranges: [[41, Number.MAX_SAFE_INTEGER, 1]] }),
    JSON.stringify({ ...valid, extra: true }),
    JSON.stringify({ ...valid, values: JSON.parse('{"constructor":{}}') }),
  ])
    expect(() => decodeIntent(value)).toThrow(OwnedTextError);
  expect(() => decodeLineage('{"source":[1,9007199254740991],"length":2}')).toThrow(OwnedTextError);
  expect(() =>
    decodeIntent(
      JSON.stringify({ ...valid, ranges: Array.from({ length: 100_001 }, () => [41, 0, 1]) })
    )
  ).toThrow(OwnedTextError);
  let nested: unknown = {};
  for (let i = 0; i < 35; i++) nested = { nested };
  expect(() => detachedAttributes(nested)).toThrow(OwnedTextError);
});

test('malformed remote records fail before text observers and keep the room failed', () => {
  const source = new Y.Doc(),
    target = new Y.Doc();
  const author = attachOwnedTextAttributes(source);
  const failures: Error[] = [];
  const receiver = attachOwnedTextAttributes(target, { onError: (error) => failures.push(error) });
  let observed = 0;
  try {
    source.getText('p').insert(0, 'Two', {});
    Y.applyUpdate(target, Y.encodeStateAsUpdate(source));
    target.getText('p').observe(() => observed++);
    // A legacy writer can send a malformed record. The registered receiver must stop.
    author.dispose();
    source.transact(() => {
      source.getMap(OWNED_TEXT_OPERATIONS).set('broken', '{}');
      source.getText('p').format(0, 3, { bold: true });
    });
    expect(() => Y.applyUpdate(target, Y.encodeStateAsUpdate(source))).toThrow(OwnedTextError);
    expect(observed).toBe(0);
    expect(failures).toHaveLength(1);
    expect(() => receiver.assertHealthy()).toThrow(OwnedTextError);
    expect(() => target.getText('p').insert(3, 'Later', {})).toThrow(OwnedTextError);
  } finally {
    author.dispose();
    receiver.dispose();
    source.destroy();
    target.destroy();
  }
});

test('delta authoring detaches caller data and retains native trailing-newline behavior', () => {
  const doc = new Y.Doc(),
    handle = attachOwnedTextAttributes(doc),
    text = doc.getText('p');
  try {
    const attributes = { payload: { value: 'original' } };
    text.applyDelta([{ insert: 'A\n', attributes }], { sanitize: false });
    attributes.payload.value = 'mutated';
    expect(text.toString()).toBe('A');
    expect(read(text)[0]!.attributes).toEqual({ payload: { value: 'original' } });
    text.applyDelta([{ retain: 1 }, { insert: '\n' }]);
    expect(text.toString()).toBe('A\n');
    handle.assertHealthy();
  } finally {
    handle.dispose();
    doc.destroy();
  }
});
