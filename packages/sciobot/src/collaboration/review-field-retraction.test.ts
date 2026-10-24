import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { attachOwnedReviewFields, type ReviewFieldReceipt } from './review-fields';

for (const nested of [false, true])
  for (const receiver of [0, 1] as const)
    for (const peerValue of ['peer', 'proposal', null] as const)
      test(`scalar proposal retraction preserves exact peer ownership and history: ${nested}/${receiver}/${peerValue}`, () => {
        const docs = [new Y.Doc(), new Y.Doc()];
        docs[0]!.clientID = 41;
        docs[1]!.clientID = 73;
        const fields = docs.map((doc) => attachOwnedReviewFields(doc, { protocol: '2' }));
        const root = docs[0]!.getMap<unknown>('native');
        if (nested) {
          const node = new Y.Map<string>();
          node.set('format', 'original');
          root.set('node', node);
        } else root.set('format', 'original');
        Y.applyUpdate(docs[1]!, Y.encodeStateAsUpdate(docs[0]!), 'relay');
        const mapOf = (doc: Y.Doc): Y.Map<string> =>
          nested ? doc.getMap<Y.Map<string>>('native').get('node')! : doc.getMap<string>('native');
        const maps = docs.map(mapOf);
        fields.forEach((field) => field.enable());
        const histories = maps.map((map) => new Y.UndoManager(map, { captureTimeout: 0 }));
        let connected = false;
        docs.forEach((doc, index) =>
          doc.on('update', (update, origin) => {
            if (connected && origin !== 'relay') Y.applyUpdate(docs[1 - index]!, update, 'relay');
          })
        );
        const check = (value: string | undefined) => {
          for (const map of maps) expect(map.get('format')).toBe(value);
        };
        let receipt: readonly ReviewFieldReceipt[] = [];
        try {
          fields[0]!.capture(
            () => docs[0]!.transact(() => maps[0]!.set('format', 'proposal')),
            1,
            (items) => {
              receipt = items;
            }
          );
          expect(receipt).toHaveLength(1);
          fields[1]!.capture(
            () =>
              docs[1]!.transact(() =>
                peerValue === null ? maps[1]!.delete('format') : maps[1]!.set('format', peerValue)
              ),
            1
          );
          connected = true;
          Y.applyUpdate(docs[receiver]!, Y.encodeStateAsUpdate(docs[1 - receiver]!), 'relay');
          Y.applyUpdate(docs[1 - receiver]!, Y.encodeStateAsUpdate(docs[receiver]!), 'relay');
          // The native rejection writes the old snapshot. The retraction must suppress
          // that write's ownership and withdraw only the exact proposal receipt.
          docs[0]!.transact(() =>
            fields[0]!.retractInTransaction(receipt, () => maps[0]!.set('format', 'original'))
          );
          check(peerValue ?? undefined);
          expect(histories[1]!.undo()).not.toBeNull();
          check('original');
          expect(histories[1]!.redo()).not.toBeNull();
          check(peerValue ?? undefined);
          expect(histories[0]!.undo()).not.toBeNull();
          expect(histories[1]!.undo()).not.toBeNull();
          check('proposal');
          expect(histories[0]!.redo()).not.toBeNull();
          check('original');
          expect(histories[1]!.redo()).not.toBeNull();
          check(peerValue ?? undefined);
          const cold = new Y.Doc(),
            scope = attachOwnedReviewFields(cold, { protocol: '2' });
          try {
            Y.applyUpdate(cold, Y.encodeStateAsUpdate(docs[0]!));
            scope.enable();
            expect(mapOf(cold).get('format')).toBe(peerValue ?? undefined);
          } finally {
            scope.dispose();
            cold.destroy();
          }
        } finally {
          histories.forEach((history) => history.destroy());
          fields.forEach((field) => field.dispose());
          docs.forEach((doc) => doc.destroy());
        }
      });

test('scalar retraction keeps a later edit by the proposal author', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFields(doc, { protocol: '2' });
  const map = doc.getMap<string>('native');
  map.set('format', 'original');
  scope.enable();
  const history = new Y.UndoManager(map, { captureTimeout: 0 });
  let receipt: readonly ReviewFieldReceipt[] = [];
  try {
    scope.capture(
      () => doc.transact(() => map.set('format', 'proposal')),
      1,
      (items) => {
        receipt = items;
      }
    );
    map.set('format', 'later');
    doc.transact(() => scope.retractInTransaction(receipt, () => map.set('format', 'original')));
    expect(map.get('format')).toBe('later');
    history.undo();
    expect(map.get('format')).toBe('later');
    history.undo();
    expect(map.get('format')).toBe('proposal');
    history.redo();
    history.redo();
    expect(map.get('format')).toBe('later');
  } finally {
    history.destroy();
    scope.dispose();
    doc.destroy();
  }
});

test('invalid scalar receipts refuse before any native or ownership write', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFields(doc, { protocol: '2' });
  const map = doc.getMap<string>('native');
  map.set('format', 'original');
  scope.enable();
  let receipt: readonly ReviewFieldReceipt[] = [];
  try {
    scope.capture(
      () => doc.transact(() => map.set('format', 'proposal')),
      1,
      (items) => {
        receipt = items;
      }
    );
    for (const invalid of [
      [receipt[0]!, receipt[0]!],
      [{ ...receipt[0]!, record: receipt[0]!.record + ' ' }],
      [{ ...receipt[0]!, id: '73:0' }],
      [{ ...receipt[0]!, extra: true }],
    ]) {
      const before = Y.encodeStateAsUpdate(doc);
      expect(() =>
        doc.transact(() => scope.retractInTransaction(invalid, () => map.set('format', 'bad')))
      ).toThrow('stale or invalid');
      expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
      expect(map.get('format')).toBe('proposal');
    }
    expect(() => scope.retractInTransaction(receipt, () => {})).toThrow('active local');
  } finally {
    scope.dispose();
    doc.destroy();
  }
});

test('protocol 1 refuses protocol 2 rooms and cannot author retractions', () => {
  const doc = new Y.Doc(),
    old = new Y.Doc();
  const scope = attachOwnedReviewFields(doc, { protocol: '2' });
  const legacy = attachOwnedReviewFields(old);
  try {
    doc.getMap('native').set('format', 'original');
    scope.enable();
    scope.capture(() => doc.transact(() => doc.getMap('native').set('format', 'proposal')), 1);
    const joined = new Y.Doc();
    try {
      Y.applyUpdate(joined, Y.encodeStateAsUpdate(doc));
      const before = Y.encodeStateAsUpdate(joined);
      expect(() => attachOwnedReviewFields(joined)).toThrow('Unsupported');
      expect(Y.encodeStateAsUpdate(joined)).toEqual(before);
    } finally {
      joined.destroy();
    }
    legacy.enable();
    const before = Y.encodeStateAsUpdate(old);
    expect(() => old.transact(() => legacy.retractInTransaction([], () => {}))).toThrow(
      'protocol 2'
    );
    expect(Y.encodeStateAsUpdate(old)).toEqual(before);
  } finally {
    scope.dispose();
    legacy.dispose();
    doc.destroy();
    old.destroy();
  }
});

test('retraction reserves its receipt and unrelated native writes before mutation', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFields(doc, {
      protocol: '2',
      limits: { targets: 1, intents: 2 },
    });
  const map = doc.getMap<string>('native');
  map.set('format', 'original');
  map.set('other', 'original');
  scope.enable();
  let receipt: readonly ReviewFieldReceipt[] = [];
  try {
    scope.capture(
      () => doc.transact(() => map.set('format', 'proposal')),
      1,
      (items) => {
        receipt = items;
      }
    );
    const before = Y.encodeStateAsUpdate(doc);
    expect(() =>
      doc.transact(() =>
        scope.retractInTransaction(
          receipt,
          () => {
            map.set('format', 'original');
            map.set('other', 'changed');
          },
          1
        )
      )
    ).toThrow('record limit');
    expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
    expect(map.get('format')).toBe('proposal');
    doc.transact(() => scope.retractInTransaction(receipt, () => map.set('format', 'original')));
    expect(map.get('format')).toBe('original');
    const cancelled = Y.encodeStateAsUpdate(doc);
    expect(() => doc.transact(() => scope.retractInTransaction(receipt, () => {}))).toThrow(
      'stale'
    );
    expect(Y.encodeStateAsUpdate(doc)).toEqual(cancelled);
  } finally {
    scope.dispose();
    doc.destroy();
  }
});

test('a third reviewer withdraws the proposal while both authors keep independent histories', () => {
  const docs = [new Y.Doc(), new Y.Doc(), new Y.Doc()];
  docs.forEach((doc, index) => {
    doc.clientID = 41 + index;
  });
  const fields = docs.map((doc) => attachOwnedReviewFields(doc, { protocol: '2' }));
  docs[0]!.getMap('native').set('format', 'original');
  for (const doc of docs.slice(1)) Y.applyUpdate(doc, Y.encodeStateAsUpdate(docs[0]!));
  const maps = docs.map((doc) => doc.getMap<string>('native'));
  fields.forEach((scope) => scope.enable());
  const histories = maps.map((map) => new Y.UndoManager(map, { captureTimeout: 0 }));
  const sync = () => {
    for (const source of docs)
      for (const target of docs)
        if (source !== target) Y.applyUpdate(target, Y.encodeStateAsUpdate(source), 'relay');
  };
  const check = (expected: string) => {
    sync();
    for (const map of maps) expect(map.get('format')).toBe(expected);
  };
  let receipt: readonly ReviewFieldReceipt[] = [];
  try {
    fields[0]!.capture(
      () => docs[0]!.transact(() => maps[0]!.set('format', 'proposal')),
      1,
      (items) => {
        receipt = items;
      }
    );
    maps[1]!.set('format', 'peer');
    sync();
    docs[2]!.transact(() =>
      fields[2]!.retractInTransaction(receipt, () => maps[2]!.set('format', 'original'))
    );
    check('peer');
    histories[1]!.undo();
    check('original');
    histories[0]!.undo();
    check('original');
    histories[0]!.redo();
    check('original');
    histories[2]!.undo();
    check('proposal');
    histories[2]!.redo();
    check('original');
    histories[1]!.redo();
    check('peer');
  } finally {
    histories.forEach((history) => history.destroy());
    fields.forEach((scope) => scope.dispose());
    docs.forEach((doc) => doc.destroy());
  }
});
