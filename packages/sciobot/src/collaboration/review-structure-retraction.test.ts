import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { attachOwnedReviewFields, type ReviewFieldReceipt } from './review-fields';
import { attachOwnedReviewFlags } from './review-flags';
import {
  type ReviewFlagReceipt,
  REVIEW_FLAG_INTENTS,
  REVIEW_FLAG_RETRACTIONS,
} from './review-flag-codec';

// A native property replacement owns both its old child's removal flag and
// the ordered child edges. Both ownership kinds must withdraw together.
for (const receiver of [0, 1] as const)
  for (const peerReplacement of [false, true])
    test(`property subtree retraction preserves peer replacement and both histories: ${receiver}/${peerReplacement}`, () => {
      const docs = [new Y.Doc(), new Y.Doc()];
      docs.forEach((doc, index) => {
        doc.clientID = index ? 73 : 41;
      });
      const fields = docs.map((doc) => attachOwnedReviewFields(doc, { protocol: '2' }));
      const flags = docs.map((doc, index) =>
        attachOwnedReviewFlags(doc, { protocol: '3', bootstrap: index ? 'join' : 'create' })
      );
      const original = new Y.Map<unknown>();
      docs[0]!.getMap('nodes').set('original', original);
      docs[0]!.getArray('children').insert(0, ['original', 'unrelated']);
      Y.applyUpdate(docs[1]!, Y.encodeStateAsUpdate(docs[0]!), 'relay');
      fields.forEach((scope) => scope.enable());
      const arrays = docs.map((doc) => doc.getArray<string>('children'));
      const nodes = docs.map((doc) => doc.getMap<Y.Map<unknown>>('nodes'));
      const histories = docs.map(
        (_, index) => new Y.UndoManager([arrays[index]!, nodes[index]!], { captureTimeout: 0 })
      );
      let connected = false;
      docs.forEach((doc, index) =>
        doc.on('update', (update, origin) => {
          if (connected && origin !== 'relay') Y.applyUpdate(docs[1 - index]!, update, 'relay');
        })
      );
      let fieldReceipt: readonly ReviewFieldReceipt[] = [];
      let flagReceipt: readonly ReviewFlagReceipt[] = [];
      const check = (children: string[], removed: boolean) => {
        for (const [index, array] of arrays.entries()) {
          expect(array.toArray()).toEqual(children);
          expect(nodes[index]!.get('original')!.get('removed') === true).toBe(removed);
        }
      };
      try {
        flags[0]!.capture(
          () =>
            fields[0]!.capture(
              () =>
                docs[0]!.transact(() => {
                  nodes[0]!.get('original')!.set('removed', true);
                  arrays[0]!.delete(0, 1);
                  arrays[0]!.insert(0, ['proposal']);
                }),
              2,
              (items) => {
                fieldReceipt = items;
              }
            ),
          1,
          (items) => {
            flagReceipt = items;
          }
        );
        expect(fieldReceipt).toHaveLength(2);
        expect(flagReceipt).toHaveLength(1);
        if (peerReplacement)
          flags[1]!.capture(
            () =>
              fields[1]!.capture(
                () =>
                  docs[1]!.transact(() => {
                    nodes[1]!.get('original')!.set('removed', true);
                    arrays[1]!.delete(0, 1);
                    arrays[1]!.insert(0, ['peer']);
                  }),
                2
              ),
            1
          );
        connected = true;
        Y.applyUpdate(docs[receiver]!, Y.encodeStateAsUpdate(docs[1 - receiver]!), 'relay');
        Y.applyUpdate(docs[1 - receiver]!, Y.encodeStateAsUpdate(docs[receiver]!), 'relay');
        docs[0]!.transact(() =>
          flags[0]!.retractInTransaction(flagReceipt, () =>
            fields[0]!.retractInTransaction(
              fieldReceipt,
              () => nodes[0]!.get('original')!.set('decision', 'rejected'),
              1
            )
          )
        );
        check(peerReplacement ? ['peer', 'unrelated'] : ['original', 'unrelated'], peerReplacement);
        if (peerReplacement) {
          expect(histories[1]!.undo()).not.toBeNull();
          check(['original', 'unrelated'], false);
          expect(histories[1]!.redo()).not.toBeNull();
          check(['peer', 'unrelated'], true);
          histories[1]!.undo();
        }
        expect(histories[0]!.undo()).not.toBeNull();
        check(['proposal', 'unrelated'], true);
        expect(histories[0]!.redo()).not.toBeNull();
        check(['original', 'unrelated'], false);
        if (peerReplacement) histories[1]!.redo();
        const cold = new Y.Doc(),
          coldFields = attachOwnedReviewFields(cold, { protocol: '2' }),
          coldFlags = attachOwnedReviewFlags(cold, { protocol: '3', bootstrap: 'join' });
        try {
          Y.applyUpdate(cold, Y.encodeStateAsUpdate(docs[0]!));
          coldFields.enable();
          coldFlags.assertHealthy();
          expect(cold.getArray('children').toArray()).toEqual(
            peerReplacement ? ['peer', 'unrelated'] : ['original', 'unrelated']
          );
          expect(
            cold.getMap<Y.Map<unknown>>('nodes').get('original')!.get('removed') === true
          ).toBe(peerReplacement);
        } finally {
          coldFields.dispose();
          coldFlags.dispose();
          cold.destroy();
        }
      } finally {
        histories.forEach((history) => history.destroy());
        fields.forEach((scope) => scope.dispose());
        flags.forEach((scope) => scope.dispose());
        docs.forEach((doc) => doc.destroy());
      }
    });

test('flag retraction refuses stale receipts before writes and retains original intent history', () => {
  const doc = new Y.Doc(),
    flags = attachOwnedReviewFlags(doc, { protocol: '3' });
  const map = new Y.Map<unknown>();
  doc.getMap('nodes').set('original', map);
  let receipt: readonly ReviewFlagReceipt[] = [];
  try {
    flags.capture(
      () => doc.transact(() => map.set('removed', true)),
      1,
      (items) => {
        receipt = items;
      }
    );
    for (const invalid of [
      [receipt[0]!, receipt[0]!],
      [{ ...receipt[0]!, id: '73:0' }],
      [{ ...receipt[0]!, record: 'invalid' }],
    ]) {
      const before = Y.encodeStateAsUpdate(doc);
      expect(() =>
        doc.transact(() => flags.retractInTransaction(invalid, () => map.set('bad', true)))
      ).toThrow('stale');
      expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
    }
    expect(() => flags.retractInTransaction(receipt, () => {})).toThrow('active local');
    doc.transact(() => flags.retractInTransaction(receipt, () => {}));
    expect(map.has('removed')).toBe(false);
    expect(doc.getMap(REVIEW_FLAG_INTENTS).size).toBe(1);
    expect(doc.getMap(REVIEW_FLAG_RETRACTIONS).size).toBe(1);
    const before = Y.encodeStateAsUpdate(doc);
    expect(() => doc.transact(() => flags.retractInTransaction(receipt, () => {}))).toThrow(
      'stale'
    );
    expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
    const legacy = new Y.Doc();
    try {
      Y.applyUpdate(legacy, Y.encodeStateAsUpdate(doc));
      expect(() => attachOwnedReviewFlags(legacy, { protocol: '2', bootstrap: 'join' })).toThrow(
        'Unsupported'
      );
    } finally {
      legacy.destroy();
    }
  } finally {
    flags.dispose();
    doc.destroy();
  }
});
