import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { attachOwnedReviewFields } from './review-fields';

for (const receiver of [0, 1] as const)
  for (const first of [0, 1] as const)
    test(
      'concurrent string sequence moves retain peer order through both undo and redo; ' +
        receiver +
        '/' +
        first,
      () => {
        const docs = [new Y.Doc(), new Y.Doc()],
          scopes = docs.map((doc) => attachOwnedReviewFields(doc));
        docs[0]!.clientID = 41;
        docs[1]!.clientID = 73;
        const array = docs[0]!.getArray<string>('order');
        array.insert(0, ['shade', 'marker', 'tail']);
        Y.applyUpdate(docs[1]!, Y.encodeStateAsUpdate(docs[0]!), 'relay');
        const arrays = docs.map((doc) => doc.getArray<string>('order')),
          history = arrays.map((value) => new Y.UndoManager(value, { captureTimeout: 0 }));
        scopes.forEach((scope) => scope.enable());
        let connected = false;
        docs.forEach((doc, index) =>
          doc.on('update', (update, origin) => {
            if (connected && origin !== 'relay') Y.applyUpdate(docs[1 - index]!, update, 'relay');
          })
        );
        const check = (expected: string[]) =>
          arrays.forEach((value) => expect(value.toArray()).toEqual(expected));
        try {
          arrays.forEach((value, index) =>
            scopes[index]!.capture(
              () =>
                docs[index]!.transact(() => {
                  value.delete(0, 1);
                  value.insert(1, ['shade']);
                }),
              2
            )
          );
          connected = true;
          Y.applyUpdate(docs[receiver]!, Y.encodeStateAsUpdate(docs[1 - receiver]!), 'relay');
          Y.applyUpdate(docs[1 - receiver]!, Y.encodeStateAsUpdate(docs[receiver]!), 'relay');
          // Both native placement entries remain. The transport selects one logical node.
          check(['marker', 'shade', 'shade', 'tail']);
          expect(history[first]!.undo()).not.toBeNull();
          check(['marker', 'shade', 'tail']);
          expect(history[1 - first]!.undo()).not.toBeNull();
          check(['shade', 'marker', 'tail']);
          expect(history[first]!.redo()).not.toBeNull();
          check(['marker', 'shade', 'tail']);
          expect(history[1 - first]!.redo()).not.toBeNull();
          check(['marker', 'shade', 'shade', 'tail']);
          expect(history[first]!.undo()).not.toBeNull();
          check(['marker', 'shade', 'tail']);
          expect(history[1 - first]!.undo()).not.toBeNull();
          check(['shade', 'marker', 'tail']);
        } finally {
          history.forEach((value) => value.destroy());
          scopes.forEach((value) => value.dispose());
          docs.forEach((value) => value.destroy());
        }
      }
    );

test('ordinary removal reuses a restored entry address and remains independent of earlier review history', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFields(doc),
    array = doc.getArray<string>('order');
  array.insert(0, ['A', 'B']);
  scope.enable();
  const history = new Y.UndoManager(array, { captureTimeout: 0 });
  try {
    scope.capture(() => doc.transact(() => array.delete(0, 1)), 1);
    expect(array.toArray()).toEqual(['B']);
    expect(history.undo()).not.toBeNull();
    expect(array.toArray()).toEqual(['A', 'B']);
    array.delete(0, 1);
    expect(array.toArray()).toEqual(['B']);
    expect(doc.getMap('sciobot-review-field-targets-v1').size).toBe(2);
    for (let pass = 0; pass < 2; pass++) {
      expect(history.undo()).not.toBeNull();
      expect(array.toArray()).toEqual(['A', 'B']);
      expect(history.redo()).not.toBeNull();
      expect(array.toArray()).toEqual(['B']);
    }
  } finally {
    history.destroy();
    scope.dispose();
    doc.destroy();
  }
});

test('batched string removals retain split item positions through repeated undo and redo', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFields(doc),
    array = doc.getArray<string>('order');
  array.insert(0, ['A', 'B', 'C', 'D']);
  scope.enable();
  const history = new Y.UndoManager(array, { captureTimeout: 0 });
  try {
    scope.capture(() => doc.transact(() => array.delete(1, 2)), 2);
    expect(array.toArray()).toEqual(['A', 'D']);
    for (let pass = 0; pass < 3; pass++) {
      expect(history.undo()).not.toBeNull();
      expect(array.toArray()).toEqual(['A', 'B', 'C', 'D']);
      expect(history.redo()).not.toBeNull();
      expect(array.toArray()).toEqual(['A', 'D']);
    }
  } finally {
    history.destroy();
    scope.dispose();
    doc.destroy();
  }
});

test('undo of one native gesture does not restore entries created and removed within that gesture', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFields(doc),
    array = doc.getArray<string>('order');
  array.insert(0, ['original']);
  scope.enable();
  const history = new Y.UndoManager(array, { captureTimeout: 1000 });
  try {
    doc.transact(() => array.insert(1, ['temporary record']));
    doc.transact(() => array.delete(1, 1));
    history.stopCapturing();
    expect(array.toArray()).toEqual(['original']);
    expect(history.undoStack).toHaveLength(1);
    for (let pass = 0; pass < 3; pass++) {
      expect(history.undo()).not.toBeNull();
      expect(array.toArray()).toEqual(['original']);
      expect(history.redo()).not.toBeNull();
      expect(array.toArray()).toEqual(['original']);
    }
  } finally {
    history.destroy();
    scope.dispose();
    doc.destroy();
  }
});

test('later ordinary insertion preserves the position of an already restored baseline entry', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFields(doc),
    array = doc.getArray<string>('order');
  array.insert(0, ['original']);
  scope.enable();
  const history = new Y.UndoManager(array, { captureTimeout: 0 });
  try {
    array.delete(0, 1);
    expect(history.undo()).not.toBeNull();
    expect(array.toArray()).toEqual(['original']);
    array.insert(1, ['later']);
    expect(array.toArray()).toEqual(['original', 'later']);
    expect(history.undo()).not.toBeNull();
    expect(array.toArray()).toEqual(['original']);
    expect(history.redo()).not.toBeNull();
    expect(array.toArray()).toEqual(['original', 'later']);
  } finally {
    history.destroy();
    scope.dispose();
    doc.destroy();
  }
});

for (const reverse of [false, true])
  for (const receiver of [0, 1] as const)
    for (const peerIndex of [1, 2, 3])
      test(`separate removals in one gesture restore original sibling anchors: ${reverse}/${receiver}/${peerIndex}`, () => {
        const docs = [new Y.Doc(), new Y.Doc()];
        docs.forEach((doc, index) => {
          doc.clientID = index ? 73 : 41;
        });
        const scopes = docs.map((doc, index) =>
          attachOwnedReviewFields(doc, { protocol: '3', bootstrap: index ? 'join' : 'create' })
        );
        const arrays = docs.map((doc) => doc.getArray<string>('children'));
        arrays[0]!.insert(0, ['head', 'middle', 'tail']);
        Y.applyUpdate(docs[1]!, Y.encodeStateAsUpdate(docs[0]!), 'relay');
        scopes.forEach((scope) => scope.enable());
        const history = new Y.UndoManager(arrays[0]!, { captureTimeout: 0 });
        let connected = false;
        docs.forEach((doc, index) =>
          doc.on('update', (update, origin) => {
            if (connected && origin !== 'relay') Y.applyUpdate(docs[1 - index]!, update, 'relay');
          })
        );
        try {
          scopes[0]!.capture(
            () =>
              docs[0]!.transact(() => {
                arrays[0]!.delete(reverse ? 2 : 1, 1);
                arrays[0]!.delete(1, 1);
              }),
            2
          );
          arrays[1]!.insert(peerIndex, ['peer']);
          const restored = ['head', 'middle', 'tail'];
          restored.splice(peerIndex, 0, 'peer');
          connected = true;
          Y.applyUpdate(docs[receiver]!, Y.encodeStateAsUpdate(docs[1 - receiver]!), 'relay');
          Y.applyUpdate(docs[1 - receiver]!, Y.encodeStateAsUpdate(docs[receiver]!), 'relay');
          for (const array of arrays) expect(array.toArray()).toEqual(['head', 'peer']);
          for (let pass = 0; pass < 3; pass++) {
            expect(history.undo()).not.toBeNull();
            for (const array of arrays) expect(array.toArray()).toEqual(restored);
            expect(history.redo()).not.toBeNull();
            for (const array of arrays) expect(array.toArray()).toEqual(['head', 'peer']);
          }
        } finally {
          history.destroy();
          scopes.forEach((scope) => scope.dispose());
          docs.forEach((doc) => doc.destroy());
        }
      });

test('ordered field profile refuses old rooms and older readers without rewriting their headers', () => {
  const old = new Y.Doc(),
    fresh = new Y.Doc();
  old.getMap('sciobot-review-field-targets-v1').set('protocol', '2');
  const before = Y.encodeStateAsUpdate(old);
  try {
    expect(() => attachOwnedReviewFields(old, { protocol: '3', bootstrap: 'create' })).toThrow(
      'Unsupported owned review field protocol'
    );
    expect(Y.encodeStateAsUpdate(old)).toEqual(before);
    expect(() => attachOwnedReviewFields(fresh, { protocol: '3' })).toThrow('bootstrap mode');
    const scope = attachOwnedReviewFields(fresh, { protocol: '3', bootstrap: 'create' });
    const cold = new Y.Doc();
    try {
      Y.applyUpdate(cold, Y.encodeStateAsUpdate(fresh));
      expect(() => attachOwnedReviewFields(cold, { protocol: '2' })).toThrow(
        'Unsupported owned review field protocol'
      );
      expect(cold.getMap('sciobot-review-field-targets-v1').get('protocol')).toBe('3');
    } finally {
      cold.destroy();
      scope.dispose();
    }
  } finally {
    old.destroy();
    fresh.destroy();
  }
});

test('an ordered join refuses insertion before room initialization without changing shared data', () => {
  const doc = new Y.Doc();
  const scope = attachOwnedReviewFields(doc, { protocol: '3', bootstrap: 'join' });
  const array = doc.getArray<string>('children');
  scope.enable();
  const before = Y.encodeStateAsUpdate(doc);
  try {
    expect(() => array.insert(0, ['premature child'])).toThrow('room header before editing');
    expect(() => scope.capture(() => doc.getText('body').insert(0, 'premature text'), 0)).toThrow(
      'room header before editing'
    );
    expect(array.toArray()).toEqual([]);
    expect(doc.getText('body').toString()).toBe('');
    expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
    doc.getMap('sciobot-review-field-targets-v1').set('protocol', '3');
    scope.capture(() => doc.getText('body').insert(0, 'ready'), 0);
    expect(doc.getText('body').toString()).toBe('ready');
  } finally {
    scope.dispose();
    doc.destroy();
  }
});
