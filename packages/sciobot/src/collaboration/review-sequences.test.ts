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
