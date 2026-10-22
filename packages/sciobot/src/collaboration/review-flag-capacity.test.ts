import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { attachOwnedReviewFlags } from './review-flags';
import {
  flagAddress,
  MAX_REVIEW_FLAGS,
  MAX_REVIEW_FLAG_INTENTS,
  REVIEW_FLAG_INTENTS,
  REVIEW_FLAG_TARGETS,
} from './review-flag-codec';
import { seedReviewCapacity } from '../../test/support/review-capacity';

test('a review command that exceeds remaining flag capacity refuses before native flags, ownership rows, or history change', () => {
  const doc = new Y.Doc(),
    records = doc.getMap<Y.Map<unknown>>('records');
  records.set('revision', new Y.Map());
  const record = records.get('revision')!,
    id = record._item!.id,
    targets = doc.getMap<string>(REVIEW_FLAG_TARGETS);
  doc.transact(() => {
    targets.set('protocol', '1');
    for (let index = 0; index < MAX_REVIEW_FLAGS - 1; index++) {
      const target = {
        client: id.client,
        clock: id.clock,
        key: `retired_${index}`,
        baseline: 'missing' as const,
      };
      targets.set(flagAddress(target), JSON.stringify(target));
    }
  }, 'seed');
  const scope = attachOwnedReviewFlags(doc, {}),
    manager = new Y.UndoManager(records, {
      trackedOrigins: new Set(['review']),
      captureTimeout: 0,
    }),
    before = Y.encodeStateVector(doc);
  try {
    expect(() =>
      scope.capture(
        () =>
          doc.transact(() => {
            record.set('first', true);
            record.set('second', true);
          }, 'review'),
        2
      )
    ).toThrow('record limit');
    expect([record.get('first'), record.get('second')]).toEqual([undefined, undefined]);
    expect(doc.getMap(REVIEW_FLAG_INTENTS).size).toBe(0);
    expect(targets.size).toBe(MAX_REVIEW_FLAGS);
    expect(manager.undoStack).toHaveLength(0);
    expect(Y.encodeStateVector(doc)).toEqual(before);
  } finally {
    manager.destroy();
    scope.dispose();
    doc.destroy();
  }
}, 30000);

for (const kind of ['targets', 'intents'] as const)
  for (const remaining of [1, 2])
    test(`a complete two-flag command checks the exact ${kind} boundary before mutation; remaining=${remaining}`, () => {
      const doc = new Y.Doc(),
        scope = attachOwnedReviewFlags(doc, {}),
        records = doc.getMap<Y.Map<unknown>>('records');
      records.set('revision', new Y.Map());
      seedReviewCapacity(doc, kind, remaining);
      const map = records.get('revision')!,
        manager = new Y.UndoManager(records, {
          trackedOrigins: new Set(['review']),
          captureTimeout: 0,
        }),
        before = Y.encodeStateVector(doc),
        targets = doc.getMap(REVIEW_FLAG_TARGETS),
        intents = doc.getMap(REVIEW_FLAG_INTENTS),
        targetCount = targets.size,
        intentCount = intents.size;
      let calls = 0;
      const commit = () =>
        scope.capture(() => {
          calls++;
          doc.transact(() => {
            map.set('first', true);
            map.set('second', true);
          }, 'review');
        }, 2);
      try {
        if (remaining === 1) {
          expect(commit).toThrow('record limit');
          expect(calls).toBe(0);
          expect([map.get('first'), map.get('second')]).toEqual([undefined, undefined]);
          expect(targets.size).toBe(targetCount);
          expect(intents.size).toBe(intentCount);
          expect(manager.undoStack).toHaveLength(0);
          expect(Y.encodeStateVector(doc)).toEqual(before);
        } else {
          commit();
          expect(calls).toBe(1);
          expect([map.get('first'), map.get('second')]).toEqual([true, true]);
          expect(targets.size).toBe(targetCount + 2);
          expect(intents.size).toBe(intentCount + 2);
          expect(kind === 'targets' ? targets.size - 1 : intents.size).toBe(
            kind === 'targets' ? MAX_REVIEW_FLAGS : MAX_REVIEW_FLAG_INTENTS
          );
          expect(manager.undoStack).toHaveLength(1);
          manager.undo();
          expect([map.get('first'), map.get('second')]).toEqual([undefined, undefined]);
          expect(intents.size).toBe(intentCount);
          expect(manager.undoStack).toHaveLength(0);
          manager.redo();
          expect([map.get('first'), map.get('second')]).toEqual([true, true]);
          expect(intents.size).toBe(intentCount + 2);
        }
      } finally {
        manager.destroy();
        scope.dispose();
        doc.destroy();
      }
    }, 30000);

test('flag reservations validate counts, release unused headroom, and capture a repeated flag once per command', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFlags(doc, {}),
    root = doc.getMap<Y.Map<unknown>>('records');
  root.set('revision', new Y.Map());
  const map = root.get('revision')!,
    before = Y.encodeStateVector(doc);
  try {
    for (const budget of [-1, 0.5, NaN, Infinity, Number.MAX_SAFE_INTEGER + 1])
      expect(() =>
        scope.capture(() => {
          throw new Error('Must not run');
        }, budget)
      ).toThrow('bounded');
    expect(Y.encodeStateVector(doc)).toEqual(before);
    expect(() =>
      scope.capture(() => {
        throw new Error('Pure refusal');
      }, 10)
    ).toThrow('Pure refusal');
    scope.capture(
      () =>
        doc.transact(() => {
          map.set('removed', true);
          map.set('removed', true);
        }, 'review'),
      1
    );
    expect(map.get('removed')).toBe(true);
    expect(doc.getMap(REVIEW_FLAG_INTENTS).size).toBe(1);
  } finally {
    scope.dispose();
    doc.destroy();
  }
});

test('nested capture reserves the parent budget and refuses a child before either can overbook the final target slot', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFlags(doc, {}),
    root = doc.getMap<Y.Map<unknown>>('records');
  root.set('revision', new Y.Map());
  seedReviewCapacity(doc, 'targets', 1);
  const map = root.get('revision')!;
  try {
    scope.capture(() => {
      const before = Y.encodeStateVector(doc);
      expect(() => scope.capture(() => doc.transact(() => map.set('child', true)), 1)).toThrow(
        'record limit'
      );
      expect(map.has('child')).toBe(false);
      expect(Y.encodeStateVector(doc)).toEqual(before);
      doc.transact(() => map.set('parent', true), 'review');
    }, 1);
    expect(map.get('parent')).toBe(true);
    expect(doc.getMap(REVIEW_FLAG_TARGETS).size).toBe(MAX_REVIEW_FLAGS + 1);
    expect(doc.getMap(REVIEW_FLAG_INTENTS).size).toBe(1);
  } finally {
    scope.dispose();
    doc.destroy();
  }
}, 30000);

test('capacity admission waits for incomplete updates and resumes after their dependencies arrive', () => {
  const doc = new Y.Doc(),
    peer = new Y.Doc(),
    scope = attachOwnedReviewFlags(doc, {}),
    root = doc.getMap<Y.Map<unknown>>('records');
  root.set('revision', new Y.Map());
  const map = root.get('revision')!,
    remote = peer.getMap('remote');
  remote.set('value', 'first');
  const prefix = Y.encodeStateVector(peer);
  remote.set('value', 'second');
  try {
    Y.applyUpdate(doc, Y.encodeStateAsUpdate(peer, prefix), 'relay');
    expect(doc.store.pendingStructs).not.toBeNull();
    const before = Y.encodeStateVector(doc);
    let calls = 0;
    const commit = () =>
      scope.capture(() => {
        calls++;
        doc.transact(() => map.set('removed', true), 'review');
      }, 1);
    expect(commit).toThrow('complete room updates');
    expect(calls).toBe(0);
    expect(map.has('removed')).toBe(false);
    expect(Y.encodeStateVector(doc)).toEqual(before);
    Y.applyUpdate(doc, Y.encodeStateAsUpdate(peer), 'relay');
    expect(doc.store.pendingStructs).toBeNull();
    commit();
    expect(calls).toBe(1);
    expect(map.get('removed')).toBe(true);
    expect(doc.getMap(REVIEW_FLAG_INTENTS).size).toBe(1);
  } finally {
    scope.dispose();
    doc.destroy();
    peer.destroy();
  }
});

test('an actor cannot spend intent capacity retained for another actor redo', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFlags(doc, {}),
    root = doc.getMap<Y.Map<unknown>>('records');
  root.set('revision', new Y.Map());
  seedReviewCapacity(doc, 'intents', 2);
  const map = root.get('revision')!,
    alice = new Y.UndoManager(root, { trackedOrigins: new Set(['alice']), captureTimeout: 0 }),
    bob = new Y.UndoManager(root, { trackedOrigins: new Set(['bob']), captureTimeout: 0 });
  try {
    scope.capture(
      () =>
        doc.transact(() => {
          map.set('first', true);
          map.set('second', true);
        }, 'alice'),
      2
    );
    expect(doc.getMap(REVIEW_FLAG_INTENTS).size).toBe(MAX_REVIEW_FLAG_INTENTS);
    alice.undo();
    expect(doc.getMap(REVIEW_FLAG_INTENTS).size).toBe(MAX_REVIEW_FLAG_INTENTS - 2);
    const before = Y.encodeStateVector(doc);
    let calls = 0;
    expect(() =>
      scope.capture(() => {
        calls++;
        doc.transact(() => map.set('peer', true), 'bob');
      }, 1)
    ).toThrow('record limit');
    expect(calls).toBe(0);
    expect(map.has('peer')).toBe(false);
    expect(Y.encodeStateVector(doc)).toEqual(before);
    expect(bob.undoStack).toHaveLength(0);
    alice.redo();
    expect([map.get('first'), map.get('second')]).toEqual([true, true]);
    expect(doc.getMap(REVIEW_FLAG_INTENTS).size).toBe(MAX_REVIEW_FLAG_INTENTS);
  } finally {
    alice.destroy();
    bob.destroy();
    scope.dispose();
    doc.destroy();
  }
}, 30000);
