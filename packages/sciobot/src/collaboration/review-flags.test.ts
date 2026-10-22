import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { attachOwnedReviewFlags } from './review-flags';
import {
  decodeFlagTarget,
  flagAddress,
  REVIEW_FLAG_INTENTS,
  REVIEW_FLAG_TARGETS,
  validateFlagIntent,
} from './review-flag-codec';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    test(`review flags retain each actor's ownership on existing Yjs frames; ${clients}/first=${first}`, () => {
      const docs = [new Y.Doc(), new Y.Doc()],
        scopes = docs.map((doc) => attachOwnedReviewFlags(doc, {}));
      docs.forEach((doc, index) => (doc.clientID = clients[index]!));
      const roots = docs.map((doc) => doc.getMap<Y.Map<unknown>>('records'));
      roots[0]!.set('revision', new Y.Map());
      roots[0]!.get('revision')!.set('explicit', false);
      Y.applyUpdate(docs[1]!, Y.encodeStateAsUpdate(docs[0]!), 'relay');
      const managers = docs.map(
          (_doc, index) =>
            new Y.UndoManager(roots[index]!, {
              trackedOrigins: new Set([index]),
              captureTimeout: 0,
            })
        ),
        flags = () => roots.map((root) => root.get('revision')!),
        sync = () => {
          Y.applyUpdate(docs[first]!, Y.encodeStateAsUpdate(docs[1 - first]!), 'relay');
          Y.applyUpdate(docs[1 - first]!, Y.encodeStateAsUpdate(docs[first]!), 'relay');
          Y.applyUpdate(docs[first]!, Y.encodeStateAsUpdate(docs[1 - first]!), 'relay');
        },
        accepted = () => {
          sync();
          for (const map of flags()) {
            expect(map.get('removed')).toBe(true);
            expect(map.get('explicit')).toBe(true);
          }
        },
        pending = () => {
          sync();
          for (const map of flags()) {
            expect(map.has('removed')).toBe(false);
            expect(map.get('explicit')).toBe(false);
          }
        };
      try {
        scopes.forEach((scope, index) =>
          scope.capture(
            () =>
              docs[index]!.transact(() => {
                flags()[index]!.set('removed', true);
                flags()[index]!.set('explicit', true);
              }, index),
            2
          )
        );
        accepted();
        expect(managers.map((manager) => manager.undoStack.length)).toEqual([1, 1]);
        managers[first]!.undo();
        accepted();
        managers[1 - first]!.undo();
        pending();
        managers[first]!.redo();
        accepted();
        managers[1 - first]!.redo();
        accepted();
        managers[first]!.undo();
        accepted();
        managers[1 - first]!.undo();
        pending();
        expect(managers.map((manager) => manager.undoStack.length)).toEqual([0, 0]);
        expect(docs.map((doc) => doc.getMap(REVIEW_FLAG_INTENTS).size)).toEqual([0, 0]);
        // Unrelated maps and ordinary writes never enter this protocol.
        roots[0]!.get('revision')!.set('ordinary', true);
        expect(docs[0]!.getMap(REVIEW_FLAG_INTENTS).size).toBe(0);
      } finally {
        managers.forEach((manager) => manager.destroy());
        scopes.forEach((scope) => scope.dispose());
        docs.forEach((doc) => doc.destroy());
      }
    });

test('review flag codecs reject ambiguous targets, unsafe keys, future fields, and invalid intent identities', () => {
  const target = { client: 41, clock: 7, key: 'removed', baseline: 'missing' as const },
    address = flagAddress(target);
  expect(decodeFlagTarget(address, JSON.stringify(target))).toEqual(target);
  for (const value of [
    null,
    [],
    { ...target, version: 2 },
    { ...target, clock: -1 },
    { ...target, client: Number.MAX_SAFE_INTEGER + 1 },
    { ...target, baseline: 'true' },
    { ...target, key: '' },
    { ...target, key: '__proto__' },
    { ...target, key: 'constructor' },
    { ...target, key: 'x\u0000y' },
    { ...target, key: 'x'.repeat(257) },
  ])
    expect(() => decodeFlagTarget(address, JSON.stringify(value))).toThrow();
  expect(() => decodeFlagTarget('other', JSON.stringify(target))).toThrow('address');
  expect(() => decodeFlagTarget(address, '{')).toThrow('JSON');
  expect(() => decodeFlagTarget(address, 'x'.repeat(2049))).toThrow('bounded');
  for (const key of ['', '41:-1', 'x:7', '41:9007199254740992'])
    expect(() => validateFlagIntent(key, address)).toThrow();
  expect(() => validateFlagIntent('41:7', 1)).toThrow();
});

test('review flag health rejects poisoned room data and prevents later command capture', () => {
  const doc = new Y.Doc(),
    errors: Error[] = [],
    scope = attachOwnedReviewFlags(doc, { onError: (error) => errors.push(error) });
  try {
    expect(() => doc.getMap(REVIEW_FLAG_TARGETS).set('protocol', '2')).toThrow(
      'Unsupported review flag protocol'
    );
    expect(errors).toHaveLength(1);
    expect(() => scope.capture(() => {}, 0)).toThrow('Unsupported review flag protocol');
    expect(() => scope.assertHealthy()).toThrow('Unsupported review flag protocol');
    scope.dispose();
    scope.dispose();
    expect(() => scope.capture(() => {}, 0)).toThrow('no longer active');
  } finally {
    scope.dispose();
    doc.destroy();
  }
});

test('review flag capture requires the native transaction and leaves unrelated boolean fields unchanged', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFlags(doc, {}),
    root = doc.getMap<Y.Map<unknown>>('nodes');
  root.set('record', new Y.Map());
  const map = root.get('record')!;
  try {
    expect(() => scope.capture(() => map.set('removed', true), 1)).toThrow('native transaction');
    expect(map.has('removed')).toBe(false);
    expect(doc.getMap(REVIEW_FLAG_INTENTS).size).toBe(0);
    map.set('ordinary', true);
    expect(doc.getMap(REVIEW_FLAG_TARGETS).size).toBe(0);
    map.set('wrong-type', 'keep');
    expect(() => scope.capture(() => doc.transact(() => map.set('wrong-type', true)), 1)).toThrow(
      'non-boolean'
    );
    expect(map.get('wrong-type')).toBe('keep');
    expect(doc.getMap(REVIEW_FLAG_INTENTS).size).toBe(0);
  } finally {
    scope.dispose();
    doc.destroy();
  }
});

test('synchronous foreign transactions inside a command keep their own history and receive no local ownership intent', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFlags(doc, {}),
    root = doc.getMap<Y.Map<unknown>>('records');
  root.set('local', new Y.Map());
  root.set('peer', new Y.Map());
  const local = root.get('local')!,
    peer = root.get('peer')!,
    manager = new Y.UndoManager(root, { trackedOrigins: new Set(['local']), captureTimeout: 0 });
  try {
    scope.capture(() => {
      doc.transact(() => local.set('removed', true), 'local');
      doc.transact(() => peer.set('removed', true), 'relay');
    }, 1);
    expect(doc.getMap(REVIEW_FLAG_INTENTS).size).toBe(1);
    expect(manager.undoStack).toHaveLength(1);
    manager.undo();
    expect(local.has('removed')).toBe(false);
    expect(peer.get('removed')).toBe(true);
    peer.delete('removed');
    expect(peer.has('removed')).toBe(false);
    expect(doc.getMap(REVIEW_FLAG_INTENTS).size).toBe(0);
  } finally {
    manager.destroy();
    scope.dispose();
    doc.destroy();
  }
});
