import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { TEXT_DELETIONS_KEY } from '../../../pro/src/collaboration/document/paragraph-text-deletions';
import { TEXT_MARKS_KEY } from '../../../pro/src/collaboration/document/paragraph-text-marks';
import { attachOwnedReviewFields } from './review-fields';
import {
  REVIEW_FIELD_TARGETS,
  REVIEW_FIELD_INTENTS,
  fieldAddress,
  encodeFieldRecord,
  decodeFieldTarget,
  decodeFieldIntent,
} from './review-field-codec';
function pair(nested = false) {
  const docs = [new Y.Doc(), new Y.Doc()],
    fields = docs.map((doc) => attachOwnedReviewFields(doc));
  docs[0]!.clientID = 41;
  docs[1]!.clientID = 73;
  const root = docs[0]!.getMap<unknown>('native');
  if (nested) {
    const child = new Y.Map<string>();
    child.set('shape', 'original');
    root.set('node', child);
  } else root.set('shape', 'original');
  Y.applyUpdate(docs[1]!, Y.encodeStateAsUpdate(docs[0]!), 'relay');
  const maps = docs.map((doc) =>
    nested ? doc.getMap<Y.Map<string>>('native').get('node')! : doc.getMap<string>('native')
  );
  fields.forEach((field) => field.enable());
  const history = maps.map((map) => new Y.UndoManager(map, { captureTimeout: 0 }));
  let connected = true;
  docs.forEach((doc, index) =>
    doc.on('update', (update, origin) => {
      if (connected && origin !== 'relay') Y.applyUpdate(docs[1 - index]!, update, 'relay');
    })
  );
  return {
    docs,
    fields,
    maps,
    history,
    pause() {
      connected = false;
    },
    connect(receiver: 0 | 1) {
      connected = true;
      Y.applyUpdate(docs[receiver]!, Y.encodeStateAsUpdate(docs[1 - receiver]!), 'relay');
      Y.applyUpdate(docs[1 - receiver]!, Y.encodeStateAsUpdate(docs[receiver]!), 'relay');
    },
    destroy() {
      history.forEach((item) => item.destroy());
      fields.forEach((item) => item.dispose());
      docs.forEach((item) => item.destroy());
    },
  };
}
for (const nested of [false, true])
  for (const receiver of [0, 1] as const)
    for (const value of ['accepted', null] as const)
      test(
        'concurrent string field ' +
          (value === null ? 'removal' : 'write') +
          ' keeps peer ownership until both undo; nested=' +
          nested +
          '/' +
          receiver,
        () => {
          const fixture = pair(nested);
          try {
            fixture.pause();
            fixture.maps.forEach((map, index) =>
              fixture.fields[index]!.capture(() => {
                fixture.docs[index]!.transact(() =>
                  value === null ? map.delete('shape') : map.set('shape', value)
                );
              }, 1)
            );
            fixture.connect(receiver);
            const check = (expected: string | undefined) => {
              for (const map of fixture.maps) {
                expect(map.get('shape')).toBe(expected);
                expect(map.has('shape')).toBe(expected !== undefined);
              }
            };
            check(value ?? undefined);
            expect(fixture.history[0]!.undo()).not.toBeNull();
            check(value ?? undefined);
            expect(fixture.history[1]!.undo()).not.toBeNull();
            check('original');
            expect(fixture.history[0]!.redo()).not.toBeNull();
            check(value ?? undefined);
            expect(fixture.history[1]!.redo()).not.toBeNull();
            check(value ?? undefined);
            const reopened = new Y.Doc(),
              scope = attachOwnedReviewFields(reopened);
            try {
              Y.applyUpdate(reopened, Y.encodeStateAsUpdate(fixture.docs[0]!));
              scope.enable();
              const map = nested
                ? reopened.getMap<Y.Map<string>>('native').get('node')!
                : reopened.getMap<string>('native');
              expect(map.get('shape')).toBe(value ?? undefined);
            } finally {
              scope.dispose();
              reopened.destroy();
            }
          } finally {
            fixture.destroy();
          }
        }
      );
for (const receiver of [0, 1] as const)
  test(
    'native root keys retain control characters, absent values and empty strings; receiver=' +
      receiver,
    () => {
      const fixture = pair();
      try {
        const key = 'node\u001fw\u001fval';
        fixture.pause();
        fixture.maps.forEach((map, index) =>
          fixture.fields[index]!.capture(() => {
            fixture.docs[index]!.transact(() => map.set(key, ''));
          }, 1)
        );
        fixture.connect(receiver);
        for (const map of fixture.maps) expect(map.get(key)).toBe('');
        fixture.history[0]!.undo();
        for (const map of fixture.maps) expect(map.get(key)).toBe('');
        fixture.history[1]!.undo();
        for (const map of fixture.maps) expect(map.has(key)).toBe(false);
      } finally {
        fixture.destroy();
      }
    }
  );
test('ordinary later native field edits survive the earlier actor undo', () => {
  const fixture = pair();
  try {
    fixture.fields[0]!.capture(
      () => fixture.docs[0]!.transact(() => fixture.maps[0]!.set('shape', 'accepted')),
      1
    );
    fixture.maps[1]!.set('shape', 'manual');
    for (const map of fixture.maps) expect(map.get('shape')).toBe('manual');
    fixture.history[0]!.undo();
    for (const map of fixture.maps) expect(map.get('shape')).toBe('manual');
    fixture.history[1]!.undo();
    for (const map of fixture.maps) expect(map.get('shape')).toBe('original');
  } finally {
    fixture.destroy();
  }
});
test('field capture coalesces repeated writes and releases nested reservations after refusal', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFields(doc, { limits: { targets: 2, intents: 2 } }),
    map = doc.getMap<string>('native');
  map.set('shape', 'original');
  scope.enable();
  try {
    scope.capture(
      () =>
        doc.transact(() => {
          map.set('shape', 'first');
          map.set('shape', 'last');
          expect(() => scope.capture(() => {}, 2)).toThrow('record limit');
        }),
      1
    );
    expect(map.get('shape')).toBe('last');
    expect(doc.getMap(REVIEW_FIELD_INTENTS).size).toBe(1);
    expect(() =>
      scope.capture(() => {
        throw new Error('stop');
      }, 1)
    ).toThrow('stop');
    scope.capture(() => doc.transact(() => map.set('other', 'retained')), 1);
    expect(map.get('other')).toBe('retained');
  } finally {
    scope.dispose();
    doc.destroy();
  }
});
for (const kind of ['targets', 'intents'] as const)
  test('complete ' + kind + ' admission refuses before callback and retains redo capacity', () => {
    const doc = new Y.Doc(),
      scope = attachOwnedReviewFields(doc, {
        limits: { targets: kind === 'targets' ? 1 : 4, intents: kind === 'intents' ? 1 : 4 },
      }),
      map = doc.getMap<string>('native');
    map.set('shape', 'original');
    scope.enable();
    const history = new Y.UndoManager(map, { captureTimeout: 0 });
    try {
      const before = Y.encodeStateVector(doc);
      let calls = 0;
      expect(() =>
        scope.capture(() => {
          calls++;
          map.set('shape', 'bad');
        }, 2)
      ).toThrow('record limit');
      expect(calls).toBe(0);
      expect(Y.encodeStateVector(doc)).toEqual(before);
      expect(history.canUndo()).toBe(false);
      scope.capture(() => doc.transact(() => map.set('shape', 'accepted')), 1);
      history.undo();
      expect(map.get('shape')).toBe('original');
      expect(() =>
        scope.capture(() => {
          calls++;
          map.set('other', 'bad');
        }, 1)
      ).toThrow('record limit');
      expect(calls).toBe(0);
      expect(history.redo()).not.toBeNull();
      expect(map.get('shape')).toBe('accepted');
    } finally {
      history.destroy();
      scope.dispose();
      doc.destroy();
    }
  });
test('bootstrap and new nested maps keep their native identities without ownership records', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFields(doc),
    root = doc.getMap<unknown>('native');
  root.set('shape', 'original');
  expect(doc.getMap(REVIEW_FIELD_TARGETS).size).toBe(0);
  scope.enable();
  try {
    doc.transact(() => {
      const node = new Y.Map<string>();
      root.set('new', node);
      node.set('shape', 'created');
    });
    expect(doc.getMap(REVIEW_FIELD_TARGETS).size).toBe(0);
    const node = root.get('new') as Y.Map<string>;
    node.set('shape', 'changed');
    expect(doc.getMap(REVIEW_FIELD_TARGETS).size).toBe(2);
  } finally {
    scope.dispose();
    doc.destroy();
  }
});
test('malformed, self-targeting, and unsupported field data refuses attachment', () => {
  for (const mode of ['version', 'address', 'self', 'missing-target'] as const) {
    const doc = new Y.Doc(),
      targets = doc.getMap<string>(REVIEW_FIELD_TARGETS),
      intents = doc.getMap<string>(REVIEW_FIELD_INTENTS);
    targets.set('protocol', mode === 'version' ? '2' : '1');
    if (mode === 'address')
      targets.set('wrong', encodeFieldRecord({ root: 'native', key: 'shape', baseline: null }));
    if (mode === 'self') {
      const target = { root: REVIEW_FIELD_INTENTS, key: 'shape', baseline: null };
      targets.set(fieldAddress(target), encodeFieldRecord(target));
    }
    if (mode === 'missing-target')
      intents.set('1:0', encodeFieldRecord({ target: 'missing', value: 'bad', stamp: [1, 1, 0] }));
    try {
      expect(() => attachOwnedReviewFields(doc)).toThrow();
    } finally {
      doc.destroy();
    }
  }
});
test('field errors remain explicit, reject scalar type changes before writes, and disposal is idempotent', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFields(doc),
    map = doc.getMap<unknown>('native');
  map.set('shape', 'original');
  scope.enable();
  try {
    map.set('shape', 'accepted');
    expect(() => map.set('shape', true)).toThrow('value type');
    expect(map.get('shape')).toBe('accepted');
    expect(() => attachOwnedReviewFields(doc)).toThrow('already attached');
    scope.dispose();
    scope.dispose();
    expect(() => scope.capture(() => {}, 0)).toThrow('no longer active');
  } finally {
    scope.dispose();
    doc.destroy();
  }
});
test('field codecs reject ambiguous metadata and retain exact root keys', () => {
  const target = { root: 'native', key: 'node\u001fw\u001fval', baseline: '' };
  expect(decodeFieldTarget(fieldAddress(target), encodeFieldRecord(target))).toEqual(target);
  for (const raw of [
    { ...target, unknown: true },
    { ...target, key: '' },
    { ...target, client: 1 },
    [],
  ])
    expect(() => decodeFieldTarget(fieldAddress(target), JSON.stringify(raw))).toThrow();
  for (const raw of [
    { target: 'x', value: 'v', stamp: [1, 2] },
    { target: 'x', value: 3, stamp: [1, 2, 3] },
    { target: 'x', value: 'v', stamp: [1, -1, 3] },
  ])
    expect(() => decodeFieldIntent('1:0', JSON.stringify(raw))).toThrow();
});

test('field admission waits for complete updates and resumes without poisoning the scope', () => {
  const doc = new Y.Doc(),
    peer = new Y.Doc(),
    scope = attachOwnedReviewFields(doc),
    map = doc.getMap<string>('native'),
    remote = peer.getMap<string>('remote');
  map.set('shape', 'original');
  scope.enable();
  remote.set('shape', 'first');
  const prefix = Y.encodeStateVector(peer);
  remote.set('shape', 'second');
  try {
    Y.applyUpdate(doc, Y.encodeStateAsUpdate(peer, prefix), 'relay');
    expect(doc.store.pendingStructs).not.toBeNull();
    const before = Y.encodeStateVector(doc);
    let calls = 0;
    const commit = () =>
      scope.capture(() => {
        calls++;
        doc.transact(() => map.set('shape', 'accepted'));
      }, 1);
    expect(commit).toThrow('complete room updates');
    expect(calls).toBe(0);
    expect(map.get('shape')).toBe('original');
    expect(Y.encodeStateVector(doc)).toEqual(before);
    expect(doc.getMap(REVIEW_FIELD_TARGETS).size).toBe(0);
    Y.applyUpdate(doc, Y.encodeStateAsUpdate(peer), 'relay');
    expect(doc.store.pendingStructs).toBeNull();
    commit();
    expect(calls).toBe(1);
    expect(map.get('shape')).toBe('accepted');
    expect(doc.getMap(REVIEW_FIELD_INTENTS).size).toBe(1);
  } finally {
    scope.dispose();
    doc.destroy();
    peer.destroy();
  }
});

test('field ownership follows a nested map through native deletion undo and redo', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewFields(doc),
    root = doc.getMap<Y.Map<string>>('native'),
    original = new Y.Map<string>();
  original.set('shape', 'original');
  root.set('node', original);
  scope.enable();
  const history = new Y.UndoManager(root, { captureTimeout: 0 });
  try {
    original.set('shape', 'accepted');
    root.delete('node');
    expect(history.undo()).not.toBeNull();
    const restored = root.get('node')!;
    expect(restored).not.toBe(original);
    expect(restored.get('shape')).toBe('accepted');
    expect(history.undo()).not.toBeNull();
    expect(restored.get('shape')).toBe('original');
    expect(history.redo()).not.toBeNull();
    expect(restored.get('shape')).toBe('accepted');
    restored.set('shape', 'manual');
    expect(doc.getMap(REVIEW_FIELD_TARGETS).size).toBe(2);
    expect(history.undo()).not.toBeNull();
    expect(restored.get('shape')).toBe('accepted');
  } finally {
    history.destroy();
    scope.dispose();
    doc.destroy();
  }
});

test('invalid limits and malformed followup records refuse explicitly and release failed attachment', () => {
  const doc = new Y.Doc(),
    errors: Error[] = [];
  try {
    for (const targets of [-1, 0.5, 50001, Number.NaN])
      expect(() => attachOwnedReviewFields(doc, { limits: { targets } })).toThrow('limits');
    const scope = attachOwnedReviewFields(doc, { onError: (error) => errors.push(error) });
    try {
      scope.enable();
      expect(() => doc.getMap(REVIEW_FIELD_TARGETS).set('protocol', '2')).toThrow('protocol');
      expect(errors).toHaveLength(1);
      let calls = 0;
      expect(() =>
        scope.capture(() => {
          calls++;
        }, 0)
      ).toThrow('protocol');
      expect(calls).toBe(0);
      expect(() => scope.assertHealthy()).toThrow('protocol');
    } finally {
      scope.dispose();
    }
  } finally {
    doc.destroy();
  }
});

test('sequence metadata requires exact bounded neighbor addresses and a removal intent', () => {
  const target = {
    sequence: true as const,
    root: 'native',
    key: '41:0',
    baseline: 'A',
    initiallyPresent: true,
    position: { left: null, right: '41:1' },
  };
  expect(decodeFieldTarget(fieldAddress(target), encodeFieldRecord(target))).toEqual(target);
  for (const record of [
    { ...target, baseline: null },
    { ...target, position: { left: null } },
    { ...target, position: { left: 0, right: null } },
    { ...target, position: { left: 'not-an-item', right: null } },
    { ...target, position: { left: null, right: null, extra: true } },
    { ...target, sequence: false },
  ])
    expect(() => decodeFieldTarget(fieldAddress(target), JSON.stringify(record))).toThrow();
  const doc = new Y.Doc();
  try {
    doc.getArray('native').insert(0, ['A']);
    doc.getMap(REVIEW_FIELD_TARGETS).set('protocol', '1');
    doc.getMap(REVIEW_FIELD_TARGETS).set(fieldAddress(target), encodeFieldRecord(target));
    doc.getMap(REVIEW_FIELD_INTENTS).set(
      '41:2',
      encodeFieldRecord({
        target: fieldAddress(target),
        value: 'write',
        stamp: [1, 41, 2],
      })
    );
    const vector = Y.encodeStateVector(doc);
    expect(() => attachOwnedReviewFields(doc)).toThrow('sequence intent');
    expect(Y.encodeStateVector(doc)).toEqual(vector);
    expect(doc.getArray('native').toArray()).toEqual(['A']);
  } finally {
    doc.destroy();
  }
});
test('native history ledgers stay outside field ownership', () => {
  const doc = new Y.Doc(),
    fields = attachOwnedReviewFields(doc);
  const ledgers = [TEXT_DELETIONS_KEY, TEXT_MARKS_KEY].map((name) => doc.getMap<string>(name));
  try {
    for (const ledger of ledgers) ledger.set('k', 'a');
    doc.getMap<string>('native').set('shape', 'original');
    fields.enable();
    // An undo records the characters it hides and its redo withdraws them, outside any review.
    for (const ledger of ledgers) {
      ledger.set('k', 'b');
      ledger.delete('k');
    }
    doc.getMap<string>('native').set('shape', 'changed');
    const targets = doc.getMap<string>(REVIEW_FIELD_TARGETS);
    expect(
      [...targets.keys()]
        .filter((key) => key !== 'protocol')
        .map((key) => decodeFieldTarget(key, targets.get(key)))
    ).toEqual([expect.objectContaining({ root: 'native', key: 'shape' })]);
    fields.assertHealthy();
  } finally {
    fields.dispose();
    doc.destroy();
  }
});
