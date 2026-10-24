import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createPropertyGroupBindings } from './property-group-bindings';
import {
  decodePropertyGroupBinding,
  encodePropertyGroupBinding,
  PROPERTY_GROUP_RECEIPTS,
  MAX_PROPERTY_GROUPS,
} from './property-group-codec';
import { attachOwnedReviewFields } from './review-fields';
import { attachOwnedReviewFlags } from './review-flags';
import { propertyScopedReceipt } from './property-review-scope';
import {
  REVIEW_FIELD_INTENTS,
  REVIEW_FIELD_TARGETS,
  decodeFieldIntent,
  decodeFieldTarget,
  MAX_REVIEW_FIELD_BYTES,
} from './review-field-codec';
const id = 'g1:stored:rg_scope';
const membersKey = JSON.stringify([JSON.stringify(['/word/document.xml', '1', 'Alice', null])]);
const valid = {
  schemaVersion: 1 as const,
  groupId: id,
  membersKey,
  fields: ['41:2'],
  flags: ['41:3'],
};

test('group binding codec enforces exact schema, bounded IDs, unique ownership, and nonempty receipts', () => {
  const raw = encodePropertyGroupBinding(valid);
  const decoded = decodePropertyGroupBinding(id, raw);
  expect(decoded).toEqual(valid);
  expect(Object.isFrozen(decoded)).toBe(true);
  expect(Object.isFrozen(decoded.fields)).toBe(true);
  for (const value of [
    null,
    [],
    { ...valid, schemaVersion: 2 },
    { ...valid, extra: true },
    { ...valid, groupId: 'other' },
    { ...valid, membersKey: '' },
    { ...valid, fields: [], flags: [] },
    { ...valid, fields: ['41:2', '41:2'] },
    { ...valid, flags: ['41:3', '41:3'] },
    { ...valid, fields: ['bad'] },
    { ...valid, flags: ['-1:3'] },
    { ...valid, flags: ['41:9007199254740992'] },
  ])
    expect(() => decodePropertyGroupBinding(id, JSON.stringify(value))).toThrow();
  expect(() => decodePropertyGroupBinding(id, ' '.repeat(MAX_REVIEW_FIELD_BYTES + 1))).toThrow(
    'bounded'
  );
});

test('group slot reservations retain undo capacity and refuse before room mutation', () => {
  const doc = new Y.Doc(),
    store = createPropertyGroupBindings(doc);
  const map = doc.getMap(PROPERTY_GROUP_RECEIPTS);
  try {
    doc.transact(() => {
      for (let index = 0; index < MAX_PROPERTY_GROUPS; index++) {
        const groupId = `g1:stored:rg_${index}`;
        map.set(groupId, encodePropertyGroupBinding({ ...valid, groupId }));
      }
    });
    const before = Y.encodeStateAsUpdate(doc);
    expect(() =>
      store.reserve([{ groupId: id, membersKey, nodeIds: new Set(['pPr']) }], 1, 1)
    ).toThrow('record limit');
    expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
    map.delete('g1:stored:rg_0');
    const removed = Y.encodeStateAsUpdate(doc);
    expect(() =>
      store.reserve([{ groupId: id, membersKey, nodeIds: new Set(['pPr']) }], 1, 1)
    ).toThrow('record limit');
    expect(Y.encodeStateAsUpdate(doc)).toEqual(removed);
  } finally {
    store.dispose();
    doc.destroy();
  }
});

test('property scope captures exact scalar, namespace, and child identities after native ID allocation', () => {
  const doc = new Y.Doc(),
    fields = attachOwnedReviewFields(doc, { protocol: '2' }),
    flags = attachOwnedReviewFlags(doc, { protocol: '3' }),
    store = createPropertyGroupBindings(doc);
  const meta = doc.getMap('docx-package-meta-v1');
  meta.set('protocolVersion', 1);
  meta.set('sharedSchemaVersion', 9);
  meta.set('repairVersion', 14);
  meta.set('canonicalModelVersion', 1);
  const nodes = doc.getMap<Y.Map<unknown>>('docx-package-nodes-v1');
  const parent = new Y.Map<unknown>(),
    child = new Y.Map<unknown>();
  parent.set('children', new Y.Array<string>());
  nodes.set('stable-pPr', parent);
  nodes.set('lid:replica:1', child);
  (parent.get('children') as Y.Array<string>).insert(0, ['lid:replica:1']);
  const attrs = doc.getMap<string>('docx-package-attributes-v1');
  const namespace = doc.getMap<string>('docx-package-bindings-v1');
  const key = 'lid:replica:1\u001fw\u001fval',
    bindingKey = 'lid:replica:1\u001fprefix';
  attrs.set(key, 'w\u001f4');
  namespace.set(bindingKey, 'old');
  attrs.set('unrelated\u001fw\u001fval', 'w\u001funtouched');
  fields.enable();
  const history = new Y.UndoManager([attrs, namespace, nodes], { captureTimeout: 0 });
  try {
    let captured: ReturnType<typeof fields.currentReceipt> = [];
    fields.capture(
      () =>
        doc.transact(() => {
          attrs.set(key, 'w\u001f1');
          namespace.set(bindingKey, 'new');
          attrs.set('unrelated\u001fw\u001fval', 'w\u001fpeer');
        }),
      3,
      (items) => {
        captured = items;
      }
    );
    const selected = propertyScopedReceipt(
      doc,
      { fields: captured, flags: [] },
      new Set(['stable-pPr', 'temporary#new:1'])
    );
    expect(selected.fields).toHaveLength(2);
    const scope = { groupId: id, membersKey, nodeIds: new Set(['stable-pPr']) };
    store.reserve([scope], 3, 0);
    doc.transact(() => {
      attrs.set('decision', 'native');
      store.publish([scope], { fields: captured, flags: [] });
    });
    const binding = decodePropertyGroupBinding(id, doc.getMap(PROPERTY_GROUP_RECEIPTS).get(id));
    expect(binding.fields).toHaveLength(2);
    for (const fieldId of binding.fields) {
      const intent = decodeFieldIntent(fieldId, doc.getMap(REVIEW_FIELD_INTENTS).get(fieldId));
      const target = decodeFieldTarget(
        intent.target,
        doc.getMap(REVIEW_FIELD_TARGETS).get(intent.target)
      );
      expect([key, bindingKey]).toContain(target.key);
    }
    const beforeRead = Y.encodeStateAsUpdate(doc);
    expect(store.read(id, membersKey)).toEqual(selected);
    expect(Y.encodeStateAsUpdate(doc)).toEqual(beforeRead);
    expect(() => store.read(id, 'changed members')).toThrow('membership changed');
    meta.set('repairVersion', 999);
    const unknownProfile = Y.encodeStateAsUpdate(doc);
    expect(() => store.reserve([{ ...scope, groupId: 'g1:stored:rg_new' }], 0, 0)).toThrow(
      'profile'
    );
    expect(() => store.read(id, membersKey)).toThrow('profile');
    expect(Y.encodeStateAsUpdate(doc)).toEqual(unknownProfile);
    meta.set('repairVersion', 14);
    expect(() =>
      doc.transact(() => store.publish([scope], { fields: captured, flags: [] }))
    ).toThrow('another group');
    expect(history.undo()).not.toBeNull();
    expect(doc.getMap(PROPERTY_GROUP_RECEIPTS).has(id)).toBe(false);
    expect(history.redo()).not.toBeNull();
    expect(store.read(id, membersKey)).toEqual(selected);
  } finally {
    history.destroy();
    fields.dispose();
    flags.dispose();
    store.dispose();
    doc.destroy();
  }
});
