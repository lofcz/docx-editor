import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { attachOwnedReviewFields } from './review-fields';
import { propertyScopedReceipt } from './property-review-scope';
import { decodeFieldIntent, decodeFieldTarget, REVIEW_FIELD_TARGETS } from './review-field-codec';

test('new property anchors exclude old tombstones, paragraph text, and unrelated scalar writes', () => {
  const doc = new Y.Doc();
  const meta = doc.getMap('docx-package-meta-v1');
  for (const [key, value] of Object.entries({
    protocolVersion: 1,
    sharedSchemaVersion: 9,
    repairVersion: 14,
    canonicalModelVersion: 1,
  }))
    meta.set(key, value);
  const nodes = doc.getMap<Y.Map<unknown>>('docx-package-nodes-v1');
  const node = (name: string) => {
    const record = new Y.Map<unknown>();
    record.set('s', ['generic', 'w', name, ''].join('\u001f'));
    record.set('children', new Y.Array<string>());
    return record;
  };
  const parent = node('p'),
    old = node('pPr');
  old.set('deleted', true);
  nodes.set('paragraph', parent);
  nodes.set('obsolete-properties', old);
  const children = parent.get('children') as Y.Array<string>;
  children.insert(0, ['obsolete-properties']);
  const attrs = doc.getMap<string>('docx-package-attributes-v1');
  attrs.set('unrelated\u001fw\u001fval', 'old');
  const fields = attachOwnedReviewFields(doc, { protocol: '2' });
  fields.enable();
  try {
    let captured: ReturnType<typeof fields.currentReceipt> = [];
    fields.capture(
      () =>
        doc.transact(() => {
          nodes.set('allocated-properties', node('pPr'));
          nodes.set('allocated-run', node('r'));
          children.insert(0, ['allocated-properties', 'allocated-run']);
          attrs.set('unrelated\u001fw\u001fval', 'peer');
        }),
      8,
      (items) => {
        captured = items;
      }
    );
    const receipt = { fields: captured, flags: [] };
    const selected = propertyScopedReceipt(doc, receipt, new Set(['temporary:new-pPr']), [
      { parentId: 'paragraph', childName: 'pPr' },
    ]);
    expect(selected.fields).toHaveLength(1);
    const intent = decodeFieldIntent(selected.fields[0]!.id, selected.fields[0]!.record);
    const target = decodeFieldTarget(
      intent.target,
      doc.getMap(REVIEW_FIELD_TARGETS).get(intent.target)
    );
    expect(target.baseline).toBe('allocated-properties');
    old.set('deleted', false);
    const before = Y.encodeStateAsUpdate(doc);
    expect(() =>
      propertyScopedReceipt(doc, receipt, new Set(['temporary:new-pPr']), [
        { parentId: 'paragraph', childName: 'pPr' },
      ])
    ).toThrow('Ambiguous');
    expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
  } finally {
    fields.dispose();
    doc.destroy();
  }
});
