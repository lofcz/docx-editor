import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { nativeBirthFieldCreation } from './review-birth-field-retention';
import { nativeNumberingChildArray } from './review-numbering-birth-retention';
import { retainedBirthSequenceEntry } from './review-sequence-retention';

for (const schema of [9, 10, 11, '10'] as const)
  test(`native birth retention uses the evaluated schema profile: ${typeof schema}/${schema}`, () => {
    const doc = new Y.Doc();
    const meta = doc.getMap('docx-package-meta-v1');
    const nodes = doc.getMap<Y.Map<unknown>>('docx-package-nodes-v1');
    for (const [key, value] of Object.entries({
      protocolVersion: 1,
      sharedSchemaVersion: schema,
      repairVersion: 17,
      canonicalModelVersion: 1,
    }))
      meta.set(key, value);
    const properties = new Y.Map<unknown>();
    properties.set('b', ['paragraphProperties', 'w', 'pPr', 'w'].join('\u001f'));
    const children = new Y.Array<unknown>();
    properties.set('children', children);
    nodes.set('properties', properties);
    const paragraph = new Y.Map<unknown>();
    paragraph.set('undone', true);
    nodes.set('paragraph', paragraph);
    children.push(['paragraph']);
    const supported = schema === 10;
    let checked = false;
    doc.once('afterTransaction', (transaction) => {
      const before = Y.encodeStateAsUpdate(doc);
      expect(
        nativeBirthFieldCreation(doc, transaction, {
          root: 'docx-package-attributes-v1',
          key: ['width', 'w', 'w'].join('\u001f'),
        })
      ).toBe(supported);
      expect(nativeNumberingChildArray(doc, children)).toBe(supported);
      expect(retainedBirthSequenceEntry(doc, children, 'paragraph')).toBe(supported);
      expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
      checked = true;
    });
    try {
      doc.transact(() => {
        const width = new Y.Map<unknown>();
        width.set('b', ['generic', 'w', 'tcW', 'w'].join('\u001f'));
        nodes.set('width', width);
      });
      expect(checked).toBe(true);
    } finally {
      doc.destroy();
    }
  });
