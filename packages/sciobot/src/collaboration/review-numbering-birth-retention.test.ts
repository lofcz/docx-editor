import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { attachOwnedReviewFields } from './review-fields';
import { decodeFieldIntent, decodeFieldTarget, REVIEW_FIELD_TARGETS } from './review-field-codec';

for (const [protocol, repair, namespace, capturedCount] of [
  ['1', 14, 'w', 0],
  ['2', 14, 'w', 2],
  ['2', 13, 'w', 0],
  ['2', 15, 'w', 0],
  ['2', 14, 'foreign', 0],
] as const)
  test(`constructor child ownership preserves native creation history and cold state: ${protocol}/${repair}/${namespace}`, () => {
    const doc = new Y.Doc(),
      meta = doc.getMap('docx-package-meta-v1');
    for (const [key, value] of Object.entries({
      protocolVersion: 1,
      sharedSchemaVersion: 9,
      repairVersion: repair,
      canonicalModelVersion: 1,
    }))
      meta.set(key, value);
    const nodes = doc.getMap<Y.Map<unknown>>('docx-package-nodes-v1');
    const fields = attachOwnedReviewFields(doc, { protocol });
    fields.enable();
    const history = new Y.UndoManager(nodes, { captureTimeout: 0 });
    let captured: ReturnType<typeof fields.currentReceipt> = [];
    const contents = (source: Y.Doc) =>
      (
        source
          .getMap<Y.Map<unknown>>('docx-package-nodes-v1')
          .get('properties')!
          .get('children') as Y.Array<string>
      ).toArray();
    try {
      fields.capture(
        () =>
          doc.transact(() => {
            const props = new Y.Map<unknown>();
            props.set('b', ['paragraphProperties', namespace, 'pPr', 'w'].join('\u001f'));
            props.set('children', new Y.Array<string>());
            nodes.set('properties', props);
            (props.get('children') as Y.Array<string>).insert(0, ['numbering', 'record']);
          }),
        8,
        (items) => {
          captured = items;
        }
      );
      expect(captured).toHaveLength(capturedCount);
      for (const entry of captured) {
        const intent = decodeFieldIntent(entry.id, entry.record);
        const target = decodeFieldTarget(
          intent.target,
          doc.getMap(REVIEW_FIELD_TARGETS).get(intent.target)
        );
        expect(target.sequence).toBe(true);
        expect(target.initiallyPresent).toBe(false);
      }
      expect(contents(doc)).toEqual(['numbering', 'record']);
      expect(history.undo()).not.toBeNull();
      expect(nodes.has('properties')).toBe(false);
      expect(history.redo()).not.toBeNull();
      expect(contents(doc)).toEqual(['numbering', 'record']);
      const cold = new Y.Doc(),
        reader = attachOwnedReviewFields(cold, { protocol });
      try {
        Y.applyUpdate(cold, Y.encodeStateAsUpdate(doc));
        reader.enable();
        expect(contents(cold)).toEqual(['numbering', 'record']);
      } finally {
        reader.dispose();
        cold.destroy();
      }
    } finally {
      history.destroy();
      fields.dispose();
      doc.destroy();
    }
  });
