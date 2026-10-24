// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { listNumbering, listParagraph } from '../../test/support/list-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import {
  PROPERTY_REVIEW_PROTOTYPE,
  type PropertyReviewPrototype,
} from '../collaboration/property-review-prototype';
import { propertyScopedReceipt } from '../collaboration/property-review-scope';
import { withoutReviewGroupRecord, readReviewGroupStore } from '../review/group-record-store';
import type { OoxmlNode, OoxmlPackage } from '@docx-editor.dev/core/store';
import { setManualList } from './lists';
const context = { documentId: 'list-sequence-peers', epoch: 'mount' };
const numbering = listNumbering.replace(
  '</w:numbering>',
  '<w:num w:numId="4"><w:abstractNumId w:val="0"/></w:num>' +
    '<w:num w:numId="5"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
);
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`property ownership withdrawal keeps a peer's conflicting tail format ${clients}`, async () => {
    const pair = await createTwoEditorFixture({
      document: docxFixture(
        listParagraph('One', '1') +
          listParagraph('Gap') +
          listParagraph('Three', '4') +
          listParagraph('Four', '4'),
        { numbering }
      ),
      documentId: context.documentId,
      clientIds: clients,
      ownedAttributes: true,
      experimentalPropertyReview: true,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors,
        a = alice!.surface!,
        b = bob!.surface!;
      const ids = a.session.paragraphIds();
      const first = { paragraphId: ids[1]!, offset: 1 },
        last = { paragraphId: ids[3]!, offset: 1 };
      a.setSelection({ anchor: first, head: first });
      b.setSelection({ anchor: last, head: last });
      expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
      const original = canonicalOoxmlFingerprint(a.session.part());
      const propertyIds = new Set<string>();
      const collect = (node: OoxmlNode, active = false) => {
        if (node.kind === 'textValue') return;
        active ||= node.kind === 'paragraphProperties';
        if (active) propertyIds.add(node.id);
        for (const child of node.children) collect(child, active);
      };
      collect(a.session.part().root);
      const properties = (a.collaborationSession() as unknown as Record<symbol, unknown>)[
        PROPERTY_REVIEW_PROTOTYPE
      ] as PropertyReviewPrototype;
      pair.pause();
      const proposed = setManualList(alice!, 'decimal', context);
      if (!proposed.success) throw new Error(proposed.error);
      collect(a.session.part().root);
      const receipt = propertyScopedReceipt(pair.documents[0]!, properties.latest()!, propertyIds);
      expect(receipt.fields.length).toBeGreaterThan(0);
      expect(receipt.flags).toHaveLength(0);
      expect(setManualList(bob!, 'circle', context).success).toBe(true);
      const expected = canonicalOoxmlFingerprint(b.session.part());

      pair.reconnect(clients[0] === 41 ? 0 : 1);
      await pair.settle();
      const pending = canonicalOoxmlFingerprint(a.session.part());
      const groups = readReviewGroups(a.session.currentPackage());
      if (!groups.success) throw new Error(groups.error);
      expect(groups.groups).toHaveLength(1);
      const store = readReviewGroupStore(a.session.currentPackage());
      if (store.status !== 'ready') throw new Error('Missing group store.');
      const groupRecord = [...store.records.keys()][0]!;
      const beforeInvalid = canonicalOoxmlFingerprint(a.session.part());
      const beforeInvalidRoom = Y.encodeStateAsUpdate(pair.documents[0]!);
      let invalidCalled = false;
      expect(() =>
        properties.withdrawDuringPublication(
          { flags: [], fields: [receipt.fields[0]!, receipt.fields[0]!] },
          () => {
            invalidCalled = true;
          }
        )
      ).toThrow('stale');
      expect(invalidCalled).toBe(false);
      expect(canonicalOoxmlFingerprint(a.session.part())).toBe(beforeInvalid);
      expect(Y.encodeStateAsUpdate(pair.documents[0]!)).toEqual(beforeInvalidRoom);
      let committed = false;
      properties.withdrawDuringPublication(receipt, () => {
        const refusal = a.commitReviewOps((historyToken) => {
          const applied = a.session.applyTreeOps(
            [],
            undefined,
            undefined,
            { kind: 'body' },
            {
              historyToken,
              packageEditsAfterOps: [
                (pkg: OoxmlPackage) => {
                  const removed = withoutReviewGroupRecord(pkg, groupRecord);
                  if (!removed.ok) throw new Error(removed.reason);
                  return removed.pkg;
                },
              ],
            }
          );
          committed = applied.committed;
          return applied;
        }, 'revision-resolve');
        if (refusal) throw new Error(String(refusal));
      });
      expect(committed).toBe(true);
      await pair.settle();
      for (const peer of pair.editors)
        expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(expected);
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const peer of pair.editors)
        expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(pending);
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const peer of pair.editors)
        expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(expected);
      expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const peer of pair.editors)
        expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(original);
      expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const peer of pair.editors)
        expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(expected);
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      const cold = await createTwoEditorFixture({
        document: docxFixture(listParagraph('Unused')),
        documentId: context.documentId,
        ownedAttributes: true,
        experimentalPropertyReview: true,
        roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        await cold.settle();
        for (const peer of cold.editors) {
          expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(expected);
          const read = readReviewGroups(peer.surface!.session.currentPackage());
          if (!read.success) throw new Error(read.error);
          expect(read.groups).toHaveLength(0);
        }
        const reopened = await cold.savedDigests();
        expect(reopened[0]).toEqual(saved[0]);
        expect(reopened[1]).toEqual(saved[0]);
      } finally {
        cold.destroy();
      }
    } finally {
      pair.destroy();
    }
  });
