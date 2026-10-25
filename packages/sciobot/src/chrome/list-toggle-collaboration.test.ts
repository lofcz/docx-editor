// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  directParagraphProperties,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { listParagraph, listNumbering } from '../../test/support/list-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { rejectReviewGroup } from '../review/reject-group';
import { readEditorRevision } from '../projection/revision-cache';
import { toggleManualList } from './lists';
const context = { documentId: 'toggle-peer-indent', epoch: 'mount' };
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`nested list removal rejection keeps the observed peer indentation field and histories: ${clients}`, async () => {
    const body = listParagraph('Item', '1', 2).replace(
      '<w:customKeep',
      '<w:ind xmlns:u="urn:opaque" w:left="2160" w:hanging="360" w:right="240" u:keep="yes"/><w:customKeep'
    );
    const pair = await createTwoEditorFixture({
      document: docxFixture(body, { numbering: listNumbering }),
      documentId: context.documentId,
      clientIds: clients,
      ownedAttributes: true,
      experimentalPropertyReview: true,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors,
        a = alice!.surface!,
        b = bob!.surface!,
        id = a.session.paragraphIds()[0]!,
        point = { paragraphId: id, offset: 2 };
      a.setSelection({ anchor: point, head: point });
      b.setSelection({ anchor: point, head: point });
      const original = a.session.part(),
        before = canonicalOoxmlFingerprint(original);
      expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
      const toggled = toggleManualList(alice!, 'decimal', context);
      if (!toggled.success) throw new Error(toggled.error);
      await pair.settle();
      const properties = directParagraphProperties(b.session.part(), id).map((property) =>
        property.localName === 'ind'
          ? { ...property, attributes: { ...property.attributes, right: '480' } }
          : property
      );
      expect(
        b.applyAutomationOps(() => [{ op: 'setParagraphProperties', paragraphId: id, properties }])
          .committed
      ).toBe(true);
      const expectedProperties = directParagraphProperties(original, id).map((property) =>
        property.localName === 'ind'
          ? { ...property, attributes: { ...property.attributes, right: '480' } }
          : property
      );
      const word = applyTreeOp(original, {
        op: 'setParagraphProperties',
        paragraphId: id,
        properties: expectedProperties,
      });
      if (!word.ok) throw new Error(word.reason);
      const expected = canonicalOoxmlFingerprint(word.part);
      await pair.settle();
      const pending = canonicalOoxmlFingerprint(a.session.part());
      const groups = readReviewGroups(a.session.currentPackage());
      if (!groups.success) throw new Error(groups.error);
      expect(groups.groups).toHaveLength(1);
      // Rehydrate the exact pending native room before withdrawing the proposal.
      const cold = await createTwoEditorFixture({
        document: docxFixture('<w:p/>'),
        documentId: context.documentId,
        ownedAttributes: true,
        experimentalPropertyReview: true,
        roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const reviewer = cold.editors[0]!;
        const decided = rejectReviewGroup(
          reviewer,
          groups.groups[0]!.id,
          readEditorRevision(reviewer, context.documentId, context.epoch)!,
          context
        );
        if (!decided.success) throw new Error(decided.error);
        await cold.settle();
        for (const editor of cold.editors) {
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(expected);
          const loaded = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!loaded.ok) throw new Error(loaded.reason);
          expect(
            canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
          ).toEqual(expected);
        }
      } finally {
        cold.destroy();
      }
      const decided = rejectReviewGroup(
        alice!,
        groups.groups[0]!.id,
        readEditorRevision(alice!, context.documentId, context.epoch)!,
        context
      );
      if (!decided.success) throw new Error(decided.error);
      await pair.settle();
      for (const editor of pair.editors)
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(expected);
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(pending);
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(expected);
      expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(before);
      expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(expected);
      const ownCaret = caretAt(a.layout(), point),
        peerCaret = caretAt(b.layout(), point);
      expect(ownCaret).not.toBeNull();
      expect(peerCaret).not.toBeNull();
      expect({ x: ownCaret!.x, y: ownCaret!.y }).toEqual({ x: peerCaret!.x, y: peerCaret!.y });
      const remote = pair.rooms[0]!.session.remoteSelections().find(
        (selection) => selection.actorId === 'bob'
      );
      expect(remote?.head.nodeId).toBe(id);
      expect(remote?.head.offset).toBe(point.offset);
      expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(0);
      for (const editor of pair.editors) {
        const loaded = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!loaded.ok) throw new Error(loaded.reason);
        expect(
          canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
        ).toEqual(expected);
      }
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
