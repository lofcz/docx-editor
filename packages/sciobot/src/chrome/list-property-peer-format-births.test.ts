// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { listNumbering } from '../../test/support/list-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { rejectReviewGroup } from '../review/reject-group';
import { readEditorRevision } from '../projection/revision-cache';
import { setManualList } from './lists';
const context = { documentId: 'peer-format-births', epoch: 'mount' };
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const mark of [false, true])
    for (const reopen of [false, true])
      test(`new list-property ancestry preserves independent peer formatting: ${clients}/mark=${mark}/cold=${reopen}`, async () => {
        const pair = await createTwoEditorFixture({
          document: docxFixture('<w:p><w:r><w:t>Item</w:t></w:r></w:p>', {
            numbering: listNumbering,
          }),
          documentId: context.documentId,
          clientIds: clients,
          ownedAttributes: true,
          experimentalPropertyReview: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        let pendingRoom: Awaited<ReturnType<typeof createTwoEditorFixture>> | undefined;
        try {
          const [alice, bob] = pair.editors,
            a = alice!.surface!,
            b = bob!.surface!;
          const id = a.session.paragraphIds()[0]!,
            point = { paragraphId: id, offset: 2 };
          a.setSelection({ anchor: point, head: point });
          b.setSelection({ anchor: point, head: point });
          const originalPart = a.session.part(),
            original = canonicalOoxmlFingerprint(originalPart);
          expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
          const changed = setManualList(alice!, 'decimal', context);
          if (!changed.success) throw new Error(changed.error);
          await pair.settle();
          const op: TreeDocOp = mark
            ? {
                op: 'setParagraphMarkProperties',
                paragraphId: id,
                properties: [{ localName: 'b', attributes: {} }],
              }
            : {
                op: 'setParagraphProperties',
                paragraphId: id,
                properties: [{ localName: 'jc', attributes: { val: 'center' } }],
              };
          expect(b.applyAutomationOps(() => [op]).committed).toBe(true);
          const word = applyTreeOp(originalPart, op);
          if (!word.ok) throw new Error(word.reason);
          const expected = canonicalOoxmlFingerprint(word.part);
          await pair.settle();
          const pending = canonicalOoxmlFingerprint(a.session.part());
          const groups = readReviewGroups(a.session.currentPackage());
          if (!groups.success) throw new Error(groups.error);
          expect(groups.groups).toHaveLength(1);
          let active = pair;
          if (reopen) {
            pendingRoom = await createTwoEditorFixture({
              document: docxFixture('<w:p/>'),
              documentId: context.documentId,
              ownedAttributes: true,
              experimentalPropertyReview: true,
              roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
              editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
            });
            active = pendingRoom;
            await active.settle();
          }
          const reviewer = active.editors[0]!;
          const rejected = rejectReviewGroup(
            reviewer,
            groups.groups[0]!.id,
            readEditorRevision(reviewer, context.documentId, context.epoch)!,
            context
          );
          if (!rejected.success) throw new Error(rejected.error);
          await active.settle();
          for (const editor of active.editors)
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(expected);
          expect(reviewer.exec({ type: 'undo' }).ok).toBe(true);
          await active.settle();
          for (const editor of active.editors)
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(pending);
          expect(reviewer.exec({ type: 'redo' }).ok).toBe(true);
          await active.settle();
          for (const editor of active.editors)
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(expected);
          if (!reopen) {
            const own = caretAt(a.layout(), point),
              peer = caretAt(b.layout(), point);
            expect(own).not.toBeNull();
            expect(peer).not.toBeNull();
            expect({ x: own!.x, y: own!.y }).toEqual({ x: peer!.x, y: peer!.y });
            const remote = pair.rooms[0]!.session.remoteSelections().find(
              (selection) => selection.actorId === 'bob'
            );
            expect(remote?.head.nodeId).toBe(id);
            expect(remote?.head.offset).toBe(2);
            expect(
              pair.containers[0]!.querySelectorAll('.docx-remote-caret').length
            ).toBeGreaterThan(0);
            expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
            await active.settle();
            for (const editor of active.editors)
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(original);
            expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
            await active.settle();
          }
          for (const editor of active.editors) {
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(expected);
            const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
            if (!reopened.ok) throw new Error(reopened.reason);
            expect(
              canonicalOoxmlFingerprint(
                reopened.package.parts.get(reopened.package.mainDocumentPart)!
              )
            ).toEqual(expected);
          }
          const saved = await active.savedDigests();
          expect(saved[0]).toEqual(saved[1]);
        } finally {
          pendingRoom?.destroy();
          pair.destroy();
        }
      });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`peer removal and its repeated history do not retain an unused list constructor: ${clients}`, async () => {
    const pair = await createTwoEditorFixture({
      document: docxFixture('<w:p><w:r><w:t>Item</w:t></w:r></w:p>', { numbering: listNumbering }),
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
      const id = a.session.paragraphIds()[0]!,
        point = { paragraphId: id, offset: 2 };
      a.setSelection({ anchor: point, head: point });
      b.setSelection({ anchor: point, head: point });
      const originalPart = a.session.part(),
        original = canonicalOoxmlFingerprint(originalPart);
      expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
      const changed = setManualList(alice!, 'decimal', context);
      if (!changed.success) throw new Error(changed.error);
      await pair.settle();
      const op: TreeDocOp = {
        op: 'setParagraphProperties',
        paragraphId: id,
        properties: [{ localName: 'jc', attributes: { val: 'center' } }],
      };
      const word = applyTreeOp(originalPart, op);
      if (!word.ok) throw new Error(word.reason);
      const formatted = canonicalOoxmlFingerprint(word.part);
      expect(b.applyAutomationOps(() => [op], undefined, undefined, false, true).committed).toBe(
        true
      );
      await pair.settle();
      expect(
        b.applyAutomationOps(
          () => [{ op: 'setParagraphProperties', paragraphId: id, properties: [] }],
          undefined,
          undefined,
          false,
          true
        ).committed
      ).toBe(true);
      await pair.settle();
      const groups = readReviewGroups(a.session.currentPackage());
      if (!groups.success) throw new Error(groups.error);
      expect(groups.groups).toHaveLength(1);
      const rejected = rejectReviewGroup(
        alice!,
        groups.groups[0]!.id,
        readEditorRevision(alice!, context.documentId, context.epoch)!,
        context
      );
      if (!rejected.success) throw new Error(rejected.error);
      await pair.settle();
      for (const editor of pair.editors)
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(original);
      for (const [command, expected] of [
        ['undo', formatted],
        ['undo', original],
        ['redo', formatted],
        ['redo', original],
      ] as const) {
        expect(bob!.exec({ type: command }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(expected);
        }
      }
    } finally {
      pair.destroy();
    }
  });
