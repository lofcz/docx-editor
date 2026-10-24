// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  applyTreeOp,
  readOoxmlPackage,
  canonicalOoxmlFingerprint,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { listNumbering, listParagraph } from '../../test/support/list-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { rejectReviewGroup } from '../review/reject-group';
import { acceptReviewGroup } from '../review/accept-group';
import { readEditorRevision } from '../projection/revision-cache';
import { manualListState, setManualList } from './lists';
const context = { documentId: 'property-child-births', epoch: 'mount' };
async function expectSavedMain(
  pair: Awaited<ReturnType<typeof createTwoEditorFixture>>,
  expected: string
) {
  for (const editor of pair.editors) {
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(
      canonicalOoxmlFingerprint(reopened.package.parts.get(reopened.package.mainDocumentPart)!)
    ).toEqual(expected);
  }
}

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const source of ['unlisting', 'new-child', 'new-container'] as const)
    for (const reopen of [false, true])
      test(`property birth withdrawal keeps peer numbering and both histories: ${clients}/${source}/cold=${reopen}`, async () => {
        const pair = await createTwoEditorFixture({
          document: docxFixture(
            source === 'new-container'
              ? '<w:p><w:r><w:t>Item</w:t></w:r></w:p>'
              : listParagraph('Item', source === 'unlisting' ? '1' : undefined),
            { numbering: listNumbering }
          ),
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
          const id = a.session.paragraphIds()[0]!;
          const point = { paragraphId: id, offset: 2 };
          a.setSelection({ anchor: point, head: point });
          b.setSelection({ anchor: point, head: point });
          const originalPart = a.session.part(),
            original = canonicalOoxmlFingerprint(originalPart);
          expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
          const removed = setManualList(
            alice!,
            source === 'unlisting' ? 'none' : 'decimal',
            context
          );
          if (!removed.success) throw new Error(removed.error);
          await pair.settle();
          if (source === 'unlisting')
            expect(manualListState(bob!, context, id)!.numbering).toBeNull();
          else expect(manualListState(bob!, context, id)!.numbering).not.toBeNull();
          const added = setManualList(bob!, 'circle', context);
          if (!added.success) throw new Error(added.error);
          const membership = manualListState(bob!, context, id)!.numbering!;
          const word = applyTreeOp(originalPart, {
            op: 'setListNumbering',
            paragraphId: id,
            numId: membership.numId,
            level: membership.level,
          });
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
              document: docxFixture(listParagraph('Unused')),
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
          const result = rejectReviewGroup(
            reviewer,
            groups.groups[0]!.id,
            readEditorRevision(reviewer, context.documentId, context.epoch)!,
            context
          );
          if (!result.success) throw new Error(result.error);
          await active.settle();
          for (const peer of active.editors)
            expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(expected);
          expect(reviewer.exec({ type: 'undo' }).ok).toBe(true);
          await active.settle();
          for (const peer of active.editors)
            expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(pending);
          expect(reviewer.exec({ type: 'redo' }).ok).toBe(true);
          await active.settle();
          for (const peer of active.editors)
            expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(expected);
          if (!reopen) {
            expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
            await active.settle();
            for (const peer of active.editors)
              expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(original);
            expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
            await active.settle();
            for (const peer of active.editors)
              expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(expected);
          }
          await expectSavedMain(active, expected);
          const saved = await active.savedDigests();
          expect(saved[0]).toEqual(saved[1]);
          const cold = await createTwoEditorFixture({
            document: docxFixture(listParagraph('Unused')),
            documentId: context.documentId,
            ownedAttributes: true,
            experimentalPropertyReview: true,
            roomSnapshot: Y.encodeStateAsUpdate(active.documents[0]!),
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          });
          try {
            await cold.settle();
            for (const peer of cold.editors)
              expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(expected);
            const reopened = await cold.savedDigests();
            expect(reopened[0]).toEqual(saved[0]);
            expect(reopened[1]).toEqual(saved[0]);
          } finally {
            cold.destroy();
          }
        } finally {
          pendingRoom?.destroy();
          pair.destroy();
        }
      });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const holder of [false, true])
    for (const peerText of [false, true])
      for (const decision of ['accept', 'reject'] as const)
        test(`numbering birth decisions retain exact Word state and unrelated peer text: ${clients}/holder=${holder}/text=${peerText}/${decision}`, async () => {
          const pair = await createTwoEditorFixture({
            document: docxFixture(
              holder ? listParagraph('Item') : '<w:p><w:r><w:t>Item</w:t></w:r></w:p>',
              { numbering: listNumbering }
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
            const id = a.session.paragraphIds()[0]!,
              point = { paragraphId: id, offset: 2 };
            a.setSelection({ anchor: point, head: point });
            b.setSelection({ anchor: point, head: point });
            const original = a.session.part();
            let expected = original;
            expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
            const changed = setManualList(alice!, 'decimal', context);
            if (!changed.success) throw new Error(changed.error);
            await pair.settle();
            if (peerText) {
              const op = { op: 'insertText' as const, paragraphId: id, offset: 2, text: ' peer' };
              expect(b.applyAutomationOps(() => [op]).committed).toBe(true);
              const result = applyTreeOp(original, op);
              if (!result.ok) throw new Error(result.reason);
              expected = result.part;
              await pair.settle();
            }
            const groups = readReviewGroups(a.session.currentPackage());
            if (!groups.success) throw new Error(groups.error);
            expect(groups.groups).toHaveLength(1);
            const pending = canonicalOoxmlFingerprint(a.session.part());
            if (decision === 'accept') {
              const accepted = applyTreeOp(a.session.part(), {
                op: 'acceptAllRevisions',
                siteNodeIds: groups.groups[0]!.native.flatMap((member) => member.siteNodeIds),
              });
              if (!accepted.ok) throw new Error(accepted.reason);
              expected = accepted.part;
            }
            const result = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
              alice!,
              groups.groups[0]!.id,
              readEditorRevision(alice!, context.documentId, context.epoch)!,
              context
            );
            if (!result.success) throw new Error(result.error);
            await pair.settle();
            for (const editor of pair.editors)
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(
                canonicalOoxmlFingerprint(expected)
              );
            expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors)
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(pending);
            expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors)
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(
                canonicalOoxmlFingerprint(expected)
              );
            await expectSavedMain(pair, canonicalOoxmlFingerprint(expected));
            const saved = await pair.savedDigests();
            expect(saved[0]).toEqual(saved[1]);
          } finally {
            pair.destroy();
          }
        });
