// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { listNumbering, listParagraph } from '../../test/support/list-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { acceptReviewGroup } from '../review/accept-group';
import { rejectReviewGroup } from '../review/reject-group';
import { readEditorRevision } from '../projection/revision-cache';
import { setManualList, manualListState, manualListMarker } from './lists';
const context = { documentId: 'removed-ancestor-peers', epoch: 'mount' };
function word(part: OoxmlPart, ops: readonly TreeDocOp[]) {
  for (const op of ops) {
    const result = applyTreeOp(part, op);
    if (!result.ok) throw new Error(result.reason);
    part = result.part;
  }
  return part;
}
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const decision of ['direct', 'accept', 'reject'] as const)
    test(`nested sequence preserves parent prefixes, disconnected peer text and independent histories: ${clients}/${decision}`, async () => {
      const pair = await createTwoEditorFixture({
        document: docxFixture(
          listParagraph('Root one', '1') +
            listParagraph('Root two', '1') +
            listParagraph('Removed child', '1', 1) +
            listParagraph('Tail child', '1', 1),
          {
            numbering: listNumbering.replace(
              '<w:lvlText w:val="%2."/>',
              '<w:lvlText w:val="%1.%2."/>'
            ),
          }
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
          b = bob!.surface!,
          ids = a.session.paragraphIds(),
          point = { paragraphId: ids[2]!, offset: 2 },
          peer = { paragraphId: ids[3]!, offset: 2 };
        a.setSelection({ anchor: point, head: point });
        b.setSelection({ anchor: peer, head: peer });
        const original = a.session.part(),
          before = canonicalOoxmlFingerprint(original);
        if (decision !== 'direct')
          expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
        pair.pause();
        const result = setManualList(alice!, 'none', context);
        if (!result.success) throw new Error(result.error);
        const tail = manualListState(alice!, context, ids[3]!)!.numbering!;
        expect(ids.map((id) => manualListMarker(alice!, id)?.text ?? null)).toEqual([
          '1.',
          '2.',
          null,
          '2.1.',
        ]);
        const edits: TreeDocOp[] = [
          { op: 'setListNumbering', paragraphId: ids[2]!, numId: '0' },
          ...[3].map((index) => ({
            op: 'setListNumbering' as const,
            paragraphId: ids[index]!,
            numId: tail.numId,
            level: 1,
          })),
        ];
        const peerEdit: TreeDocOp = {
          op: 'insertText',
          paragraphId: peer.paragraphId,
          offset: 2,
          text: '!',
        };
        const expected = canonicalOoxmlFingerprint(
          word(original, [...(decision === 'reject' ? [] : edits), peerEdit])
        );
        b.insertPlainText('!');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        const pending = canonicalOoxmlFingerprint(a.session.part());
        if (decision !== 'direct') {
          const groups = readReviewGroups(a.session.currentPackage());
          if (!groups.success) throw new Error(groups.error);
          expect(groups.groups).toHaveLength(1);
          const decided = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
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
        }
        for (const editor of pair.editors)
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(expected);
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        const noPeer =
          decision === 'reject' ? before : canonicalOoxmlFingerprint(word(original, edits));
        for (const editor of pair.editors)
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(noPeer);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(expected);
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
