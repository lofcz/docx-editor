// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { checklistParagraph, listNumbering } from '../../test/support/list-document';
import { checklistIndentOracle } from '../../test/support/checklist-indent-oracle';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { acceptReviewGroup } from '../review/accept-group';
import { rejectReviewGroup } from '../review/reject-group';
import { readEditorRevision } from '../projection/revision-cache';
import { indentManualChecklist, manualListState } from './lists';
const context = { documentId: 'checklist-indent-shared', epoch: 'mount' };
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const prior of [false, true])
    for (const decision of ['direct', 'accept', 'reject'] as const)
      test(`captured checklist action retains concurrent text, histories and native Word output: ${clients}/${prior}/${decision}`, async () => {
        const pair = await createTwoEditorFixture({
          document: docxFixture(
            prior
              ? checklistParagraph(false).replace('Task', 'Alpha beta')
              : '<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p>',
            { numbering: listNumbering }
          ),
          documentId: context.documentId,
          clientIds: clients,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors,
            a = alice!.surface!,
            b = bob!.surface!;
          const id = a.session.paragraphIds()[0]!,
            original = a.session.part();
          const point = { paragraphId: id, offset: prior ? 6 : 5 };
          a.setSelection({ anchor: point, head: point });
          b.setSelection({ anchor: point, head: point });
          const op = {
            op: 'insertText' as const,
            paragraphId: id,
            offset: point.offset,
            text: ' peer',
          };
          const originalPeer = applyTreeOp(original, op);
          if (!originalPeer.ok) throw new Error(originalPeer.reason);
          if (decision !== 'direct')
            expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
          pair.pause();
          const changed = indentManualChecklist(alice!, context);
          if (!changed.success) throw new Error(changed.error);
          const member = manualListState(alice!, context, id)!.numbering!;
          expect(b.applyAutomationOps(() => [op]).committed).toBe(true);
          pair.reconnect(clients[0] === 41 ? 0 : 1);
          await pair.settle();
          const pending = canonicalOoxmlFingerprint(a.session.part());
          if (decision !== 'direct') {
            const groups = readReviewGroups(a.session.currentPackage());
            if (!groups.success) throw new Error(groups.error);
            expect(groups.groups).toHaveLength(1);
            const result = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
              alice!,
              groups.groups[0]!.id,
              readEditorRevision(alice!, context.documentId, context.epoch)!,
              context
            );
            if (!result.success) throw new Error(result.error);
            await pair.settle();
          }
          const expected = canonicalOoxmlFingerprint(
            checklistIndentOracle(
              original,
              id,
              prior,
              decision !== 'reject',
              member.numId,
              member.level,
              'Alpha peer beta'
            )
          );
          for (const editor of pair.editors)
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(expected);
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(
              decision === 'direct' ? canonicalOoxmlFingerprint(originalPeer.part) : pending
            );
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(expected);
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          const withoutPeer = canonicalOoxmlFingerprint(
            checklistIndentOracle(
              original,
              id,
              prior,
              decision !== 'reject',
              member.numId,
              member.level,
              'Alpha beta'
            )
          );
          for (const editor of pair.editors)
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(withoutPeer);
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
