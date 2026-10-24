// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { listDocument, checklistParagraph, listParagraph } from '../../test/support/list-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { acceptReviewGroup } from '../review/accept-group';
import { rejectReviewGroup } from '../review/reject-group';
import { readEditorRevision } from '../projection/revision-cache';
import { handleManualListKey } from './list-input';
import { manualListState } from './lists';
const context = { documentId: 'manual-lists', epoch: 'mount' };
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const split of [5, 3, 2] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`tracked checklist continuation retains offline peer text, review and independent histories ${clients}/${split}/${decision}`, async () => {
        const pair = await createTwoEditorFixture({
          clientIds: clients,
          documentId: context.documentId,
          ownedAttributes: true,
          document: listDocument(checklistParagraph(true, 1) + listParagraph('Tail')),
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors,
            a = alice!.surface!,
            b = bob!.surface!,
            id = a.session.paragraphIds()[0]!;
          const end = { paragraphId: id, offset: split },
            peer = { paragraphId: id, offset: 1 };
          a.setSelection({ anchor: end, head: split === 2 ? { ...end, offset: 4 } : end });
          b.setSelection({ anchor: peer, head: peer });
          expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
          const before = canonicalOoxmlFingerprint(a.session.part());
          pair.pause();
          expect(handleManualListKey(alice!, 'Enter', context)).toBe(true);
          const next = a.state().selection.head.paragraphId;
          b.insertPlainText('!');
          b.flushPendingInput();
          pair.reconnect(clients[0] === 41 ? 0 : 1);
          await pair.settle();
          const after = canonicalOoxmlFingerprint(a.session.part()),
            selection = a.state().selection;
          expect(canonicalOoxmlFingerprint(b.session.part())).toEqual(after);
          expect(paragraphTextOf(a.session.part(), id)).toBe(
            split === 5 ? '☒!Task' : split === 3 ? '☒!Ta' : '☒!Tas'
          );
          expect(manualListState(alice!, context, next)?.checkbox).toMatchObject({
            checked: split !== 5,
            level: 1,
          });
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(a.session.paragraphIds()).toHaveLength(2);
          expect(paragraphTextOf(a.session.part(), id)).toBe('☒!Task');
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(before);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(after);
          expect(a.state().selection).toEqual(selection);
          const groups = readReviewGroups(a.session.currentPackage());
          if (!groups.success) throw new Error(groups.error);
          expect(groups.groups).toHaveLength(1);
          const decide = decision === 'accept' ? acceptReviewGroup : rejectReviewGroup;
          expect(
            decide(
              alice!,
              groups.groups[0]!.id,
              readEditorRevision(alice!, context.documentId, context.epoch)!,
              context
            ).success
          ).toBe(true);
          await pair.settle();
          const resolved = canonicalOoxmlFingerprint(a.session.part());
          expect(canonicalOoxmlFingerprint(b.session.part())).toEqual(resolved);
          const saved = [];
          for (const editor of pair.editors) {
            const loaded = readOoxmlPackage(new Uint8Array(await editor.save()));
            if (!loaded.ok) throw new Error(loaded.reason);
            saved.push(semanticDigest(loaded.package.parts.values()));
            const fresh = createDocxEditor({
              container: document.createElement('div'),
              document: new Uint8Array(await editor.save()),
            });
            try {
              const paragraphs = fresh.surface!.session.paragraphIds();
              expect(
                paragraphs.map((p) => paragraphTextOf(fresh.surface!.session.part(), p))
              ).toEqual(
                decision === 'accept'
                  ? split === 5
                    ? ['☒!Task', '☐', 'Tail']
                    : split === 3
                      ? ['☒!Ta', '☒sk', 'Tail']
                      : ['☒!T', '☒k', 'Tail']
                  : ['☒!Task', 'Tail']
              );
              expect(manualListState(fresh, context, paragraphs[0]!)?.checkbox).toMatchObject({
                checked: true,
                level: 1,
              });
            } finally {
              fresh.destroy();
            }
          }
          expect(saved[0]).toEqual(saved[1]);
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(after);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(resolved);
        } finally {
          pair.destroy();
        }
      });
