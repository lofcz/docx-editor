import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { paragraphPlan, TEXT_DOCUMENT_ID, TEXT_EPOCH } from '../../test/support/text-plan';
import { readEditorRevision } from '../projection/revision-cache';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTextPlan } from './track-text';
const context = {
  documentId: TEXT_DOCUMENT_ID,
  epoch: TEXT_EPOCH,
  author: 'ScioBot',
  date: '2026-10-19T09:00:00Z',
};
async function views(editor: DocxEditorInstance) {
  const revision = readEditorRevision(editor, TEXT_DOCUMENT_ID, TEXT_EPOCH)!;
  const original = await readCommittedView(editor.surface!.session, revision),
    working = await readWorkingView(editor.surface!.session, revision);
  if (!original.success || !working.success) throw new Error('Missing review view');
  return [original.view.projection.text, working.view.projection.text];
}
for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    for (const action of ['acceptReviewItem', 'rejectReviewItem'] as const)
      test(`whole body replacement preserves unseen suffix through ${action}; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
        const pair = await createTwoEditorFixture({
          document: docxFixture(
            '<w:p><w:r><w:t>Two</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
          ),
          ownedAttributes: true,
          clientIds,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors,
            id = alice!.surface!.session.paragraphIds()[0]!;
          const plan = paragraphPlan(alice!, id, 'Changed');
          pair.pause();
          expect(trackTextPlan(alice!, plan, context).success).toBe(true);
          expect(
            bob!.surface!.applyAutomationOps(() => [
              { op: 'insertText', paragraphId: id, offset: 3, text: ' Bob' },
            ]).rejected
          ).toBe(false);
          pair.reconnect(first);
          await pair.settle();
          const pending = ['Two Bob\n\nTail', 'Changed Bob\n\nTail'];
          for (const editor of pair.editors) {
            expect(await views(editor)).toEqual(pending);
            expect(editor.getReviewItems()).toHaveLength(1);
          }
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(await views(editor)).toEqual([pending[0], pending[0]]);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(await views(editor)).toEqual(pending);
          expect(bob![action](bob!.getReviewItems()[0]!.key).ok).toBe(true);
          await pair.settle();
          const resolved = action === 'acceptReviewItem' ? pending[1] : pending[0];
          for (const editor of pair.editors) {
            expect(await views(editor)).toEqual([resolved, resolved]);
            expect(editor.getReviewItems()).toHaveLength(0);
          }
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(await views(editor)).toEqual(pending);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(await views(editor)).toEqual([resolved, resolved]);
          const saved = await pair.savedDigests();
          expect(saved[0]).toEqual(saved[1]);
          const reopened = await createTwoEditorFixture({
            document: new Uint8Array(await alice!.save()),
            ownedAttributes: true,
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          });
          try {
            for (const editor of reopened.editors)
              expect(await views(editor)).toEqual([resolved, resolved]);
          } finally {
            reopened.destroy();
          }
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 30_000);
