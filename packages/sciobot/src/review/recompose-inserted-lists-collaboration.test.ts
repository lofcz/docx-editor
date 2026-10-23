// The test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { propertyContext, propertyEditor } from '../../test/support/property-execution';
import { groupToken } from '../../test/support/review-groups';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { recomposeOwnedProposals } from './recompose';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
const context = { ...propertyContext, date: '2026-10-23T00:00:00Z' };
async function views(editor: DocxEditorInstance) {
  const a = await readCommittedView(editor.surface!.session, groupToken(editor)),
    b = await readWorkingView(editor.surface!.session, groupToken(editor));
  if (!a.success || !b.success) throw Error('Missing views');
  return [a.view.projection.text.trimEnd(), b.view.projection.text.trimEnd()];
}
for (const checklist of [false, true])
  for (const receiver of [0, 1] as const)
    test(`born ${checklist ? 'checklists' : 'lists'} retain peer text, decisions, history, geometry, and Word reopen; receiver ${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: docxFixture(
          '<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>'
        ),
        documentId: context.documentId,
        ownedAttributes: true,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors;
        const firstBody = checklist ? '- [ ] One\n- [x] Two' : '- One\n- Two';
        const nextBody = checklist
          ? '- [x] One revised\n- [x] Two\n- [ ] Three'
          : '- One revised\n- Two\n- Three';
        const firstWord = firstBody.replace(/\n/g, '\n\n'),
          nextWord = nextBody.replace(/\n/g, '\n\n');
        const first = await recomposeOwnedProposals(
          alice!,
          {
            revision: groupToken(alice!),
            targetMarkdown: `Before\n\n${firstBody}\n\nAfter`,
            ownership: [],
            contentProofs: [],
          },
          context
        );
        if (!first.success) throw Error(first.error);
        await pair.settle();
        pair.pause();
        const paragraphId = bob!
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob!.surface!.session.part(), id) === 'Before')!;
        expect(
          bob!.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId, offset: 6, text: ' Peer' },
          ]).rejected
        ).toBe(false);
        const second = await recomposeOwnedProposals(
          alice!,
          {
            revision: groupToken(alice!),
            targetMarkdown: `Before\n\n${nextBody}\n\nAfter`,
            ownership: first.ownership,
            contentProofs: first.contentProofs,
          },
          context
        );
        if (!second.success) throw Error(second.error);
        pair.reconnect(receiver);
        await pair.settle();
        const original = 'Before Peer\n\nAfter',
          proposed = `Before Peer\n\n${nextWord}\n\nAfter`;
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([original, proposed]);
        bob!.surface!.setSelection({
          anchor: { paragraphId, offset: 11 },
          head: { paragraphId, offset: 11 },
        });
        await pair.settle();
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (value) => value.actorId === 'bob'
        );
        expect(remote?.head.nodeId).toBe(paragraphId);
        expect(remote?.head.offset).toBe(11);
        const carets = pair.editors.map((editor) =>
          caretAt(editor.surface!.layout(), { paragraphId, offset: 11 })
        );
        expect(carets[0]).not.toBeNull();
        expect(carets[1]).not.toBeNull();
        expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({ x: carets[1]!.x, y: carets[1]!.y });
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([original, `Before Peer\n\n${firstWord}\n\nAfter`]);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([original, proposed]);
        for (const action of ['accept', 'reject'] as const) {
          expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(await views(editor)).toEqual(
              action === 'accept' ? [proposed, proposed] : [original, original]
            );
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(await views(editor)).toEqual([original, proposed]);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
        }
        const saved = readOoxmlPackage(new Uint8Array(await alice!.save()));
        if (!saved.ok) throw Error(saved.reason);
        const reopened = propertyEditor(saved.package);
        try {
          expect(await views(reopened)).toEqual([original, proposed]);
        } finally {
          reopened.destroy();
        }
        pair.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        pair.destroy();
      }
    });
