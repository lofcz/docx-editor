import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { bodySectionIndexOf } from '../../../core/src/editor/section-scope';
import { collectRevisionSites, findNode, type TreeDocOp } from '@docx-editor.dev/core/store';
import pairs from '../__tests__/fixtures/ai-pairs.json';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import {
  blockAst,
  blockContext,
  blockP,
  blockViews,
  decideBlocks,
  manualBlockPlan,
  targetBlockPlan,
} from '../../test/support/block-execution';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackSiblingPlan } from './track-blocks';

type Pair = Awaited<ReturnType<typeof createTwoEditorFixture>>;
function apply(editor: DocxEditorInstance, ops: readonly TreeDocOp[]) {
  const result = editor.surface!.applyAutomationOps(
    () => ops,
    { kind: 'body' },
    undefined,
    false,
    true
  );
  if (result.rejected || !result.committed) throw new Error(JSON.stringify(result));
}
async function assertViews(pair: Pair, original: string, working = original) {
  for (const editor of pair.editors)
    expect((await blockViews(editor)).map(blockAst)).toEqual([
      blockAst(original),
      blockAst(working),
    ]);
}
async function assertSaved(pair: Pair) {
  const digests = await pair.savedDigests();
  expect(digests[0]).toEqual(digests[1]);
  for (const editor of pair.editors) {
    const reopened = createDocxEditor({
      container: document.createElement('div'),
      document: new Uint8Array(await editor.save()),
      modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
    });
    try {
      expect((await blockViews(reopened)).map(blockAst)).toEqual(
        (await blockViews(editor)).map(blockAst)
      );
    } finally {
      reopened.destroy();
    }
  }
}

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const index of [2, 5])
      for (const decision of ['acceptRevision', 'rejectRevision'] as const)
        test(`C01 sibling ${index} ${decision} keeps offline peer text, selection, and history; clients ${clients.join('/')}, receiver ${receiver}`, async () => {
          const fixture = pairs.cases[index]!,
            pair = await createTwoEditorFixture({
              document: docxFixture(fixture.beforeMarkdown.split('\n\n').map(blockP).join('')),
              clientIds: clients,
              editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
            });
          try {
            const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
              ids = alice.surface!.session.paragraphIds(),
              tail = ids.at(-1)!,
              plan = await targetBlockPlan(alice, fixture.afterMarkdown);
            bob.surface!.setSelection({
              anchor: { paragraphId: tail, offset: 0 },
              head: { paragraphId: tail, offset: 2 },
            });
            await pair.settle();
            pair.pause();
            const tracked = trackSiblingPlan(alice, plan, blockContext);
            if (!tracked.success) throw new Error(tracked.error);
            apply(bob, [{ op: 'insertText', paragraphId: ids[0]!, offset: 0, text: 'Peer: ' }]);
            pair.reconnect(receiver);
            await pair.settle();
            const original = `Peer: ${fixture.beforeMarkdown}`,
              working = `Peer: ${fixture.afterMarkdown}`;
            await assertViews(pair, original, working);
            for (const editor of pair.editors) {
              expect(findNode(editor.surface!.session.part(), ids[0]!)).not.toBeNull();
              expect(findNode(editor.surface!.session.part(), tail)).not.toBeNull();
            }
            const remote = pair.rooms[0]!.session.remoteSelections().find(
              (selection) => selection.actorId === 'bob'
            )!;
            expect(remote.anchor.nodeId).toBe(tail);
            expect(remote.head.nodeId).toBe(tail);
            expect(
              caretAt(alice.surface!.layout(), { paragraphId: tail, offset: remote.head.offset })
            ).not.toBeNull();
            expect(
              pair.containers[0]!.querySelectorAll('.docx-remote-selection-rect').length
            ).toBeGreaterThan(0);
            await assertSaved(pair);
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await assertViews(pair, original);
            expect(alice.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await assertViews(pair, original, working);
            decideBlocks(alice, tracked.proposals, decision);
            await pair.settle();
            const expected = decision === 'acceptRevision' ? working : original;
            await assertViews(pair, expected);
            for (const editor of pair.editors)
              expect(collectRevisionSites(editor.surface!.session.part())).toHaveLength(0);
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await assertViews(pair, original, working);
            expect(alice.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await assertViews(pair, expected);
            await assertSaved(pair);
          } finally {
            pair.destroy();
          }
        }, 30_000);

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const change of ['insertion', 'deletion'] as const)
      test(`resolving a middle paragraph ${change} preserves unseen peer text at the native join; clients ${clients.join('/')}, receiver ${receiver}`, async () => {
        const pair = await createTwoEditorFixture({
          document: docxFixture(
            blockP('Opening') + (change === 'deletion' ? blockP('Old') : '') + blockP('Closing')
          ),
          clientIds: clients,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
          const tracked = trackSiblingPlan(
            alice,
            manualBlockPlan(
              alice,
              1,
              change === 'deletion' ? 1 : 0,
              change === 'deletion' ? '' : 'New'
            ),
            blockContext
          );
          if (!tracked.success) throw new Error(tracked.error);
          await pair.settle();
          // Allocated canonical IDs can change during native journal publication. Read the current tree.
          const id = bob.surface!.session.paragraphIds()[1]!;
          const tail = bob.surface!.session.paragraphIds().at(-1)!;
          for (const editor of pair.editors)
            expect(bodySectionIndexOf(editor.surface!.session, id)).toBe(0);
          for (const editor of pair.editors)
            editor.surface!.setSelection({
              anchor: { paragraphId: id, offset: 3 },
              head: { paragraphId: id, offset: 3 },
            });
          await pair.settle();
          pair.pause();
          decideBlocks(
            alice,
            tracked.proposals,
            change === 'deletion' ? 'acceptRevision' : 'rejectRevision'
          );
          const selection = alice.surface!.state().selection!;
          expect(caretAt(alice.surface!.layout(), selection.head)).not.toBeNull();
          expect(bodySectionIndexOf(alice.surface!.session, tail)).toBe(0);
          apply(bob, [
            { op: 'insertText', paragraphId: id, offset: 3, text: 'Peer', bias: 'right' },
          ]);
          bob.surface!.setSelection({
            anchor: { paragraphId: id, offset: 7 },
            head: { paragraphId: id, offset: 7 },
          });
          pair.reconnect(receiver);
          await pair.settle();
          await assertViews(pair, 'Opening\n\nPeerClosing');
          for (const editor of pair.editors)
            expect(bodySectionIndexOf(editor.surface!.session, id)).toBe(0);
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (entry) => entry.actorId === 'bob'
          )!;
          expect(remote.head.nodeId).toBe(id);
          expect(remote.head.offset).toBe(4);
          expect(caretAt(alice.surface!.layout(), { paragraphId: id, offset: 4 })).not.toBeNull();
          expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(
            0
          );
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await assertViews(
            pair,
            change === 'deletion' ? 'Opening\n\nOldPeer\n\nClosing' : 'Opening\n\nPeerClosing',
            change === 'deletion' ? 'Opening\n\nPeerClosing' : 'Opening\n\nNewPeer\n\nClosing'
          );
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          // Undoing owner creation removes the owner's own text. The peer's committed text
          // stays as its own paragraph; redo must restore the owner around it.
          await assertViews(
            pair,
            change === 'deletion' ? 'Opening\n\nOldPeer\n\nClosing' : 'Opening\n\nPeer\n\nClosing'
          );
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await assertViews(
            pair,
            change === 'deletion' ? 'Opening\n\nOldPeer\n\nClosing' : 'Opening\n\nPeerClosing',
            change === 'deletion' ? 'Opening\n\nPeerClosing' : 'Opening\n\nNewPeer\n\nClosing'
          );
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await assertViews(pair, 'Opening\n\nPeerClosing');
          await assertSaved(pair);
        } finally {
          pair.destroy();
        }
      }, 30_000);

for (const change of ['insertion', 'deletion'] as const)
  test(`visible peer text at a middle paragraph ${change} follows the native forward-join oracle`, async () => {
    const pair = await createTwoEditorFixture({
      document: docxFixture(
        blockP('Opening') + (change === 'deletion' ? blockP('Old') : '') + blockP('Closing')
      ),
      clientIds: [41, 73],
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const tracked = trackSiblingPlan(
        alice,
        manualBlockPlan(
          alice,
          1,
          change === 'deletion' ? 1 : 0,
          change === 'deletion' ? '' : 'New'
        ),
        blockContext
      );
      if (!tracked.success) throw new Error(tracked.error);
      await pair.settle();
      const id = bob.surface!.session.paragraphIds()[1]!;
      apply(bob, [{ op: 'insertText', paragraphId: id, offset: 3, text: 'Peer', bias: 'right' }]);
      await pair.settle();
      const original =
          change === 'deletion' ? 'Opening\n\nOldPeer\n\nClosing' : 'Opening\n\nPeerClosing',
        working =
          change === 'deletion' ? 'Opening\n\nPeerClosing' : 'Opening\n\nNewPeer\n\nClosing';
      await assertViews(pair, original, working);
      decideBlocks(
        alice,
        tracked.proposals,
        change === 'deletion' ? 'acceptRevision' : 'rejectRevision'
      );
      await pair.settle();
      await assertViews(pair, 'Opening\n\nPeerClosing');
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      await assertViews(pair, original, working);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      await assertViews(pair, 'Opening\n\nPeerClosing');
      await assertSaved(pair);
    } finally {
      pair.destroy();
    }
  }, 30_000);
