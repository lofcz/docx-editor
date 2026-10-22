import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  bodyStoryRoot,
  findNode,
  type OoxmlElement,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import {
  blockAst,
  blockContext,
  blockEditor,
  blockP,
  blockToken,
  blockViews,
  decideBlocks,
  manualBlockPlan,
} from '../../test/support/block-execution';
import { createEditPlan } from '../diff/plan';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackSiblingPlan } from './track-blocks';

type Pair = Awaited<ReturnType<typeof createTwoEditorFixture>>;
const options = (body: string, clientIds: readonly [number, number]) => ({
  document: docxFixture(body),
  clientIds,
  editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
});
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
async function views(pair: Pair, original: string, working = original) {
  for (const editor of pair.editors)
    expect((await blockViews(editor)).map(blockAst)).toEqual([
      blockAst(original),
      blockAst(working),
    ]);
}
async function saved(pair: Pair) {
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
  for (const receiver of [0, 1] as const) {
    test(`explicit source deletion wins over an offline edit after a resolved middle join; clients ${clients.join('/')}, receiver ${receiver}`, async () => {
      const pair = await createTwoEditorFixture(
        options(blockP('Opening') + blockP('Closing'), clients)
      );
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const tracked = trackSiblingPlan(alice, manualBlockPlan(alice, 1, 0, 'New'), blockContext);
        if (!tracked.success) throw new Error(tracked.error);
        await pair.settle();
        const id = bob.surface!.session.paragraphIds()[1]!;
        decideBlocks(alice, tracked.proposals, 'rejectRevision');
        await pair.settle();
        pair.pause();
        apply(alice, [{ op: 'deleteBlock', blockId: id }]);
        apply(bob, [{ op: 'insertText', paragraphId: id, offset: 0, text: 'Peer' }]);
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors)
          expect(findNode(editor.surface!.session.part(), id)).toBeNull();
        await views(pair, 'Opening\n\nClosing');
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await views(pair, 'Opening\n\nPeerClosing');
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await views(pair, 'Opening\n\nClosing');
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30_000);

    test(`rejecting a whole inserted section keeps peer text from its last paragraph; clients ${clients.join('/')}, receiver ${receiver}`, async () => {
      const pair = await createTwoEditorFixture(
        options(blockP('Opening') + blockP('Closing'), clients)
      );
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const tracked = trackSiblingPlan(
          alice,
          manualBlockPlan(alice, 1, 0, '## New\n\nFirst\n\nSecond'),
          blockContext
        );
        if (!tracked.success) throw new Error(tracked.error);
        await pair.settle();
        const id = bob.surface!.session.paragraphIds()[3]!;
        pair.pause();
        decideBlocks(alice, tracked.proposals, 'rejectRevision');
        apply(bob, [{ op: 'insertText', paragraphId: id, offset: 6, text: 'Peer', bias: 'right' }]);
        pair.reconnect(receiver);
        await pair.settle();
        await views(pair, 'Opening\n\nPeerClosing');
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await views(
          pair,
          'Opening\n\nPeerClosing',
          'Opening\n\n## New\n\nFirst\n\nSecondPeer\n\nClosing'
        );
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await views(pair, 'Opening\n\nPeerClosing');
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30_000);
  }

test('a recovered source supports ordinary split/join editing and independent peer history', async () => {
  const pair = await createTwoEditorFixture(
    options(blockP('Opening') + blockP('Closing'), [41, 73])
  );
  try {
    const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
    const tracked = trackSiblingPlan(alice, manualBlockPlan(alice, 1, 0, 'New'), blockContext);
    if (!tracked.success) throw new Error(tracked.error);
    await pair.settle();
    const id = bob.surface!.session.paragraphIds()[1]!,
      opening = bob.surface!.session.paragraphIds()[0]!;
    pair.pause();
    decideBlocks(alice, tracked.proposals, 'rejectRevision');
    apply(bob, [{ op: 'insertText', paragraphId: id, offset: 3, text: 'Peer', bias: 'right' }]);
    pair.reconnect();
    await pair.settle();
    apply(bob, [{ op: 'splitParagraph', paragraphId: id, offset: 2 }]);
    await pair.settle();
    await views(pair, 'Opening\n\nPe\n\nerClosing');
    apply(alice, [{ op: 'insertText', paragraphId: opening, offset: 0, text: 'Alice: ' }]);
    await pair.settle();
    expect(bob.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    await views(pair, 'Alice: Opening\n\nPeerClosing');
    expect(bob.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    await views(pair, 'Alice: Opening\n\nPe\n\nerClosing');
    const tail = bob.surface!.session.paragraphIds()[2]!;
    bob.surface!.setSelection({
      anchor: { paragraphId: tail, offset: 1 },
      head: { paragraphId: tail, offset: 1 },
    });
    await pair.settle();
    expect(caretAt(alice.surface!.layout(), { paragraphId: tail, offset: 1 })).not.toBeNull();
    apply(bob, [{ op: 'joinParagraphs', firstId: id, secondId: tail }]);
    await pair.settle();
    await views(pair, 'Alice: Opening\n\nPeerClosing');
    await saved(pair);
  } finally {
    pair.destroy();
  }
}, 30_000);

test('a paragraph before a table retains offline text, cell geometry, review history, and Word convergence', async () => {
  const table = (value: string) =>
    `<w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc>${blockP(value)}</w:tc></w:tr></w:tbl>`;
  const pair = await createTwoEditorFixture(
    options(blockP('Opening') + table('Cell') + blockP('Closing'), [41, 73])
  );
  const oracle = blockEditor(blockP('Opening') + table('PeerCell') + blockP('Closing'));
  try {
    const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
      root = bodyStoryRoot(alice.surface!.session.part()) as OoxmlElement;
    const made = createEditPlan(blockToken(alice), [
      {
        kind: 'siblings',
        partName: alice.surface!.session.part().name,
        parentId: root.id,
        beforeId: root.children[0]!.id,
        afterId: root.children[1]!.id,
        removedIds: [],
        inserted: blockAst('New').children as import('../projection/parse-target').TargetAstNode[],
      },
    ]);
    if (!made.success) throw new Error(made.error);
    const tracked = trackSiblingPlan(alice, made.plan, blockContext);
    if (!tracked.success) throw new Error(tracked.error);
    await pair.settle();
    const id = bob.surface!.session.paragraphIds()[1]!;
    pair.pause();
    decideBlocks(alice, tracked.proposals, 'rejectRevision');
    apply(bob, [{ op: 'insertText', paragraphId: id, offset: 3, text: 'Peer', bias: 'right' }]);
    bob.surface!.setSelection({
      anchor: { paragraphId: id, offset: 7 },
      head: { paragraphId: id, offset: 7 },
    });
    pair.reconnect();
    await pair.settle();
    const expected = (await blockViews(oracle.editor)).map(blockAst);
    for (const editor of pair.editors)
      expect((await blockViews(editor)).map(blockAst)).toEqual(expected);
    expect(caretAt(alice.surface!.layout(), { paragraphId: id, offset: 4 })).not.toBeNull();
    expect(alice.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    for (const editor of pair.editors)
      expect(blockAst((await blockViews(editor))[0]!)).toEqual(expected[0]);
    expect(alice.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    for (const editor of pair.editors)
      expect((await blockViews(editor)).map(blockAst)).toEqual(expected);
    const cell = bob.surface!.session.paragraphIds()[2]!;
    bob.surface!.setSelection({
      anchor: { paragraphId: cell, offset: 0 },
      head: { paragraphId: cell, offset: 0 },
    });
    bob.surface!.deleteBackward();
    await pair.settle();
    const shortened = blockEditor(blockP('Opening') + table('PeeCell') + blockP('Closing'));
    try {
      for (const editor of pair.editors)
        expect((await blockViews(editor)).map(blockAst)).toEqual(
          (await blockViews(shortened.editor)).map(blockAst)
        );
    } finally {
      shortened.destroy();
    }
    bob.surface!.setSelection({
      anchor: { paragraphId: id, offset: 1 },
      head: { paragraphId: cell, offset: 1 },
    });
    bob.surface!.type('X');
    await pair.settle();
    const replaced = blockEditor(blockP('Opening') + table('PXell') + blockP('Closing'));
    try {
      for (const editor of pair.editors)
        expect((await blockViews(editor)).map(blockAst)).toEqual(
          (await blockViews(replaced.editor)).map(blockAst)
        );
    } finally {
      replaced.destroy();
    }
    await saved(pair);
  } finally {
    pair.destroy();
    oracle.destroy();
  }
}, 30_000);
