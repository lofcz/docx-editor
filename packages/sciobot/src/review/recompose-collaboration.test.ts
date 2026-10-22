import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  collectRevisionSites,
  findNode,
  paragraphTextOf,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { groupToken, groups } from '../../test/support/review-groups';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readReviewGroups } from './groups';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';
import { recomposeOwnedProposals } from './recompose';
type Wave = Extract<Awaited<ReturnType<typeof recomposeOwnedProposals>>, { success: true }>;

const context = {
  documentId: 'property-records',
  epoch: 'load-1',
  author: 'Bot',
  date: '2026-10-21T09:00:00.000Z',
};
const options = () => ({
  document: docxFixture(
    ['Alpha end', 'Keep', 'Stay', 'Tail']
      .map((text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`)
      .join('')
  ),
  documentId: context.documentId,
  ownedAttributes: true,
  editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
});
const pkg = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();
async function wave(
  editor: DocxEditorInstance,
  targetMarkdown: string,
  previous?: Pick<Wave, 'ownership' | 'contentProofs'>
) {
  const revision = groupToken(editor);
  const result = await recomposeOwnedProposals(
    editor,
    {
      revision,
      targetMarkdown,
      ownership: previous?.ownership ?? [],
      ...(previous ? { contentProofs: previous.contentProofs } : {}),
    },
    { ...context, epoch: revision.epoch }
  );
  if (!result.success) throw new Error(result.error);
  expect(result.committed).toBe(true);
  return { ownership: result.ownership, contentProofs: result.contentProofs };
}
function unnested(editor: DocxEditorInstance) {
  const visit = (node: OoxmlNode, inside = false) => {
    if (node.kind === 'textValue') return;
    const revision = node.kind === 'revisionInsert' || node.kind === 'revisionDelete';
    if (revision) expect(inside).toBe(false);
    for (const child of node.children) visit(child, inside || revision);
  };
  visit(editor.surface!.session.part().root);
  const sites = collectRevisionSites(editor.surface!.session.part());
  expect(new Set(sites.map((site) => site.node.id)).size).toBe(sites.length);
}
function append(editor: DocxEditorInstance, paragraphId: string, text: string) {
  const offset = paragraphTextOf(editor.surface!.session.part(), paragraphId)!.length;
  const result = editor.surface!.applyAutomationOps(
    () => [{ op: 'insertText', paragraphId, offset, text, bias: 'right' }],
    { kind: 'body' },
    [],
    false,
    true // Each append represents a separate user action, independent of the capture window.
  );
  expect(result).toMatchObject({ committed: true, rejected: false });
}
async function views(editor: DocxEditorInstance) {
  const [before, after] = await Promise.all([
    readCommittedView(editor.surface!.session, groupToken(editor)),
    readWorkingView(editor.surface!.session, groupToken(editor)),
  ]);
  if (!before.success || !after.success) throw new Error('Cannot read native staged views.');
  return [before.view.projection.text, after.view.projection.text];
}
async function geometry(pair: Awaited<ReturnType<typeof createTwoEditorFixture>>, id: string) {
  const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
    offset = paragraphTextOf(bob.surface!.session.part(), id)!.length;
  bob.surface!.setSelection({
    anchor: { paragraphId: id, offset },
    head: { paragraphId: id, offset },
  });
  await pair.settle();
  const remote = pair.rooms[0]!.session.remoteSelections().find((r) => r.actorId === 'bob');
  expect(remote?.head.nodeId).toBe(id);
  expect(remote?.head.offset).toBe(offset);
  const a = caretAt(alice.surface!.layout(), { paragraphId: id, offset }),
    b = caretAt(bob.surface!.layout(), { paragraphId: id, offset });
  expect(a).not.toBeNull();
  expect(b).not.toBeNull();
  expect({ x: a!.x, y: a!.y }).toEqual({ x: b!.x, y: b!.y });
  expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(0);
}
async function saved(pair: Awaited<ReturnType<typeof createTwoEditorFixture>>) {
  const expected = await views(pair.editors[0]!);
  expect(await views(pair.editors[1]!)).toEqual(expected);
  const digests = await pair.savedDigests();
  expect(digests[0]).toEqual(digests[1]);
  const read = readReviewGroups(pkg(pair.editors[0]!));
  if (!read.success) throw new Error(read.error);
  for (const room of [false, true]) {
    const reopened = await createTwoEditorFixture({
      ...options(),
      document: new Uint8Array(await pair.editors[0]!.save()),
      clientIds: room ? [401, 433] : [181, 211],
      ...(room ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
    });
    try {
      for (const editor of reopened.editors) {
        expect(await views(editor)).toEqual(expected);
        const groups = readReviewGroups(pkg(editor));
        if (!groups.success) throw new Error(groups.error);
        expect(groups.groups.map((g) => g.id)).toEqual(read.groups.map((g) => g.id));
      }
    } finally {
      reopened.destroy();
    }
  }
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const) {
    test(`later text waves preserve mixed decisions, manual text, concurrent peers, history, geometry, and reopening; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          [main, accepted, rejected, tail] = alice.surface!.session.paragraphIds();
        const ownership = await wave(alice, 'Alpha old end\n\nKept\n\nGone\n\nTail');
        await pair.settle();
        for (const [paragraph, decide] of [
          [accepted, acceptReviewGroup],
          [rejected, rejectReviewGroup],
        ] as const) {
          const selected = groups(bob).filter((group) => group.anchor?.paragraphId === paragraph);
          expect(selected.length).toBeGreaterThan(0);
          for (const group of selected) {
            const revision = groupToken(bob);
            expect(decide(bob, group.id, revision, revision).success).toBe(true);
            await pair.settle();
          }
        }
        append(bob, main!, ' manual');
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([
            'Alpha end manual\n\nKept\n\nStay\n\nTail',
            'Alpha old end manual\n\nKept\n\nStay\n\nTail',
          ]);
        pair.pause();
        const next = await wave(alice, 'Alpha new end manual\n\nKept\n\nStay\n\nTail', ownership);
        expect(next.ownership.length).toBeGreaterThan(0);
        append(bob, tail!, ' Peer');
        pair.reconnect(receiver);
        await pair.settle();
        const committed = 'Alpha end manual\n\nKept\n\nStay\n\nTail Peer',
          working = 'Alpha new end manual\n\nKept\n\nStay\n\nTail Peer',
          prior = 'Alpha old end manual\n\nKept\n\nStay\n\nTail Peer';
        const check = async (text: string) => {
          for (const editor of pair.editors) {
            expect(await views(editor)).toEqual([committed, text]);
            expect(groups(editor).every((group) => group.anchor?.paragraphId === main)).toBe(true);
            unnested(editor);
          }
        };
        await check(working);
        for (let cycle = 0; cycle < 2; cycle++) {
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(prior);
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(working);
        }
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([
            committed.replace(' Peer', ''),
            working.replace(' Peer', ''),
          ]);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await geometry(pair, main!);
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30000);

    test(`later waves cannot recreate a concurrently deleted owner and retain both histories; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          id = alice.surface!.session.paragraphIds()[0]!;
        const ownership = await wave(alice, 'Alpha old end\n\nKeep\n\nStay\n\nTail');
        await pair.settle();
        pair.pause();
        await wave(alice, 'Alpha new end\n\nKeep\n\nStay\n\nTail', ownership);
        expect(
          bob.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: id }])
        ).toMatchObject({ committed: true, rejected: false });
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors)
          expect(findNode(editor.surface!.session.part(), id)).toBeNull();
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([
            'Alpha end\n\nKeep\n\nStay\n\nTail',
            'Alpha new end\n\nKeep\n\nStay\n\nTail',
          ]);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([
            'Alpha end\n\nKeep\n\nStay\n\nTail',
            'Alpha old end\n\nKeep\n\nStay\n\nTail',
          ]);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await geometry(pair, id);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(findNode(editor.surface!.session.part(), id)).toBeNull();
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30000);
  }

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`later wave refuses peer changes inside owned text without losing history; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          id = alice.surface!.session.paragraphIds()[0]!,
          tail = alice.surface!.session.paragraphIds().at(-1)!;
        const previous = await wave(alice, 'Alpha old end\n\nKeep\n\nStay\n\nTail');
        await pair.settle();
        append(alice, tail, ' redo');
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        pair.pause();
        expect(
          bob.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: id, offset: 7, text: ' peer', bias: 'right' },
          ])
        ).toMatchObject({ committed: true, rejected: false });
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([
            'Alpha end\n\nKeep\n\nStay\n\nTail',
            'Alpha o peerld end\n\nKeep\n\nStay\n\nTail',
          ]);
        const before = pair.editors.map((editor) => ({
            pkg: pkg(editor),
            revision: editor.surface!.session.packageRevision(),
            groups: readReviewGroups(pkg(editor)),
          })),
          vectors = pair.documents.map((doc) => Y.encodeStateVector(doc)),
          history = pair.rooms.map((room) => [room.session.canUndo(), room.session.canRedo()]);
        expect(history[0]![1]).toBe(true);
        const revision = groupToken(alice),
          result = await recomposeOwnedProposals(
            alice,
            {
              revision,
              targetMarkdown: 'Alpha new end\n\nKeep\n\nStay\n\nTail',
              ...previous,
            },
            { ...context, epoch: revision.epoch }
          );
        expect(result).toMatchObject({ success: false, code: 'ambiguous-target' });
        for (const [index, editor] of pair.editors.entries()) {
          expect(pkg(editor)).toBe(before[index]!.pkg);
          expect(editor.surface!.session.packageRevision()).toBe(before[index]!.revision);
          expect(readReviewGroups(pkg(editor))).toEqual(before[index]!.groups);
          expect(Y.encodeStateVector(pair.documents[index]!)).toEqual(vectors[index]);
        }
        expect(pair.rooms.map((room) => [room.session.canUndo(), room.session.canRedo()])).toEqual(
          history
        );
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([
            'Alpha end\n\nKeep\n\nStay\n\nTail redo',
            'Alpha o peerld end\n\nKeep\n\nStay\n\nTail redo',
          ]);
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30000);

test('nested recomposition previews cannot leak native removals into an outer shared transaction', async () => {
  const pair = await createTwoEditorFixture({ ...options(), clientIds: [41, 73] });
  try {
    const alice = pair.editors[0]!,
      session = alice.surface!.session,
      tail = session.paragraphIds().at(-1)!;
    const previous = await wave(alice, 'Alpha old end\n\nKeep\n\nStay\n\nTail');
    await pair.settle();
    const revision = groupToken(alice),
      originalGroups = readReviewGroups(pkg(alice));
    let pending: ReturnType<typeof recomposeOwnedProposals> | undefined;
    const result = session.applyTreeOps(
      [],
      undefined,
      undefined,
      { kind: 'body' },
      {
        stage: (ctx) => {
          pending = recomposeOwnedProposals(
            alice,
            {
              revision,
              targetMarkdown: 'Alpha new end\n\nKeep\n\nStay\n\nTail',
              ...previous,
            },
            { ...context, epoch: revision.epoch }
          );
          expect(
            ctx.apply({
              op: 'insertText',
              paragraphId: tail,
              offset: 4,
              text: ' outer',
              bias: 'right',
            })
          ).toBe(true);
        },
      }
    );
    expect(result).toMatchObject({ committed: true, rejected: false });
    expect(pending).toBeDefined();
    expect(await pending!).toMatchObject({ success: false, code: 'stale-revision' });
    await pair.settle();
    for (const editor of pair.editors) {
      expect(await views(editor)).toEqual([
        'Alpha end\n\nKeep\n\nStay\n\nTail outer',
        'Alpha old end\n\nKeep\n\nStay\n\nTail outer',
      ]);
      expect(readReviewGroups(pkg(editor))).toEqual(originalGroups);
      unnested(editor);
    }
    expect(alice.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    for (const editor of pair.editors)
      expect(await views(editor)).toEqual([
        'Alpha end\n\nKeep\n\nStay\n\nTail',
        'Alpha old end\n\nKeep\n\nStay\n\nTail',
      ]);
    expect(alice.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    await saved(pair);
  } finally {
    pair.destroy();
  }
}, 30000);
