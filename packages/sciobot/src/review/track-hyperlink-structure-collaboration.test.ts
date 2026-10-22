// The existing test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  paragraphTextOf,
  findNode,
  paragraphOffsetIndex,
  collectRevisionSites,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  document_,
  paragraph,
  linked,
  context,
  plan,
  projection,
  ast,
} from '../../test/support/link-structure-execution';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackHyperlinkPlan } from './track-hyperlinks';
import { groups } from '../../test/support/review-groups';
type Kind = 'insert' | 'remove';
type Pair = Awaited<ReturnType<typeof createTwoEditorFixture>>;
const target = (kind: Kind) =>
  kind === 'insert' ? '[**Alpha new end**](https://example.org/new "New")' : '**Alpha new end**';
const options = (kind: Kind) => ({
  document: document_(kind === 'remove' ? paragraph(linked('Alpha end')) : undefined),
  documentId: context.documentId,
  ownedAttributes: true,
  editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
});
async function same(pair: Pair, original: string, proposed: string) {
  for (const editor of pair.editors) {
    expect(ast((await projection(editor, true)).text)).toBe(ast(original));
    expect(ast((await projection(editor)).text)).toBe(ast(proposed));
  }
}
async function geometry(pair: Pair) {
  const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
  const part = bob.surface!.session.part(),
    id = bob.surface!.session.paragraphIds()[0]!,
    p = findNode(part, id)!;
  if (p.kind !== 'paragraph') throw Error('paragraph');
  const offsets = paragraphOffsetIndex(p),
    insert = collectRevisionSites(part).find((site) => site.node.kind === 'revisionInsert')!;
  const span = offsets.spanOf(insert.node);
  if (!span) throw Error('insertion span');
  const offset = span.start + Math.min(2, span.end - span.start);
  bob.surface!.setSelection({
    anchor: { paragraphId: id, offset },
    head: { paragraphId: id, offset },
  });
  await pair.settle();
  const remote = pair.rooms[0]!.session.remoteSelections().find((item) => item.actorId === 'bob');
  expect(remote?.head.nodeId).toBe(id);
  expect(remote?.head.offset).toBe(offset);
  const a = caretAt(alice.surface!.layout(), { paragraphId: id, offset }),
    b = caretAt(bob.surface!.layout(), { paragraphId: id, offset });
  expect(a).not.toBeNull();
  expect(b).not.toBeNull();
  expect({ x: a!.x, y: a!.y }).toEqual({ x: b!.x, y: b!.y });
  expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(0);
}
async function saved(pair: Pair, kind: Kind) {
  const expected = [
    ast((await projection(pair.editors[0]!, true)).text),
    ast((await projection(pair.editors[0]!)).text),
  ];
  const digests = await pair.savedDigests();
  expect(digests[0]).toEqual(digests[1]);
  for (const room of [false, true]) {
    const again = await createTwoEditorFixture({
      ...options(kind),
      document: new Uint8Array(await pair.editors[0]!.save()),
      clientIds: [181, 211],
      ...(room ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
    });
    try {
      for (const editor of again.editors)
        expect([
          ast((await projection(editor, true)).text),
          ast((await projection(editor)).text),
        ]).toEqual(expected);
    } finally {
      again.destroy();
    }
  }
}
for (const kind of ['insert', 'remove'] as const)
  for (const clientIds of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      for (const action of ['accept', 'reject'] as const)
        test(`private D27 ${kind}: histories, ${action}, reconnect ${clientIds}/${receiver}`, async () => {
          const pair = await createTwoEditorFixture({ ...options(kind), clientIds });
          try {
            const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
            const original = (await projection(alice, true)).text,
              tail = bob
                .surface!.session.paragraphIds()
                .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Tail')!;
            pair.pause();
            const result = trackHyperlinkPlan(
              alice,
              await plan(alice, target(kind) + '\n\nTail'),
              context
            );
            if (!result.success) throw Error(result.error);
            expect(result.proposals).toHaveLength(1);
            expect(
              bob.surface!.applyAutomationOps(
                () => [{ op: 'insertText', paragraphId: tail, offset: 4, text: ' Peer' }],
                { kind: 'body' },
                [],
                false,
                true
              )
            ).toMatchObject({ committed: true, rejected: false });
            pair.reconnect(receiver);
            await pair.settle();
            const before = original.replace('Tail', 'Tail Peer'),
              after = target(kind) + '\n\nTail Peer';
            await same(pair, before, after);
            expect(groups(alice).map((group) => group.id)).toEqual(
              groups(bob).map((group) => group.id)
            );
            expect(groups(alice).map((group) => group.id)).toEqual(
              result.proposals.map((proposal) => proposal.identity)
            );
            await geometry(pair);
            await saved(pair, kind);
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await same(pair, before, before);
            expect(alice.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await same(pair, before, after);
            expect(bob.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await same(pair, original, target(kind) + '\n\nTail');
            expect(bob.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await same(pair, before, after);
            expect(bob.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
            await pair.settle();
            const decided = action === 'accept' ? after : before;
            await same(pair, decided, decided);
            expect(bob.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await same(pair, before, after);
            expect(bob.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await same(pair, decided, decided);
            await saved(pair, kind);
            pair.ownership.forEach((handle) => handle.assertHealthy());
          } finally {
            pair.destroy();
          }
        }, 30000);

for (const kind of ['insert', 'remove'] as const)
  for (const clientIds of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      test(`private D27 ${kind}: peer owner deletion and both histories ${clientIds}/${receiver}`, async () => {
        const pair = await createTwoEditorFixture({ ...options(kind), clientIds });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
          const original = (await projection(alice, true)).text;
          const owner = bob.surface!.session.paragraphIds()[0]!,
            tail = bob.surface!.session.paragraphIds()[1]!;
          pair.pause();
          const result = trackHyperlinkPlan(
            alice,
            await plan(alice, target(kind) + '\n\nTail'),
            context
          );
          if (!result.success) throw Error(result.error);
          expect(
            bob.surface!.applyAutomationOps(
              () => [{ op: 'insertText', paragraphId: tail, offset: 4, text: ' Peer' }],
              { kind: 'body' },
              [],
              false,
              true
            )
          ).toMatchObject({ committed: true, rejected: false });
          expect(
            bob.surface!.applyAutomationOps(
              () => [{ op: 'deleteBlock', blockId: owner }],
              { kind: 'body' },
              [],
              false,
              true
            )
          ).toMatchObject({ committed: true, rejected: false });
          pair.reconnect(receiver);
          await pair.settle();
          await same(pair, 'Tail Peer', 'Tail Peer');
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await same(pair, 'Tail Peer', 'Tail Peer');
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await same(pair, 'Tail Peer', 'Tail Peer');
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await same(pair, original.replace('Tail', 'Tail Peer'), target(kind) + '\n\nTail Peer');
          await geometry(pair);
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await same(pair, 'Tail Peer', 'Tail Peer');
          await saved(pair, kind);
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 30000);
