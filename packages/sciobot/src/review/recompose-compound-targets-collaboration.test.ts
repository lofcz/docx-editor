import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  paragraphTextOf,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { document_, target as renderTarget, cell } from '../../test/support/compound-target-table';
import { groupToken, groups } from '../../test/support/review-groups';
import { targetAstKey } from '../diff/target-key';
import { parseTargetMarkdown } from '../projection/parse-target';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { recomposeOwnedProposals } from './recompose';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';

type Pair = Awaited<ReturnType<typeof createTwoEditorFixture>>;
type Wave = Extract<Awaited<ReturnType<typeof recomposeOwnedProposals>>, { success: true }>;
const context = {
  documentId: 'property-records',
  epoch: 'load-1',
  author: 'Bot',
  date: '2026-10-21T14:00:00.000Z',
};
function options() {
  return {
    document: document_(),
    documentId: context.documentId,
    ownedAttributes: true,
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
  };
}
function ast(markdown: string) {
  const parsed = parseTargetMarkdown(markdown);
  if (!parsed.success) throw new Error(parsed.error);
  return targetAstKey(parsed.target.normalizedAst);
}
async function views(editor: DocxEditorInstance) {
  const [before, after] = await Promise.all([
    readCommittedView(editor.surface!.session, groupToken(editor), { tableSyntax: 'auto' }),
    readWorkingView(editor.surface!.session, groupToken(editor), { tableSyntax: 'auto' }),
  ]);
  if (!before.success || !after.success) throw new Error('Missing mixed shared views.');
  return [before.view.projection.text, after.view.projection.text];
}
async function wave(editor: DocxEditorInstance, targetMarkdown: string, previous?: Wave) {
  const revision = groupToken(editor);
  const publications: { origin: string; fingerprint: string }[] = [];
  const unsubscribe = editor.surface!.session.subscribe((change) => {
    const source = editor.surface!.session.currentPackage();
    publications.push({
      origin: change.origin,
      fingerprint: JSON.stringify(
        [...source.parts]
          .map(([name, part]) => {
            // OPC relationships form a keyed set. Room reflection can change their XML order.
            const canonical =
              name.endsWith('.rels') && part.root.kind === 'generic'
                ? {
                    ...part,
                    root: {
                      ...part.root,
                      children: [...part.root.children].sort((a, b) =>
                        canonicalOoxmlFingerprint(a).localeCompare(canonicalOoxmlFingerprint(b))
                      ),
                    },
                  }
                : part;
            return [name, canonicalOoxmlFingerprint(canonical)];
          })
          .sort(([a], [b]) => a!.localeCompare(b!))
      ),
    });
  });
  try {
    const result = await recomposeOwnedProposals(
      editor,
      {
        revision,
        targetMarkdown,
        ownership: previous?.ownership ?? [],
        contentProofs: previous?.contentProofs ?? [],
      },
      { ...context, epoch: revision.epoch }
    );
    if (!result.success) throw new Error(`${result.code}: ${result.error}`);
    expect(result.committed).toBe(true);
    expect(publications.filter((entry) => !entry.origin.endsWith('.remote'))).toHaveLength(1);
    expect(new Set(publications.map((entry) => entry.fingerprint)).size).toBe(1);
    return result;
  } finally {
    unsubscribe();
  }
}
function append(editor: DocxEditorInstance, paragraphId: string, text: string) {
  const offset = paragraphTextOf(editor.surface!.session.part(), paragraphId)!.length;
  expect(
    editor.surface!.applyAutomationOps(
      () => [{ op: 'insertText', paragraphId, offset, text, bias: 'right' }],
      { kind: 'body' },
      [],
      false,
      true
    )
  ).toMatchObject({ committed: true, rejected: false });
}
function unnested(editor: DocxEditorInstance) {
  const visit = (node: OoxmlNode, inside = false) => {
    if (node.kind === 'textValue') return;
    const contentRevision = node.kind === 'revisionInsert' || node.kind === 'revisionDelete';
    if (contentRevision) expect(inside).toBe(false);
    for (const child of node.children) visit(child, inside || contentRevision);
  };
  visit(editor.surface!.session.part().root);
  const sites = collectRevisionSites(editor.surface!.session.part());
  expect(new Set(sites.map((site) => site.node.id)).size).toBe(sites.length);
}
async function check(pair: Pair, original: string, proposed: string) {
  for (const editor of pair.editors) {
    expect((await views(editor)).map(ast)).toEqual([original, proposed].map(ast));
    unnested(editor);
  }
}
async function geometry(pair: Pair, id: string) {
  const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
  const offset = paragraphTextOf(bob.surface!.session.part(), id)!.length;
  bob.surface!.setSelection({
    anchor: { paragraphId: id, offset },
    head: { paragraphId: id, offset },
  });
  await pair.settle();
  const remote = pair.rooms[0]!.session.remoteSelections().find((entry) => entry.actorId === 'bob');
  expect(remote?.head.nodeId).toBe(id);
  expect(remote?.head.offset).toBe(offset);
  const a = caretAt(alice.surface!.layout(), { paragraphId: id, offset });
  const b = caretAt(bob.surface!.layout(), { paragraphId: id, offset });
  expect(a).not.toBeNull();
  expect(b).not.toBeNull();
  expect({ x: a!.x, y: a!.y }).toEqual({ x: b!.x, y: b!.y });
  expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(0);
}
async function saved(pair: Pair) {
  const expected = await views(pair.editors[0]!);
  const digests = await pair.savedDigests();
  expect(digests[0]).toEqual(digests[1]);
  const expectedGroups = groups(pair.editors[0]!).map((group) => group.id);
  for (const room of [false, true]) {
    const reopened = await createTwoEditorFixture({
      ...options(),
      document: new Uint8Array(await pair.editors[0]!.save()),
      clientIds: room ? [401, 433] : [181, 211],
      ...(room ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
    });
    try {
      for (const editor of reopened.editors) {
        expect((await views(editor)).map(ast)).toEqual(expected.map(ast));
        expect(groups(editor).map((group) => group.id)).toEqual(expectedGroups);
      }
    } finally {
      reopened.destroy();
    }
  }
}

const firstTarget = renderTarget([
  [cell('Title', 2), cell('H3')],
  [cell('A'), cell('B'), cell('C')],
  [cell('D'), cell('E'), cell('F')],
]);
function completeTarget(shape: 'row' | 'column') {
  return shape === 'row'
    ? renderTarget([
        [cell('Final **bold**', 2), cell('H3')],
        [cell('A'), cell('B'), cell('C')],
        [cell('D'), cell('E'), cell('F')],
        [cell('**New** [Link](https://example.org/compound)', 2), cell('New tail')],
      ])
    : renderTarget([
        [cell('Final **bold**', 2), cell('H3\n\nNew', 2)],
        [cell('A'), cell('B'), cell('C'), cell('[Link](https://example.org/compound)')],
        [cell('D'), cell('E'), cell('F'), cell('New tail')],
      ]);
}
for (const clientIds of [
  [47, 79],
  [79, 47],
] as const)
  for (const receiver of [0, 1] as const)
    for (const shape of ['row', 'column'] as const)
      for (const decision of ['accept', 'reject'] as const)
        test(`complete target ${shape} after owned merge; ${decision}, both histories, reconnect ${clientIds}/${receiver}`, async () => {
          const pair = await createTwoEditorFixture({ ...options(), clientIds });
          try {
            const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
              original = (await views(alice))[0]!,
              peer = alice
                .surface!.session.paragraphIds()
                .find((id) => paragraphTextOf(alice.surface!.session.part(), id) === 'D')!;
            const first = await wave(alice, firstTarget);
            await pair.settle();
            pair.pause();
            const expected = completeTarget(shape),
              second = await wave(alice, expected, first);
            append(bob, peer, ' Peer');
            pair.reconnect(receiver);
            await pair.settle();
            const peerOriginal = original.replace(/\bD\b/, 'D Peer'),
              peerExpected = expected.replace(/\bD\b/, 'D Peer');
            await check(pair, peerOriginal, peerExpected);
            await geometry(pair, peer);
            await saved(pair);
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, peerOriginal, firstTarget.replace(/\bD\b/, 'D Peer'));
            expect(alice.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, peerOriginal, peerExpected);
            expect(bob.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, original, expected);
            expect(bob.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, peerOriginal, peerExpected);
            for (const group of [...groups(bob)]) {
              const result = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
                bob,
                group.id,
                groupToken(bob),
                groupToken(bob)
              );
              if (!result.success) throw new Error(result.error);
            }
            await pair.settle();
            const decided = decision === 'accept' ? peerExpected : peerOriginal;
            await check(pair, decided, decided);
            expect(bob.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, peerOriginal, peerExpected);
            expect(bob.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, decided, decided);
            const raw = await readWorkingView(alice.surface!.session, groupToken(alice));
            if (!raw.success) throw new Error(raw.error);
            await wave(alice, raw.view.projection.text.replace('Lead', 'Latest Lead'), second);
            await pair.settle();
            await check(pair, decided, decided.replace('Lead', 'Latest Lead'));
            await saved(pair);
          } finally {
            pair.destroy();
          }
        }, 30000);

for (const clientIds of [
  [47, 79],
  [79, 47],
] as const)
  for (const receiver of [0, 1] as const)
    for (const owner of ['old-row', 'whole-table'] as const)
      test(`complete target compiler retains offline ${owner} deletion through both histories; ${clientIds}/${receiver}`, async () => {
        const pair = await createTwoEditorFixture({ ...options(), clientIds });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
            original = (await views(alice))[0]!,
            tail = alice
              .surface!.session.paragraphIds()
              .find((id) => paragraphTextOf(alice.surface!.session.part(), id) === 'Tail')!;
          const first = await wave(alice, firstTarget);
          await pair.settle();
          const kids = (node: OoxmlNode) => (node.kind === 'textValue' ? [] : node.children),
            body = kids(bob.surface!.session.part().root).find((node) => node.kind === 'body')!,
            table = kids(body).find((node) => node.kind === 'table')!,
            row = kids(table)
              .filter((node) => node.kind === 'tableRow')
              .at(-1)!;
          pair.pause();
          await wave(alice, completeTarget('column'), first);
          append(bob, tail, ' Peer');
          expect(
            bob.surface!.applyAutomationOps(
              () =>
                owner === 'old-row'
                  ? [{ op: 'deleteTableRow', tableId: table.id, rowId: row.id }]
                  : [{ op: 'deleteBlock', blockId: table.id }],
              { kind: 'body' },
              [],
              false,
              true,
              'after'
            ).committed
          ).toBe(true);
          pair.reconnect(receiver);
          await pair.settle();
          const allOriginal = original.replace('Tail', 'Tail Peer'),
            allProposed = completeTarget('column').replace('Tail', 'Tail Peer'),
            removed = 'Lead\n\nTail Peer';
          const committed =
            owner === 'old-row' ? allOriginal.replace(/^\| D \| E \| F \|\n?/m, '') : removed;
          const proposed =
            owner === 'old-row'
              ? renderTarget([
                  [cell('Final **bold**', 2), cell('H3\n\nNew', 2)],
                  [cell('A'), cell('B'), cell('C'), cell('[Link](https://example.org/compound)')],
                ]).replace('Tail', 'Tail Peer')
              : removed;
          const prior =
            owner === 'old-row'
              ? renderTarget([
                  [cell('Title', 2), cell('H3')],
                  [cell('A'), cell('B'), cell('C')],
                ]).replace('Tail', 'Tail Peer')
              : removed;
          await check(pair, committed, proposed);
          await geometry(pair, tail);
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, committed, prior);
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, committed, proposed);
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, allOriginal, allProposed);
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, committed, proposed);
          await saved(pair);
        } finally {
          pair.destroy();
        }
      }, 30000);
