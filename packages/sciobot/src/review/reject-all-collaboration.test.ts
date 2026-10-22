import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  groupFixture,
  groups,
  groupToken,
  proposeGroupText,
} from '../../test/support/review-groups';
import {
  rejectedMixedPackage,
  bulkSource,
  payloadSnapshot,
  proposalSnapshot,
  proposeMixedBulk,
  typeBulkSentence,
} from '../../test/support/bulk-decisions';
import { propertyContext } from '../../test/support/property-execution';
import { buildBlockIndex } from '../blocks/block-index';
import { rejectAllReviewGroups } from './reject-all';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

const options = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
function word(pkg: OoxmlPackage) {
  const before = canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw new Error(read.reason);
  expect(canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!)).toBe(before);
  return canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!);
}
function type(editor: DocxEditorInstance) {
  const surface = editor.surface!,
    id = surface.session
      .paragraphIds()
      .find((id) => paragraphTextOf(surface.session.part(), id) === 'Manual sentence. Keep tail')!;
  surface.setSelection({
    anchor: { paragraphId: id, offset: 0 },
    head: { paragraphId: id, offset: 0 },
  });
  surface.insertPlainText('Peer: ');
  surface.flushPendingInput();
  return id;
}
function check(
  pair: Awaited<ReturnType<typeof createTwoEditorFixture>>,
  ids: readonly string[],
  expected: OoxmlPackage,
  raw: ReturnType<typeof proposalSnapshot>
) {
  for (const editor of pair.editors) {
    expect(groups(editor).map((group) => group.id)).toEqual([...ids]);
    expect(word(bulkSource(editor))).toBe(word(expected));
    expect(payloadSnapshot(bulkSource(editor))).toEqual(payloadSnapshot(expected));
    expect(proposalSnapshot(bulkSource(editor))).toEqual(raw);
  }
}
async function saved(
  pair: Awaited<ReturnType<typeof createTwoEditorFixture>>,
  ids: readonly string[],
  expected: OoxmlPackage
) {
  const digests = await pair.savedDigests();
  expect(digests[0]).toEqual(digests[1]);
  const reopened = await createTwoEditorFixture({
    document: new Uint8Array(await pair.editors[0]!.save()),
    roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
    ownedAttributes: true,
    editor: options,
  });
  try {
    for (const editor of reopened.editors) {
      expect(groups(editor).map((group) => group.id)).toEqual([...ids]);
      expect(word(bulkSource(editor))).toBe(word(expected));
      expect(payloadSnapshot(bulkSource(editor))).toEqual(payloadSnapshot(expected));
    }
  } finally {
    reopened.destroy();
  }
}
for (const mode of ['same-bulk', 'later-group', 'owner-deletion'] as const)
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      test(
        'bulk rejection ' +
          mode +
          ' retains peer text, both actors history, geometry and room/DOCX state; ' +
          clients +
          '/' +
          receiver,
        async () => {
          const pair = await createTwoEditorFixture({
            document: writeOoxmlPackage(groupFixture()),
            ownedAttributes: true,
            clientIds: clients,
            editor: options,
          });
          try {
            const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
            proposeMixedBulk(alice, () => typeBulkSentence(alice));
            await pair.settle();
            const original = groups(alice),
              ids = original.map((group) => group.id),
              raw = proposalSnapshot(bulkSource(alice)),
              selected = new Map<string, string[]>();
            for (const group of original)
              for (const member of group.native)
                selected.set(member.partName, [
                  ...(selected.get(member.partName) ?? []),
                  ...member.siteNodeIds,
                ]);
            pair.pause();
            expect(rejectAllReviewGroups(alice, groupToken(alice), propertyContext).success).toBe(
              true
            );
            const peer = type(bob),
              pendingPeer = bulkSource(bob);
            let expected = rejectedMixedPackage(pendingPeer),
              finalIds: readonly string[] = [];
            const visual = buildBlockIndex(pendingPeer).byId.get('visual')!.controlId;
            if (mode === 'same-bulk')
              expect(rejectAllReviewGroups(bob, groupToken(bob), propertyContext).success).toBe(
                true
              );
            else if (mode === 'later-group') {
              proposeGroupText(
                bob,
                'Peer: Manual sentence. Keep tail',
                'Peer: Manual sentence. Keep tail later.'
              );
              finalIds = groups(bob)
                .map((group) => group.id)
                .filter((id) => !ids.includes(id));
              expect(finalIds).toHaveLength(1);
              expected = rejectedMixedPackage(bulkSource(bob), selected);
            } else {
              const removed = applyTreeOp(expected.parts.get(expected.mainDocumentPart)!, {
                op: 'removeContentControl',
                controlId: visual,
                keepContent: false,
              });
              if (!removed.ok) throw new Error(removed.reason);
              expected = withPart(expected, removed.part);
              expect(
                bob.surface!.applyAutomationOps(
                  () => [{ op: 'removeContentControl', controlId: visual, keepContent: false }],
                  { kind: 'body' },
                  undefined,
                  false,
                  true
                ).committed
              ).toBe(true);
            }
            const pendingLater = mode === 'later-group' ? bulkSource(bob) : pendingPeer;
            const pendingLaterIds = groups(bob).map((group) => group.id);
            pair.reconnect(receiver);
            await pair.settle();
            check(pair, finalIds, expected, raw);
            for (const editor of pair.editors) {
              expect(paragraphTextOf(editor.surface!.session.part(), peer)).toBe(
                mode === 'later-group'
                  ? 'Peer: Manual sentence. Keep tail later.'
                  : 'Peer: Manual sentence. Keep tail'
              );
              expect(
                caretAt(editor.surface!.layout(), { paragraphId: peer, offset: 0 })
              ).not.toBeNull();
            }
            const remote = pair.rooms[0]!.session.remoteSelections().find(
              (row) => row.actorId === 'bob'
            );
            expect(remote).toBeDefined();
            expect(
              caretAt(alice.surface!.layout(), {
                paragraphId: remote!.head.nodeId,
                offset: remote!.head.offset,
              })
            ).not.toBeNull();
            if (mode === 'same-bulk') {
              expect(alice.exec({ type: 'undo' }).ok).toBe(true);
              await pair.settle();
              check(pair, [], expected, raw);
              expect(bob.exec({ type: 'undo' }).ok).toBe(true);
              await pair.settle();
              check(pair, ids, pendingPeer, raw);
              expect(alice.exec({ type: 'redo' }).ok).toBe(true);
              await pair.settle();
              check(pair, [], expected, raw);
              expect(bob.exec({ type: 'redo' }).ok).toBe(true);
              await pair.settle();
              check(pair, [], expected, raw);
            } else if (mode === 'later-group') {
              expect(alice.exec({ type: 'undo' }).ok).toBe(true);
              await pair.settle();
              check(pair, pendingLaterIds, pendingLater, raw);
              expect(bob.exec({ type: 'undo' }).ok).toBe(true);
              await pair.settle();
              check(pair, ids, pendingPeer, raw);
              expect(alice.exec({ type: 'redo' }).ok).toBe(true);
              await pair.settle();
              check(pair, [], rejectedMixedPackage(pendingPeer), raw);
              expect(bob.exec({ type: 'redo' }).ok).toBe(true);
              await pair.settle();
              check(pair, finalIds, expected, raw);
            } else {
              expect(bob.exec({ type: 'undo' }).ok).toBe(true);
              await pair.settle();
              check(pair, [], rejectedMixedPackage(pendingPeer), raw);
              expect(alice.exec({ type: 'undo' }).ok).toBe(true);
              await pair.settle();
              check(pair, ids, pendingPeer, raw);
              expect(alice.exec({ type: 'redo' }).ok).toBe(true);
              await pair.settle();
              check(pair, [], rejectedMixedPackage(pendingPeer), raw);
              expect(bob.exec({ type: 'redo' }).ok).toBe(true);
              await pair.settle();
              check(pair, [], expected, raw);
              for (const editor of pair.editors)
                expect(buildBlockIndex(bulkSource(editor)).byId.get('visual')).toBeUndefined();
            }
            await saved(pair, finalIds, expected);
          } finally {
            pair.destroy();
          }
        },
        30000
      );
