import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { groupFixture, groups, groupToken } from '../../test/support/review-groups';
import {
  acceptedMixedPackage,
  rejectedMixedPackage,
  bulkSource,
  payloadSnapshot,
  proposalSnapshot,
  proposeMixedBulk,
  typeBulkSentence,
} from '../../test/support/bulk-decisions';
import { atomicVisual } from '../../test/support/atomic-execution';
import { propertyContext } from '../../test/support/property-execution';
import { buildBlockIndex } from '../blocks/block-index';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { restoreReviewState } from './restore';

const options = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
const context = { ...propertyContext, epoch: 'restored-mount' };
function state(editor: DocxEditorInstance) {
  const result = restoreReviewState(bulkSource(editor), {
    ...groupToken(editor),
    epoch: context.epoch,
  });
  if (!result.success) throw new Error(result.error);
  return result.state;
}
function word(pkg: OoxmlPackage) {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw new Error(read.reason);
  return canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!);
}
function append(editor: DocxEditorInstance, paragraphId: string, text: string) {
  const surface = editor.surface!,
    offset = paragraphTextOf(surface.session.part(), paragraphId)!.length;
  surface.setSelection({
    anchor: { paragraphId, offset },
    head: { paragraphId, offset },
  });
  surface.insertPlainText(text);
  surface.flushPendingInput();
}

for (const mode of ['docx', 'room'] as const)
  for (const action of ['accept', 'reject'] as const)
    for (const clients of [
      [41, 73],
      [73, 41],
    ] as const)
      for (const receiver of [0, 1] as const)
        test(
          'restored ' +
            mode +
            ' review retains mixed decisions, peer edits, ' +
            action +
            ' history and deletion; ' +
            clients +
            '/' +
            receiver,
          async () => {
            const original = await createTwoEditorFixture({
              document: writeOoxmlPackage(groupFixture()),
              ownedAttributes: true,
              clientIds: clients,
              editor: options,
            });
            let reopened: Awaited<ReturnType<typeof createTwoEditorFixture>> | undefined;
            try {
              const alice = original.editors[0]!;
              proposeMixedBulk(alice, () => typeBulkSentence(alice));
              await original.settle();
              expect(
                acceptReviewGroup(alice, groups(alice)[0]!.id, groupToken(alice), propertyContext)
                  .success
              ).toBe(true);
              const visual = groups(alice).find((group) =>
                group.properties.some((member) => member.blockId === 'visual')
              )!;
              expect(
                rejectReviewGroup(alice, visual.id, groupToken(alice), propertyContext).success
              ).toBe(true);
              await original.settle();
              const pendingIds = state(alice).groups.map((group) => group.id),
                raw = proposalSnapshot(bulkSource(alice)),
                saved = new Uint8Array(await alice.save()),
                room = mode === 'room' ? Y.encodeStateAsUpdate(original.documents[0]!) : undefined;
              expect(pendingIds).toHaveLength(4);
              original.destroy();
              reopened = await createTwoEditorFixture({
                document: saved,
                roomSnapshot: room,
                ownedAttributes: true,
                clientIds: mode === 'room' ? [clients[0] + 100, clients[1] + 100] : clients,
                editor: options,
              });
              const [a, b] = reopened.editors as [DocxEditorInstance, DocxEditorInstance];
              for (const editor of reopened.editors) {
                const recovered = state(editor);
                expect(recovered.groups.map((group) => group.id)).toEqual(pendingIds);
                expect(recovered.reviewableGroupIds).toEqual(pendingIds);
                expect(recovered.pendingProperties).toHaveLength(2);
                expect(recovered.hasPendingReview).toBe(true);
                expect(editor.surface!.session.canUndo()).toBe(false);
                expect(editor.surface!.session.bodyText()).toContain('The car sat.');
              }
              const peer = b
                .surface!.session.paragraphIds()
                .find(
                  (id) =>
                    paragraphTextOf(b.surface!.session.part(), id) === 'Manual sentence. Keep tail'
                )!;
              reopened.pause();
              append(b, peer, ' offline peer');
              const pending = bulkSource(b);
              let expected =
                action === 'accept' ? acceptedMixedPackage(pending) : rejectedMixedPackage(pending);
              if (action === 'accept') {
                // The visual group was already rejected before saving. Bulk review must retain it.
                const originalVisual = withPayloadRecord(
                  expected,
                  atomicVisual,
                  'Stored visual label'
                );
                if (!originalVisual.ok) throw new Error(originalVisual.reason);
                expected = originalVisual.pkg;
              }
              const decide = action === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups;
              expect(decide(a, state(a).revision, context).success).toBe(true);
              reopened.reconnect(receiver);
              await reopened.settle();
              const check = (resolved: boolean) => {
                for (const editor of reopened!.editors) {
                  const recovered = state(editor);
                  expect(recovered.groups.map((group) => group.id)).toEqual(
                    resolved ? [] : pendingIds
                  );
                  expect(recovered.pendingProperties).toHaveLength(resolved ? 0 : 2);
                  expect(recovered.hasPendingReview).toBe(!resolved);
                  expect(word(bulkSource(editor))).toBe(word(resolved ? expected : pending));
                  expect(payloadSnapshot(bulkSource(editor))).toEqual(
                    payloadSnapshot(resolved ? expected : pending)
                  );
                  expect(proposalSnapshot(bulkSource(editor))).toEqual(raw);
                  expect(paragraphTextOf(editor.surface!.session.part(), peer)).toBe(
                    'Manual sentence. Keep tail offline peer'
                  );
                  expect(
                    caretAt(editor.surface!.layout(), { paragraphId: peer, offset: 0 })
                  ).not.toBeNull();
                }
              };
              check(true);
              const remote = reopened.rooms[0]!.session.remoteSelections().find(
                (selection) => selection.actorId === 'bob'
              );
              expect(remote).toBeDefined();
              expect(
                caretAt(a.surface!.layout(), {
                  paragraphId: remote!.head.nodeId,
                  offset: remote!.head.offset,
                })
              ).not.toBeNull();
              expect(a.exec({ type: 'undo' }).ok).toBe(true);
              await reopened.settle();
              check(false);
              expect(a.exec({ type: 'redo' }).ok).toBe(true);
              await reopened.settle();
              check(true);
              expect(b.exec({ type: 'undo' }).ok).toBe(true);
              await reopened.settle();
              for (const editor of reopened.editors) {
                expect(state(editor).hasPendingReview).toBe(false);
                expect(paragraphTextOf(editor.surface!.session.part(), peer)).toBe(
                  'Manual sentence. Keep tail'
                );
              }
              expect(b.exec({ type: 'redo' }).ok).toBe(true);
              await reopened.settle();
              check(true);
              const digests = await reopened.savedDigests();
              expect(digests[0]).toEqual(digests[1]);

              const owner = buildBlockIndex(bulkSource(a)).byId.get('activity')!,
                label = b
                  .surface!.session.paragraphIds()
                  .find(
                    (id) =>
                      paragraphTextOf(b.surface!.session.part(), id) ===
                      (action === 'accept'
                        ? 'New activity — 15 minutes'
                        : 'Old activity — 10 minutes')
                  )!;
              reopened.pause();
              append(b, label, ' unseen peer');
              expect(
                a.surface!.applyAutomationOps(() => [
                  { op: 'removeContentControl', controlId: owner.controlId, keepContent: false },
                ]).committed
              ).toBe(true);
              reopened.reconnect(receiver);
              await reopened.settle();
              for (const editor of reopened.editors) {
                expect(buildBlockIndex(bulkSource(editor)).byId.has('activity')).toBe(false);
                expect(state(editor)).toMatchObject({
                  hasPendingReview: false,
                  groups: [],
                  pendingProperties: [],
                  reviewableGroupIds: [],
                });
                const read = readOoxmlPackage(new Uint8Array(await editor.save()));
                if (!read.ok) throw new Error(read.reason);
                expect(restoreReviewState(read.package, state(editor).revision)).toMatchObject({
                  success: true,
                  state: { hasPendingReview: false, groups: [], pendingProperties: [] },
                });
              }
              const deleted = await reopened.savedDigests();
              expect(deleted[0]).toEqual(deleted[1]);
            } finally {
              original.destroy();
              reopened?.destroy();
            }
          },
          30000
        );
