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
import { externalMixedReviewDocument } from '../../test/support/external-review';
import {
  acceptedMixedPackage,
  bulkSource,
  payloadSnapshot,
  proposalSnapshot,
} from '../../test/support/bulk-decisions';
import { propertyContext, propertyLabel } from '../../test/support/property-execution';
import { groupToken, groups } from '../../test/support/review-groups';
import { buildBlockIndex } from '../blocks/block-index';
import { inspectExternalReview, reconcileExternalReview } from './reconcile-external';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

const options = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
function word(pkg: OoxmlPackage) {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw new Error(read.reason);
  return canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!);
}
function append(editor: DocxEditorInstance, paragraphId: string, text: string) {
  const surface = editor.surface!,
    offset = paragraphTextOf(surface.session.part(), paragraphId)!.length;
  surface.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
  surface.insertPlainText(text);
  surface.flushPendingInput();
}
function state(editor: DocxEditorInstance) {
  const read = inspectExternalReview(bulkSource(editor));
  if (!read.success) throw new Error(read.error);
  return read.report;
}

for (const mode of ['same-reconcile', 'peer-fallback-edit', 'owner-deletion'] as const)
  for (const decision of ['accept', 'reject'] as const)
    for (const clients of [
      [41, 73],
      [73, 41],
    ] as const)
      for (const receiver of [0, 1] as const)
        test(
          'external ' +
            decision +
            ' metadata retains Word, peer input, histories and reopening; ' +
            mode +
            '/' +
            clients +
            '/' +
            receiver,
          async () => {
            const pair = await createTwoEditorFixture({
              document: await externalMixedReviewDocument(decision),
              ownedAttributes: true,
              clientIds: clients,
              editor: options,
            });
            try {
              const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
                initial = bulkSource(alice),
                initialWord = word(initial),
                raw = proposalSnapshot(initial),
                beforePayloads = payloadSnapshot(initial),
                afterPayloads =
                  decision === 'accept'
                    ? payloadSnapshot(acceptedMixedPackage(initial))
                    : beforePayloads,
                tail = bob
                  .surface!.session.paragraphIds()
                  .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Keep tail')!,
                label = propertyLabel(bob).id;
              const beforeVector = pair.documents.map((doc) => Y.encodeStateVector(doc));
              for (const editor of pair.editors) {
                const read = state(editor);
                expect(read.hasConflicts).toBe(false);
                expect(read.entries.map((entry) => entry.status)).toEqual(
                  Array(3).fill(decision === 'accept' ? 'accepted' : 'rejected')
                );
              }
              pair.documents.forEach((doc, index) =>
                expect(Y.encodeStateVector(doc)).toEqual(beforeVector[index])
              );
              pair.pause();
              expect(
                reconcileExternalReview(alice, groupToken(alice), propertyContext)
              ).toMatchObject({ success: true, reconciled: expect.any(Array) });
              append(bob, mode === 'same-reconcile' ? tail : label, ' Peer input');
              const withPeer = bulkSource(bob),
                peerWord = word(withPeer);
              let expectedWord = peerWord;
              if (mode === 'same-reconcile')
                expect(reconcileExternalReview(bob, groupToken(bob), propertyContext).success).toBe(
                  true
                );
              if (mode === 'owner-deletion') {
                const controlId = buildBlockIndex(bulkSource(bob)).byId.get('activity')!.controlId;
                expect(
                  bob.surface!.applyAutomationOps(() => [
                    { op: 'removeContentControl', controlId, keepContent: false },
                  ]).committed
                ).toBe(true);
                expectedWord = word(bulkSource(bob));
              }
              pair.reconnect(receiver);
              await pair.settle();
              const check = (
                resolved: boolean,
                expected: string,
                ownerPresent = true,
                edited = true
              ) => {
                for (const editor of pair.editors) {
                  expect(word(bulkSource(editor))).toBe(expected);
                  expect(payloadSnapshot(bulkSource(editor))).toEqual(
                    resolved ? afterPayloads : beforePayloads
                  );
                  expect(proposalSnapshot(bulkSource(editor))).toEqual(raw);
                  expect(buildBlockIndex(bulkSource(editor)).byId.has('activity')).toBe(
                    ownerPresent
                  );
                  const report = state(editor);
                  for (const entry of report.entries)
                    expect(entry.status).toBe(
                      resolved
                        ? 'resolved'
                        : mode !== 'same-reconcile' && edited && entry.blockId === 'activity'
                          ? 'conflict'
                          : decision === 'accept'
                            ? 'accepted'
                            : 'rejected'
                    );
                  expect(groups(editor)).toHaveLength(ownerPresent ? (resolved ? 3 : 6) : 1);
                }
              };
              check(true, expectedWord, mode !== 'owner-deletion');
              if (mode === 'same-reconcile') {
                // Both first-undo actors and two complete cycles retain independent metadata ownership.
                for (const order of [
                  [0, 1],
                  [1, 0],
                ] as const) {
                  expect(pair.editors[order[0]]!.exec({ type: 'undo' }).ok).toBe(true);
                  await pair.settle();
                  check(true, peerWord);
                  expect(pair.editors[order[1]]!.exec({ type: 'undo' }).ok).toBe(true);
                  await pair.settle();
                  check(false, peerWord);
                  expect(pair.editors[order[0]]!.exec({ type: 'redo' }).ok).toBe(true);
                  await pair.settle();
                  check(true, peerWord);
                  expect(pair.editors[order[1]]!.exec({ type: 'redo' }).ok).toBe(true);
                  await pair.settle();
                  check(true, peerWord);
                }
              } else if (mode === 'peer-fallback-edit') {
                expect(alice.exec({ type: 'undo' }).ok).toBe(true);
                await pair.settle();
                check(false, peerWord);
                expect(alice.exec({ type: 'redo' }).ok).toBe(true);
                await pair.settle();
                check(true, peerWord);
                expect(bob.exec({ type: 'undo' }).ok).toBe(true);
                await pair.settle();
                check(true, initialWord, true, false);
                expect(bob.exec({ type: 'redo' }).ok).toBe(true);
                await pair.settle();
                check(true, peerWord);
              } else {
                expect(bob.exec({ type: 'undo' }).ok).toBe(true);
                await pair.settle();
                check(true, peerWord);
                expect(alice.exec({ type: 'undo' }).ok).toBe(true);
                await pair.settle();
                check(false, peerWord);
                expect(bob.exec({ type: 'undo' }).ok).toBe(true);
                await pair.settle();
                check(false, initialWord, true, false);
                expect(alice.exec({ type: 'redo' }).ok).toBe(true);
                await pair.settle();
                check(true, initialWord, true, false);
                expect(bob.exec({ type: 'redo' }).ok).toBe(true);
                await pair.settle();
                check(true, peerWord);
                expect(bob.exec({ type: 'redo' }).ok).toBe(true);
                await pair.settle();
                check(true, expectedWord, false);
              }
              bob.surface!.setSelection({
                anchor: { paragraphId: tail, offset: 0 },
                head: { paragraphId: tail, offset: 0 },
              });
              await pair.settle();
              const remote = pair.rooms[0]!.session.remoteSelections().find(
                (selection) => selection.actorId === 'bob'
              );
              expect(remote).toBeDefined();
              for (const editor of pair.editors)
                expect(
                  caretAt(editor.surface!.layout(), {
                    paragraphId: remote!.head.nodeId,
                    offset: remote!.head.offset,
                  })
                ).not.toBeNull();
              const digests = await pair.savedDigests();
              expect(digests[0]).toEqual(digests[1]);
              const bytes = new Uint8Array(await alice.save());
              for (const snapshot of [
                undefined,
                Y.encodeStateAsUpdate(pair.documents[0]!),
              ] as const) {
                const reopened = await createTwoEditorFixture({
                  document: bytes,
                  roomSnapshot: snapshot,
                  ownedAttributes: true,
                  clientIds: [clients[0] + 100, clients[1] + 100],
                  editor: options,
                });
                try {
                  for (const editor of reopened.editors) {
                    expect(word(bulkSource(editor))).toBe(expectedWord);
                    expect(payloadSnapshot(bulkSource(editor))).toEqual(afterPayloads);
                    expect(
                      state(editor).entries.every((entry) => entry.status === 'resolved')
                    ).toBe(true);
                  }
                } finally {
                  reopened.destroy();
                }
              }
            } finally {
              pair.destroy();
            }
          },
          30000
        );
