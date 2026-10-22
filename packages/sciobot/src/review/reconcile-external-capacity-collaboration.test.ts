import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { externalMixedReviewDocument } from '../../test/support/external-review';
import { seedReviewFieldCapacity } from '../../test/support/review-field-capacity';
import { seedReviewCapacity } from '../../test/support/review-capacity';
import { bulkSource, payloadSnapshot, proposalSnapshot } from '../../test/support/bulk-decisions';
import { groupToken } from '../../test/support/review-groups';
import { propertyContext } from '../../test/support/property-execution';
import { REVIEW_FIELD_TARGETS, REVIEW_FIELD_INTENTS } from '../collaboration/review-field-codec';
import { REVIEW_FLAG_TARGETS, REVIEW_FLAG_INTENTS } from '../collaboration/review-flag-codec';
import { inspectExternalReview, reconcileExternalReview } from './reconcile-external';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

const options = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
for (const kind of ['targets', 'intents'] as const)
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      test(
        'complete external metadata ' +
          kind +
          ' admission refuses before mutation and retains prior redo; ' +
          clients +
          '/' +
          receiver,
        async () => {
          const pair = await createTwoEditorFixture({
            document: await externalMixedReviewDocument('accept'),
            ownedAttributes: true,
            clientIds: clients,
            editor: options,
          });
          try {
            const alice = pair.editors[0]!,
              surface = alice.surface!,
              tail = surface.session
                .paragraphIds()
                .find((id) => paragraphTextOf(surface.session.part(), id) === 'Keep tail')!;
            surface.setSelection({
              anchor: { paragraphId: tail, offset: 9 },
              head: { paragraphId: tail, offset: 9 },
            });
            surface.insertPlainText(' retained');
            surface.flushPendingInput();
            await pair.settle();
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            seedReviewFieldCapacity(pair.documents[0]!, kind, 1);
            await pair.settle();
            const before = pair.editors.map((editor) => ({
                bytes: writeOoxmlPackage(bulkSource(editor)),
                revision: editor.surface!.session.packageRevision(),
                payloads: payloadSnapshot(bulkSource(editor)),
                records: proposalSnapshot(bulkSource(editor)),
              })),
              vectors = pair.documents.map((doc) => Y.encodeStateVector(doc)),
              counts = pair.documents.map((doc) => [
                doc.getMap(REVIEW_FIELD_TARGETS).size,
                doc.getMap(REVIEW_FIELD_INTENTS).size,
              ]),
              history = pair.rooms.map((room) => [room.session.canUndo(), room.session.canRedo()]);
            expect(history[0]![1]).toBe(true);
            pair.pause();
            for (const editor of pair.editors)
              expect(
                reconcileExternalReview(editor, groupToken(editor), propertyContext)
              ).toMatchObject({ success: false, error: expect.stringContaining('record limit') });
            pair.reconnect(receiver);
            await pair.settle();
            pair.editors.forEach((editor, index) => {
              expect(writeOoxmlPackage(bulkSource(editor))).toEqual(before[index]!.bytes);
              expect(editor.surface!.session.packageRevision()).toBe(before[index]!.revision);
              expect(payloadSnapshot(bulkSource(editor))).toEqual(before[index]!.payloads);
              expect(proposalSnapshot(bulkSource(editor))).toEqual(before[index]!.records);
              expect(Y.encodeStateVector(pair.documents[index]!)).toEqual(vectors[index]);
              expect([
                pair.documents[index]!.getMap(REVIEW_FIELD_TARGETS).size,
                pair.documents[index]!.getMap(REVIEW_FIELD_INTENTS).size,
              ]).toEqual(counts[index]);
            });
            expect(
              pair.rooms.map((room) => [room.session.canUndo(), room.session.canRedo()])
            ).toEqual(history);
            expect(alice.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors)
              expect(paragraphTextOf(editor.surface!.session.part(), tail)).toBe(
                'Keep tail retained'
              );
          } finally {
            pair.destroy();
          }
        },
        30000
      );

for (const kind of ['targets', 'intents'] as const)
  for (const receiver of [0, 1] as const)
    test(
      'external metadata reconciliation at full removal ' +
        kind +
        ' capacity uses no removed Word nodes; ' +
        receiver,
      async () => {
        const pair = await createTwoEditorFixture({
          document: await externalMixedReviewDocument('accept'),
          ownedAttributes: true,
          clientIds: receiver ? [73, 41] : [41, 73],
          editor: options,
        });
        try {
          const alice = pair.editors[0]!,
            before = canonicalOoxmlFingerprint(alice.surface!.session.part());
          seedReviewCapacity(pair.documents[0]!, kind, 0);
          await pair.settle();
          const counts = pair.documents.map((doc) => [
            doc.getMap(REVIEW_FLAG_TARGETS).size,
            doc.getMap(REVIEW_FLAG_INTENTS).size,
          ]);
          pair.pause();
          expect(reconcileExternalReview(alice, groupToken(alice), propertyContext)).toMatchObject({
            success: true,
            reconciled: expect.any(Array),
          });
          pair.reconnect(receiver);
          await pair.settle();
          pair.editors.forEach((editor, index) => {
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
            expect(inspectExternalReview(bulkSource(editor))).toMatchObject({
              success: true,
              report: {
                entries: expect.arrayContaining([
                  {
                    proposalId: 'property-proposal',
                    blockId: 'activity',
                    status: 'resolved',
                    conflict: null,
                  },
                ]),
              },
            });
            expect([
              pair.documents[index]!.getMap(REVIEW_FLAG_TARGETS).size,
              pair.documents[index]!.getMap(REVIEW_FLAG_INTENTS).size,
            ]).toEqual(counts[index]);
          });
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
        } finally {
          pair.destroy();
        }
      },
      30000
    );
