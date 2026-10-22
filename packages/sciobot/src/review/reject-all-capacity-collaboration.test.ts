import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { OoxmlNode, OoxmlPackage } from '@docx-editor.dev/core/store';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { seedReviewCapacity } from '../../test/support/review-capacity';
import { groupFixture, groups, groupToken } from '../../test/support/review-groups';
import {
  rejectedMixedPackage,
  bulkSource,
  payloadSnapshot,
  proposeMixedBulk,
} from '../../test/support/bulk-decisions';
import { groupDecisionFixture, decisionBlocks } from '../../test/support/group-decisions';
import { propertyContext } from '../../test/support/property-execution';
import { REVIEW_FLAG_TARGETS, REVIEW_FLAG_INTENTS } from '../collaboration/review-flag-codec';
import { planDomainProperties } from '../diff/domain-properties';
import { createEditPlan } from '../diff/plan';
import { trackPropertyPlan } from './track-properties';
import { buildBlockIndex } from '../blocks/block-index';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { rejectAllReviewGroups } from './reject-all';
const options = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
function ids(pkg: OoxmlPackage) {
  const result = new Set<string>(),
    visit = (node: OoxmlNode, partName: string) => {
      result.add(partName + ':' + node.id);
      if (node.kind !== 'textValue') for (const child of node.children) visit(child, partName);
    };
  for (const part of pkg.parts.values()) visit(part.root, part.name);
  return result;
}
for (const kind of ['targets', 'intents'] as const)
  for (const admit of [false, true])
    for (const receiver of [0, 1] as const)
      test(
        'bulk ' +
          kind +
          ' complete admission boundary is atomic and retains redo/history; admit=' +
          admit +
          '/' +
          receiver,
        async () => {
          const pair = await createTwoEditorFixture({
            document: writeOoxmlPackage(groupFixture()),
            ownedAttributes: true,
            clientIds: receiver ? [73, 41] : [41, 73],
            editor: options,
          });
          try {
            const [alice, bob] = pair.editors;
            proposeMixedBulk(alice!);
            await pair.settle();
            const surface = bob!.surface!,
              peer = surface.session
                .paragraphIds()
                .find((id) => paragraphTextOf(surface.session.part(), id) === 'Keep tail')!;
            surface.setSelection({
              anchor: { paragraphId: peer, offset: 0 },
              head: { paragraphId: peer, offset: 0 },
            });
            surface.insertPlainText('Peer: ');
            surface.flushPendingInput();
            await pair.settle();
            const a = alice!.surface!,
              end = paragraphTextOf(a.session.part(), peer)!.length;
            a.setSelection({
              anchor: { paragraphId: peer, offset: end },
              head: { paragraphId: peer, offset: end },
            });
            a.insertPlainText(' note');
            a.flushPendingInput();
            await pair.settle();
            expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            const pending = bulkSource(alice!),
              expected = rejectedMixedPackage(pending),
              after = ids(expected),
              footprint = [...ids(pending)].filter((id) => !after.has(id)).length,
              pendingIds = groups(alice!).map((group) => group.id);
            expect(footprint).toBeGreaterThan(1);
            expect(pendingIds).toHaveLength(6);
            seedReviewCapacity(pair.documents[0]!, kind, admit ? footprint : footprint - 1);
            await pair.settle();
            const before = pair.editors.map((editor) => ({
                bytes: writeOoxmlPackage(bulkSource(editor)),
                revision: editor.surface!.session.packageRevision(),
                payload: payloadSnapshot(bulkSource(editor)),
              })),
              vectors = pair.documents.map((doc) => Y.encodeStateVector(doc)),
              counts = pair.documents.map((doc) => [
                doc.getMap(REVIEW_FLAG_TARGETS).size,
                doc.getMap(REVIEW_FLAG_INTENTS).size,
              ]);
            pair.pause();
            const result = rejectAllReviewGroups(alice!, groupToken(alice!), propertyContext);
            if (admit) expect(result.success).toBe(true);
            else
              expect(result).toMatchObject({
                success: false,
                error: expect.stringContaining('record limit'),
              });
            pair.reconnect(receiver);
            await pair.settle();
            pair.editors.forEach((editor, index) => {
              expect(groups(editor).map((group) => group.id)).toEqual(admit ? [] : pendingIds);
              if (!admit) {
                expect(writeOoxmlPackage(bulkSource(editor))).toEqual(before[index]!.bytes);
                expect(editor.surface!.session.packageRevision()).toBe(before[index]!.revision);
                expect(payloadSnapshot(bulkSource(editor))).toEqual(before[index]!.payload);
                expect(Y.encodeStateVector(pair.documents[index]!)).toEqual(vectors[index]);
                expect([
                  pair.documents[index]!.getMap(REVIEW_FLAG_TARGETS).size,
                  pair.documents[index]!.getMap(REVIEW_FLAG_INTENTS).size,
                ]).toEqual(counts[index]);
              } else expect(payloadSnapshot(bulkSource(editor))).toEqual(payloadSnapshot(expected));
            });
            if (admit) {
              expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
              await pair.settle();
              for (const editor of pair.editors) {
                expect(groups(editor).map((group) => group.id)).toEqual(pendingIds);
                expect(payloadSnapshot(bulkSource(editor))).toEqual(payloadSnapshot(pending));
              }
              expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
              await pair.settle();
              for (const editor of pair.editors) expect(groups(editor)).toEqual([]);
            } else {
              expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
              await pair.settle();
              for (const editor of pair.editors) {
                expect(paragraphTextOf(editor.surface!.session.part(), peer)).toBe(
                  'Peer: Keep tail note'
                );
                expect(groups(editor).map((group) => group.id)).toEqual(pendingIds);
              }
            }
            const digests = await pair.savedDigests();
            expect(digests[0]).toEqual(digests[1]);
          } finally {
            pair.destroy();
          }
        },
        30000
      );

for (const kind of ['targets', 'intents'] as const)
  for (const receiver of [0, 1] as const)
    test(
      'metadata-only bulk rejection needs no removal capacity; ' + kind + '/' + receiver,
      async () => {
        const pair = await createTwoEditorFixture({
          document: writeOoxmlPackage(groupDecisionFixture()),
          ownedAttributes: true,
          clientIds: [41, 73],
          editor: options,
        });
        try {
          const alice = pair.editors[0]!,
            pkg = bulkSource(alice),
            revision = groupToken(alice),
            operations = decisionBlocks.flatMap((blockId) => {
              const planned = planDomainProperties(
                { pkg, revision, blockId },
                { set: { name: 'New ' + blockId, duration: 15 } }
              );
              if (!planned.success) throw new Error(planned.error);
              return planned.plan.operations;
            }),
            plan = createEditPlan(revision, operations);
          if (!plan.success) throw new Error(plan.error);
          expect(
            trackPropertyPlan(
              alice,
              plan.plan,
              decisionBlocks.map((blockId) => ({
                blockId,
                before: 'Old ' + blockId,
                after: 'Old ' + blockId,
              })),
              { ...propertyContext, proposalId: 'all-metadata' }
            ).success
          ).toBe(true);
          await pair.settle();
          const pending = groups(alice).map((group) => group.id),
            before = canonicalOoxmlFingerprint(alice.surface!.session.part());
          expect(pending).toHaveLength(3);
          seedReviewCapacity(pair.documents[0]!, kind, 0);
          await pair.settle();
          const counts = pair.documents.map((doc) => [
            doc.getMap(REVIEW_FLAG_TARGETS).size,
            doc.getMap(REVIEW_FLAG_INTENTS).size,
          ]);
          pair.pause();
          expect(rejectAllReviewGroups(alice, groupToken(alice), propertyContext).success).toBe(
            true
          );
          pair.reconnect(receiver);
          await pair.settle();
          pair.editors.forEach((editor, index) => {
            expect(groups(editor)).toEqual([]);
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
            for (const id of decisionBlocks)
              expect(
                buildBlockIndex(bulkSource(editor)).byId.get(id)!.payload.attributes.name
              ).toBe('Old ' + id);
            expect([
              pair.documents[index]!.getMap(REVIEW_FLAG_TARGETS).size,
              pair.documents[index]!.getMap(REVIEW_FLAG_INTENTS).size,
            ]).toEqual(counts[index]);
          });
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(groups(editor).map((group) => group.id)).toEqual(pending);
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
