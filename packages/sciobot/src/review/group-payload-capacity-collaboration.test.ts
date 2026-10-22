import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  canonicalOoxmlFingerprint,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  acceptBlockGroup,
  groupDecisionFixture,
  proposeThreeProperties,
} from '../../test/support/group-decisions';
import { seedReviewCapacity } from '../../test/support/review-capacity';
import { groups, groupToken } from '../../test/support/review-groups';
import { propertyContext } from '../../test/support/property-execution';
import { rejectReviewGroup } from './reject-group';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { REVIEW_FLAG_INTENTS, REVIEW_FLAG_TARGETS } from '../collaboration/review-flag-codec';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

const editorOptions = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
function payloadSnapshot(pkg: OoxmlPackage) {
  const values = readPayloadStore(pkg);
  if (values.status !== 'ready') throw new Error('Missing payloads');
  return [...values.entries].map(([id, entry]) => [id, entry.raw, entry.label]);
}

for (const action of ['accept', 'reject'] as const)
  for (const kind of ['targets', 'intents'] as const)
    for (const receiver of [0, 1] as const)
      for (const metadataOnly of [false, true])
        test(`${action} complete payload previews preserve source state on capacity refusal and admit metadata-only writes at full ${kind} capacity; ${receiver}/metadata=${metadataOnly}`, async () => {
          const pair = await createTwoEditorFixture({
            document: writeOoxmlPackage(groupDecisionFixture()),
            ownedAttributes: true,
            clientIds: [41, 73],
            editor: editorOptions,
          });
          try {
            const alice = pair.editors[0]!;
            proposeThreeProperties(alice, metadataOnly);
            await pair.settle();
            const beforeGroups = groups(alice).map((group) => group.id);
            seedReviewCapacity(pair.documents[0]!, kind, metadataOnly ? 0 : 1);
            await pair.settle();
            const before = pair.editors.map((editor) => ({
                bytes: writeOoxmlPackage(editor.surface!.session.currentPackage()),
                revision: editor.surface!.session.packageRevision(),
                payloads: payloadSnapshot(editor.surface!.session.currentPackage()),
                word: canonicalOoxmlFingerprint(editor.surface!.session.part()),
              })),
              vectors = pair.documents.map((doc) => Y.encodeStateVector(doc)),
              rows = pair.documents.map((doc) => [
                doc.getMap(REVIEW_FLAG_TARGETS).size,
                doc.getMap(REVIEW_FLAG_INTENTS).size,
              ]);
            pair.pause();
            const result =
              action === 'accept'
                ? acceptBlockGroup(alice)
                : rejectReviewGroup(alice, beforeGroups[1]!, groupToken(alice), propertyContext);
            if (!metadataOnly)
              expect(result).toMatchObject({
                success: false,
                error: expect.stringContaining('record limit'),
              });
            else expect(result.success).toBe(true);
            pair.reconnect(receiver);
            await pair.settle();
            pair.editors.forEach((editor, index) => {
              const pkg = editor.surface!.session.currentPackage();
              expect(groups(editor).map((group) => group.id)).toEqual(
                metadataOnly ? [beforeGroups[0]!, beforeGroups[2]!] : beforeGroups
              );
              expect(buildBlockIndex(pkg).byId.get('middle')!.payload.attributes.name).toBe(
                metadataOnly && action === 'accept' ? 'New middle' : 'Old middle'
              );
              const payloads = readPayloadStore(pkg);
              expect(payloads.status === 'ready' && payloads.entries.get('middle')!.label).toBe(
                metadataOnly && action === 'accept' ? 'Old middle' : 'Stored middle'
              );
              if (action === 'reject')
                expect(payloadSnapshot(pkg)).toEqual(before[index]!.payloads);
              if (metadataOnly)
                expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
                  before[index]!.word
                );
              expect([
                pair.documents[index]!.getMap(REVIEW_FLAG_TARGETS).size,
                pair.documents[index]!.getMap(REVIEW_FLAG_INTENTS).size,
              ]).toEqual(rows[index]);
              if (!metadataOnly) {
                expect(writeOoxmlPackage(pkg)).toEqual(before[index]!.bytes);
                expect(editor.surface!.session.packageRevision()).toBe(before[index]!.revision);
                expect(Y.encodeStateVector(pair.documents[index]!)).toEqual(vectors[index]);
              }
            });
            if (metadataOnly) {
              expect(alice.exec({ type: 'undo' }).ok).toBe(true);
              await pair.settle();
              for (const editor of pair.editors)
                expect(groups(editor).map((group) => group.id)).toEqual(beforeGroups);
              expect(alice.exec({ type: 'redo' }).ok).toBe(true);
              await pair.settle();
            }
            const digests = await pair.savedDigests();
            expect(digests[0]).toEqual(digests[1]);
          } finally {
            pair.destroy();
          }
        }, 30000);
