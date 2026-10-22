import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  atomicFixture,
  atomicVisual,
  atomicListening,
  atomicLabel,
  executeAtomic,
  proposedVisual,
  proposedListening,
} from '../../test/support/atomic-execution';
import { groups, groupToken } from '../../test/support/review-groups';
import { propertyContext } from '../../test/support/property-execution';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { readPropertyReviewStore } from './property-record-store';
import { rejectReviewGroup } from './reject-group';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

const options = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
const source = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();
function stored(editor: DocxEditorInstance) {
  const payloads = readPayloadStore(source(editor)),
    records = readPropertyReviewStore(source(editor));
  if (payloads.status !== 'ready' || records.status !== 'ready') throw new Error('Missing stores');
  return {
    payloads: [...payloads.entries].map(([id, entry]) => [id, entry.raw, entry.label]),
    record: records.records.get('atomic-proposal')!.raw,
  };
}
function reject(editor: DocxEditorInstance, id: string) {
  const group = groups(editor).find((group) =>
    group.properties.some((member) => member.blockId === id)
  );
  if (!group) throw new Error('Missing atomic group');
  return rejectReviewGroup(editor, group.id, groupToken(editor), propertyContext);
}
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`independent atomic rejection retains original assets, peer text, history and room/DOCX reopening; ${clients}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: writeOoxmlPackage(atomicFixture()),
        clientIds: clients,
        ownedAttributes: true,
        editor: options,
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        expect(executeAtomic(alice, [proposedVisual, proposedListening]).success).toBe(true);
        await pair.settle();
        const ids = groups(alice).map((group) => group.id),
          before = stored(alice);
        expect(ids).toHaveLength(2);
        const tail = bob
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Keep tail')!;
        pair.pause();
        expect(reject(alice, 'visual').success).toBe(true);
        bob.surface!.setSelection({
          anchor: { paragraphId: tail, offset: 0 },
          head: { paragraphId: tail, offset: 0 },
        });
        bob.surface!.insertPlainText('Peer: ');
        bob.surface!.flushPendingInput();
        expect(reject(bob, 'audio').success).toBe(true);
        pair.reconnect(receiver);
        await pair.settle();
        const state = (pending: readonly string[]) => {
          for (const editor of pair.editors) {
            expect(groups(editor).map((group) => group.id)).toEqual([...pending]);
            expect(stored(editor)).toEqual(before);
            const index = buildBlockIndex(source(editor));
            expect(index.byId.get('visual')!.payload).toEqual(atomicVisual);
            expect(index.byId.get('audio')!.payload).toEqual(atomicListening);
            expect(paragraphTextOf(editor.surface!.session.part(), tail)).toBe('Peer: Keep tail');
            if (!pending.includes(ids[0]!))
              expect(
                paragraphTextOf(editor.surface!.session.part(), atomicLabel(editor, 'visual').id)
              ).toBe('Old visual fallback');
            if (!pending.includes(ids[1]!))
              expect(
                paragraphTextOf(editor.surface!.session.part(), atomicLabel(editor, 'audio').id)
              ).toBe('Old audio fallback');
          }
        };
        state([]);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        state([ids[0]!]);
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        state(ids);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        state([ids[1]!]);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        state([]);
        const carets = pair.editors.map((editor) =>
          caretAt(editor.surface!.layout(), { paragraphId: tail, offset: 6 })
        );
        expect(carets[0]).not.toBeNull();
        expect(carets[1]).not.toBeNull();
        expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({ x: carets[1]!.x, y: carets[1]!.y });
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (row) => row.actorId === 'bob'
        )!;
        expect(remote.head.nodeId).toBe(tail);
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        const bytes = new Uint8Array(await alice.save());
        const saved = readOoxmlPackage(bytes);
        if (!saved.ok) throw new Error(saved.reason);
        expect(buildBlockIndex(saved.package).byId.get('visual')!.payload).toEqual(atomicVisual);
        expect(buildBlockIndex(saved.package).byId.get('audio')!.payload).toEqual(atomicListening);
        for (const mode of [
          { document: bytes },
          {
            document: new Uint8Array(),
            roomSnapshot: Y.encodeStateAsUpdate(pair.documents[receiver]!),
          },
        ]) {
          const reopened = await createTwoEditorFixture({
            ...mode,
            ownedAttributes: true,
            editor: options,
          });
          try {
            await reopened.settle();
            for (const editor of reopened.editors) {
              expect(groups(editor)).toEqual([]);
              expect(buildBlockIndex(source(editor)).byId.get('visual')!.payload).toEqual(
                atomicVisual
              );
              expect(buildBlockIndex(source(editor)).byId.get('audio')!.payload).toEqual(
                atomicListening
              );
              expect(editor.surface!.session.bodyText()).toContain('Peer: Keep tail');
            }
          } finally {
            reopened.destroy();
          }
        }
        pair.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        pair.destroy();
      }
    }, 60000);
