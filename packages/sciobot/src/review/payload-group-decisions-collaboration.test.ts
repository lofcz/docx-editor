import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  decisionFixture,
  decisionReceipt,
  decisionState,
  decisionStates,
  multiBlockProposal,
  mutateDecision,
} from '../../test/support/payload-group-decisions';
import { propertyActivity } from '../../test/support/property-record';
import { readPropertyReviewStore } from './property-record-store';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

const source = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();
function originalRecord(editor: DocxEditorInstance) {
  const store = readPropertyReviewStore(source(editor));
  if (store.status !== 'ready') throw new Error('Missing property store');
  return store.records.get(multiBlockProposal.id)!.raw;
}
function paragraph(editor: DocxEditorInstance, text: string) {
  const session = editor.surface!.session,
    id = session.paragraphIds().find((id) => paragraphTextOf(session.part(), id) === text);
  if (!id) throw new Error('Missing ' + text);
  return id;
}
function type(editor: DocxEditorInstance, id: string, text: string) {
  const surface = editor.surface!;
  surface.setSelection({
    anchor: { paragraphId: id, offset: 0 },
    head: { paragraphId: id, offset: 0 },
  });
  surface.insertPlainText(text);
  surface.flushPendingInput();
}
function savedStates(bytes: Uint8Array) {
  const read = readOoxmlPackage(bytes);
  if (!read.ok) throw new Error(read.reason);
  return decisionStates(read.package);
}
const editorOptions = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`independent block receipts preserve shared proposal JSON, peer text, history, geometry, and reopening; clients ${clients.join('/')}, receiver ${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: writeOoxmlPackage(decisionFixture()),
        clientIds: clients,
        editor: editorOptions,
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          id = paragraph(alice, 'Body'),
          original = originalRecord(alice),
          geometry = caretAt(alice.surface!.layout(), { paragraphId: id, offset: 0 });
        expect(geometry).not.toBeNull();
        pair.pause();
        mutateDecision(alice, decisionReceipt());
        expect(caretAt(alice.surface!.layout(), { paragraphId: id, offset: 0 })).toEqual(geometry);
        mutateDecision(bob, decisionReceipt('rg_bob', 'phase', 'reject'));
        type(bob, id, 'Peer: ');
        pair.reconnect(receiver);
        await pair.settle();
        const expected: ReturnType<typeof decisionStates> = [
          ['activity', 'accepted', ['rg_alice']],
          ['phase', 'rejected', ['rg_bob']],
        ];
        for (const editor of pair.editors) {
          expect(decisionStates(source(editor))).toEqual(expected);
          expect(originalRecord(editor)).toBe(original);
          expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Peer: Body');
        }
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (value) => value.actorId === 'bob'
        );
        expect(remote).toBeDefined();
        expect(
          caretAt(alice.surface!.layout(), {
            paragraphId: remote!.head.nodeId,
            offset: remote!.head.offset,
          })
        ).not.toBeNull();
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(decisionStates(source(editor))).toEqual([['phase', 'rejected', ['rg_bob']]]);
          expect(originalRecord(editor)).toBe(original);
          expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Peer: Body');
        }
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(decisionStates(source(editor))).toEqual(expected);
          expect(savedStates(new Uint8Array(await editor.save()))).toEqual(expected);
        }
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        const snapshot = Y.encodeStateAsUpdate(pair.documents[receiver]!);
        const reopened = await createTwoEditorFixture({
          document: new Uint8Array(),
          roomSnapshot: snapshot,
          editor: editorOptions,
        });
        try {
          await reopened.settle();
          for (const editor of reopened.editors) {
            expect(decisionStates(source(editor))).toEqual(expected);
            expect(originalRecord(editor)).toBe(original);
            expect(
              paragraphTextOf(editor.surface!.session.part(), paragraph(editor, 'Peer: Body'))
            ).toBe('Peer: Body');
          }
        } finally {
          reopened.destroy();
        }
      } finally {
        pair.destroy();
      }
    }, 30_000);

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const peerDecision of ['accept', 'reject'] as const)
      test(`same-group ${peerDecision} receipts retain both actors and preserve conflict through history; clients ${clients.join('/')}, receiver ${receiver}`, async () => {
        const pair = await createTwoEditorFixture({
          document: writeOoxmlPackage(decisionFixture()),
          clientIds: clients,
          editor: editorOptions,
        });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
            original = originalRecord(alice),
            status = peerDecision === 'accept' ? 'accepted' : 'conflict';
          pair.pause();
          mutateDecision(alice, decisionReceipt());
          mutateDecision(bob, decisionReceipt('rg_bob', 'activity', peerDecision));
          pair.reconnect(receiver);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(decisionState(source(editor))).toMatchObject({
              status,
              receiptIds: ['rg_alice', 'rg_bob'],
            });
            expect(originalRecord(editor)).toBe(original);
          }
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(decisionState(source(editor))).toMatchObject({
              status: peerDecision === 'accept' ? 'accepted' : 'rejected',
              receiptIds: ['rg_bob'],
            });
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(decisionState(source(editor))).toMatchObject({
              status: 'accepted',
              receiptIds: ['rg_alice'],
            });
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(decisionState(source(editor))).toMatchObject({
              status,
              receiptIds: ['rg_alice', 'rg_bob'],
            });
            expect(savedStates(new Uint8Array(await editor.save()))).toEqual([
              ['activity', status, ['rg_alice', 'rg_bob']],
            ]);
          }
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
        } finally {
          pair.destroy();
        }
      }, 30_000);

for (const receiver of [0, 1] as const)
  test(`owner deletion preserves a peer receipt and restores peer edits inside the deleted owner on undo; receiver ${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(decisionFixture()),
      clientIds: [41, 73],
      editor: editorOptions,
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
        owner = propertyActivity(source(alice)),
        id = paragraph(alice, 'Body'),
        original = originalRecord(alice);
      pair.pause();
      const deleted = alice.surface!.applyAutomationOps(
        () => [{ op: 'removeContentControl', controlId: owner.controlId, keepContent: false }],
        { kind: 'body' },
        undefined,
        false,
        true
      );
      expect(deleted.committed).toBe(true);
      mutateDecision(bob, decisionReceipt('rg_bob'));
      type(bob, id, 'Peer: ');
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(propertyActivity(source(editor))).toBeUndefined();
        expect(decisionState(source(editor))).toMatchObject({
          status: 'accepted',
          receiptIds: ['rg_bob'],
        });
        expect(originalRecord(editor)).toBe(original);
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(propertyActivity(source(editor))).toBeDefined();
        expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Peer: Body');
        expect(decisionState(source(editor))).toMatchObject({
          status: 'accepted',
          receiptIds: ['rg_bob'],
        });
        expect(caretAt(editor.surface!.layout(), { paragraphId: id, offset: 0 })).not.toBeNull();
      }
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(propertyActivity(source(editor))).toBeUndefined();
        expect(decisionState(source(editor))).toMatchObject({
          status: 'accepted',
          receiptIds: ['rg_bob'],
        });
      }
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
    } finally {
      pair.destroy();
    }
  }, 30_000);
