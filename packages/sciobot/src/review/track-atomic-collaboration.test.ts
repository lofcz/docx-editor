import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  atomicFixture,
  atomicLabel,
  atomicListening,
  atomicVisual,
  executeAtomic,
  proposedListening,
  proposedVisual,
} from '../../test/support/atomic-execution';
import { propertyContext, propertyToken } from '../../test/support/property-execution';
import { buildBlockIndex } from '../blocks/block-index';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readPropertyReviewStore } from './property-record-store';
import { decideAtomicProposal } from './track-atomic';

function records(editor: DocxEditorInstance) {
  const read = readPropertyReviewStore(editor.surface!.session.currentPackage());
  if (read.status !== 'ready') throw new Error('Missing review store');
  return [...read.records.keys()].sort();
}
const payload = (editor: DocxEditorInstance, id: string) =>
  buildBlockIndex(editor.surface!.session.currentPackage()).byId.get(id)?.payload;
async function views(editor: DocxEditorInstance) {
  const token = propertyToken(editor),
    session = editor.surface!.session,
    [committed, working] = await Promise.all([
      readCommittedView(session, token),
      readWorkingView(session, token),
    ]);
  if (!committed.success) throw new Error(committed.error);
  if (!working.success) throw new Error(working.error);
  return [committed.view.projection.text, working.view.projection.text];
}
function insert(editor: DocxEditorInstance, paragraphId: string, offset: number, text: string) {
  const surface = editor.surface!;
  surface.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
  surface.insertPlainText(text);
  surface.flushPendingInput();
}

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`atomic ${decision} retains peer assets, text, history, geometry, reconnect, and saved convergence; clients ${clients.join('/')}, receiver ${receiver}`, async () => {
        const pair = await createTwoEditorFixture({
          document: writeOoxmlPackage(atomicFixture()),
          clientIds: clients,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
            tailId = alice
              .surface!.session.paragraphIds()
              .find((id) => paragraphTextOf(alice.surface!.session.part(), id) === 'Keep tail')!,
            phaseId = buildBlockIndex(alice.surface!.session.currentPackage()).byId.get(
              'phase'
            )!.controlId;
          bob.surface!.setSelection({
            anchor: { paragraphId: tailId, offset: 0 },
            head: { paragraphId: tailId, offset: 0 },
          });
          await pair.settle();
          pair.pause();
          const first = executeAtomic(alice, [proposedVisual], 'alice'),
            second = executeAtomic(bob, [proposedListening], 'bob');
          if (!first.success) throw new Error(first.error);
          if (!second.success) throw new Error(second.error);
          insert(bob, tailId, 0, 'Bob: ');
          pair.reconnect(receiver);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(records(editor)).toEqual(['alice', 'bob']);
            expect(payload(editor, 'visual')).toEqual(atomicVisual);
            expect(payload(editor, 'audio')).toEqual(atomicListening);
            const [committed, working] = await views(editor);
            expect(committed).toContain('src="vfs://pristine.png"');
            expect(committed).toContain('audio="original-audio"');
            expect(working).toContain('src="vfs://new.png"');
            expect(working).toContain('audio="new-audio"');
            expect(paragraphTextOf(editor.surface!.session.part(), tailId)).toBe('Bob: Keep tail');
          }
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          );
          expect(remote).toBeDefined();
          expect(remote!.head.offset).toBe(bob.surface!.state().selection!.head.offset);
          expect(
            caretAt(alice.surface!.layout(), {
              paragraphId: remote!.head.nodeId,
              offset: remote!.head.offset,
            })
          ).not.toBeNull();
          const pending = await pair.savedDigests();
          expect(pending[0]).toEqual(pending[1]);
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(records(editor)).toEqual(['bob']);
            expect(payload(editor, 'visual')).toEqual(atomicVisual);
            expect(paragraphTextOf(editor.surface!.session.part(), atomicLabel(editor).id)).toBe(
              'Old visual fallback'
            );
            expect(paragraphTextOf(editor.surface!.session.part(), tailId)).toBe('Bob: Keep tail');
          }
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          const decided = decideAtomicProposal(
            bob,
            'alice',
            decision,
            propertyToken(bob),
            propertyContext
          );
          if (!decided.success) throw new Error(decided.error);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(records(editor)).toEqual(['bob']);
            expect(payload(editor, 'visual')).toEqual(
              decision === 'accept' ? proposedVisual : atomicVisual
            );
            expect(payload(editor, 'audio')).toEqual(atomicListening);
            expect(paragraphTextOf(editor.surface!.session.part(), atomicLabel(editor).id)).toBe(
              `${decision === 'accept' ? 'New' : 'Old'} visual fallback`
            );
            expect(
              buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('phase')!.controlId
            ).toBe(phaseId);
          }
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(records(editor)).toEqual(['alice', 'bob']);
            expect(payload(editor, 'visual')).toEqual(atomicVisual);
            expect(paragraphTextOf(editor.surface!.session.part(), tailId)).toBe('Bob: Keep tail');
          }
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          const listeningDecision = decideAtomicProposal(
            alice,
            'bob',
            decision,
            propertyToken(alice),
            propertyContext
          );
          if (!listeningDecision.success) throw new Error(listeningDecision.error);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(records(editor)).toEqual([]);
            expect(payload(editor, 'audio')).toEqual(
              decision === 'accept' ? proposedListening : atomicListening
            );
            expect(paragraphTextOf(editor.surface!.session.part(), tailId)).toBe('Bob: Keep tail');
          }
          const final = await pair.savedDigests();
          expect(final[0]).toEqual(final[1]);
        } finally {
          pair.destroy();
        }
      }, 30_000);

for (const receiver of [0, 1] as const)
  test(`atomic owner deletion during a peer proposal retains assets without recreating the deleted control; receiver ${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(atomicFixture()),
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
        owner = buildBlockIndex(alice.surface!.session.currentPackage()).byId.get('visual')!;
      pair.pause();
      expect(
        alice.surface!.applyAutomationOps(
          () => [{ op: 'removeContentControl', controlId: owner.controlId, keepContent: false }],
          { kind: 'body' },
          undefined,
          false,
          true
        ).rejected
      ).toBe(false);
      expect(executeAtomic(bob).success).toBe(true);
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(payload(editor, 'visual')).toBeUndefined();
        expect(records(editor)).toEqual(['atomic-proposal']);
        expect(
          decideAtomicProposal(
            editor,
            'atomic-proposal',
            'accept',
            propertyToken(editor),
            propertyContext
          ).success
        ).toBe(false);
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(payload(editor, 'visual')).toEqual(atomicVisual);
        expect((await views(editor))[1]).toContain('src="vfs://new.png"');
      }
      expect(
        decideAtomicProposal(bob, 'atomic-proposal', 'reject', propertyToken(bob), propertyContext)
          .success
      ).toBe(true);
      await pair.settle();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  }, 30_000);

for (const decision of ['accept', 'reject'] as const)
  test(`atomic ${decision} and undo preserve peer text inside the fallback`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(atomicFixture()),
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      expect(executeAtomic(alice).success).toBe(true);
      await pair.settle();
      const id = atomicLabel(bob).id;
      insert(bob, id, paragraphTextOf(bob.surface!.session.part(), id)!.length, ' Bob');
      await pair.settle();
      const remote = pair.rooms[0]!.session.remoteSelections().find(
        (selection) => selection.actorId === 'bob'
      );
      expect(remote?.head.nodeId).toBe(id);
      expect(remote!.head.offset).toBe(bob.surface!.state().selection!.head.offset);
      expect(
        caretAt(alice.surface!.layout(), { paragraphId: id, offset: remote!.head.offset })
      ).not.toBeNull();
      const result = decideAtomicProposal(
        alice,
        'atomic-proposal',
        decision,
        propertyToken(alice),
        propertyContext
      );
      if (!result.success) throw new Error(result.error);
      await pair.settle();
      for (const editor of pair.editors)
        expect(paragraphTextOf(editor.surface!.session.part(), atomicLabel(editor).id)).toBe(
          `${decision === 'accept' ? 'New' : 'Old'} visual fallback Bob`
        );
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(records(editor)).toEqual(['atomic-proposal']);
        expect(payload(editor, 'visual')).toEqual(atomicVisual);
        expect(paragraphTextOf(editor.surface!.session.part(), atomicLabel(editor).id)).toContain(
          ' Bob'
        );
      }
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  }, 30_000);

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`concurrent replacements of one atomic block retain both asset pairs and require group reconciliation; clients ${clients.join('/')}`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(atomicFixture()),
      clientIds: clients,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
        peerTarget = {
          ...proposedVisual,
          attributes: { ...proposedVisual.attributes, src: 'vfs://peer.png', fileId: 'peer-file' },
        };
      pair.pause();
      expect(executeAtomic(alice, [proposedVisual], 'alice').success).toBe(true);
      expect(executeAtomic(bob, [peerTarget], 'bob').success).toBe(true);
      pair.reconnect(0);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(records(editor)).toEqual(['alice', 'bob']);
        expect(payload(editor, 'visual')).toEqual(atomicVisual);
        const store = readPropertyReviewStore(editor.surface!.session.currentPackage());
        if (store.status !== 'ready') throw new Error('Missing records');
        for (const [id, target] of [
          ['alice', proposedVisual],
          ['bob', peerTarget],
        ] as const) {
          const record = store.records.get(id)!.decoded;
          if (record.status !== 'supported') throw new Error('Invalid record');
          expect(record.value.blocks[0]!.changes[0]!.before).toEqual({
            present: true,
            value: atomicVisual,
          });
          expect(record.value.blocks[0]!.changes[0]!.after).toEqual({
            present: true,
            value: target,
          });
        }
        const before = writeOoxmlPackage(editor.surface!.session.currentPackage());
        expect(
          decideAtomicProposal(editor, 'alice', 'accept', propertyToken(editor), propertyContext)
            .success
        ).toBe(false);
        expect(writeOoxmlPackage(editor.surface!.session.currentPackage())).toEqual(before);
      }
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  }, 30_000);
