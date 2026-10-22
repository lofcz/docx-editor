import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  groupFixture,
  groupParagraph,
  groups,
  groupToken,
  proposeGroupText,
} from '../../test/support/review-groups';
import { propertyEditor } from '../../test/support/property-execution';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readReviewGroupStore } from './group-record-store';
import type { NativeReviewGroup } from './groups';

const markdownText = (text: string) => text.replaceAll('.', '\\.');
const groupIds = (editor: DocxEditorInstance) => groups(editor).map((group) => group.id);
const wordIds = (group: NativeReviewGroup) => group.native.map((member) => member.address.id);
function decide(
  editor: DocxEditorInstance,
  group: NativeReviewGroup,
  op: 'acceptRevision' | 'rejectRevision'
) {
  const result = editor.surface!.applyAutomationOps(
    () =>
      group.native.map((member) => ({
        op,
        revision: member.address,
        siteNodeIds: member.siteNodeIds,
      })),
    { kind: 'body' },
    undefined,
    true,
    true
  );
  if (result.rejected || !result.committed) throw new Error(String(result.reason));
}
async function texts(editor: DocxEditorInstance) {
  const session = editor.surface!.session,
    token = groupToken(editor),
    original = await readCommittedView(session, token),
    working = await readWorkingView(session, token);
  if (!original.success || !working.success) throw new Error('Cannot read group views');
  return [original.view.projection.text, working.view.projection.text] as const;
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
    for (const action of ['acceptRevision', 'rejectRevision'] as const)
      test(`retired Word IDs survive same-paragraph proposals, cross-actor history, deletion, geometry, and reopening; clients ${clients.join('/')}, receiver ${receiver}, ${action}`, async () => {
        const document = writeOoxmlPackage(groupFixture()),
          pair = await createTwoEditorFixture({
            document,
            clientIds: clients,
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
            catId = groupParagraph(alice, 'The cat sat.'),
            cellId = groupParagraph(alice, 'A dog ran.');
          proposeGroupText(alice, 'The cat sat.', 'The car sat.');
          await pair.settle();
          const retired = groups(alice)[0]!;
          decide(bob, groups(bob)[0]!, action);
          await pair.settle();
          const base = action === 'acceptRevision' ? 'The car sat.' : 'The cat sat.',
            target = base.replace('sat', 'sit');
          expect(groupIds(alice)).toEqual([]);
          // Both persisted forms must retain the inactive address reservation before a later write.
          const saved = readOoxmlPackage(new Uint8Array(await alice.save()));
          if (!saved.ok) throw new Error(saved.reason);
          const solo = propertyEditor(saved.package);
          try {
            proposeGroupText(solo, base, target);
            const next = groups(solo)[0]!;
            expect(next.id).not.toBe(retired.id);
            expect(wordIds(next).some((id) => wordIds(retired).includes(id))).toBe(false);
          } finally {
            solo.destroy();
          }
          pair.flush();
          const reopened = await createTwoEditorFixture({
            document,
            roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
            clientIds: [211, 307],
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          });
          try {
            proposeGroupText(reopened.editors[0]!, base, target);
            await reopened.settle();
            const next = groups(reopened.editors[0]!)[0]!;
            expect(next.id).not.toBe(retired.id);
            expect(wordIds(next).some((id) => wordIds(retired).includes(id))).toBe(false);
            expect(groupIds(reopened.editors[1]!)).toEqual([next.id]);
            const digests = await reopened.savedDigests();
            expect(digests[0]).toEqual(digests[1]);
          } finally {
            reopened.destroy();
          }
          pair.pause();
          proposeGroupText(alice, base, target);
          const later = groups(alice)[0]!;
          proposeGroupText(bob, 'A dog ran.', 'A hen ran.');
          const peer = groups(bob)[0]!;
          bob.surface!.setSelection({
            anchor: { paragraphId: cellId, offset: 2 },
            head: { paragraphId: cellId, offset: 2 },
          });
          pair.reconnect(receiver);
          await pair.settle();
          expect(wordIds(later).some((id) => wordIds(retired).includes(id))).toBe(false);
          expect(wordIds(peer).some((id) => wordIds(retired).includes(id))).toBe(false);
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          );
          expect(remote?.head.nodeId).toBe(cellId);
          expect(remote!.head.offset).toBe(bob.surface!.state().selection!.head.offset);
          for (const editor of pair.editors) {
            expect(groupIds(editor)).toEqual([later.id, peer.id]);
            expect(groups(editor).every((group) => !group.readOnly)).toBe(true);
            expect(
              caretAt(editor.surface!.layout(), { paragraphId: cellId, offset: 2 })
            ).not.toBeNull();
            const [original, working] = await texts(editor);
            expect(original).toContain(markdownText(base));
            expect(working).toContain(markdownText(target));
          }
          // Bob removes his later unit, then undoes the older decision. Alice's later unit stays.
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(new Set(groupIds(editor))).toEqual(new Set([retired.id, later.id]));
            expect(groups(editor).every((group) => !group.readOnly)).toBe(true);
            const [original, working] = await texts(editor);
            expect(original).toContain(markdownText('The cat sat.'));
            expect(working).toContain(markdownText('The car sit.'));
          }
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(groupIds(editor)).toEqual([later.id, peer.id]);
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(groupIds(editor)).toEqual([peer.id]);
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          // A native edit races deletion in the affected paragraph. Undo restores peer text and group IDs.
          pair.pause();
          insert(alice, catId, 0, 'Peer: ');
          const removed = bob.surface!.applyAutomationOps(
            () => [{ op: 'deleteBlock', blockId: catId }],
            { kind: 'body' },
            undefined,
            false,
            true
          );
          expect(removed.rejected).toBe(false);
          pair.reconnect(receiver);
          await pair.settle();
          for (const editor of pair.editors) expect(groupIds(editor)).toEqual([peer.id]);
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(groupIds(editor)).toEqual([later.id, peer.id]);
            expect(paragraphTextOf(editor.surface!.session.part(), catId)).toStartWith('Peer: ');
            expect(
              caretAt(editor.surface!.layout(), { paragraphId: catId, offset: 2 })
            ).not.toBeNull();
            const store = readReviewGroupStore(editor.surface!.session.currentPackage());
            expect(store.status === 'ready' && store.retainedRecords.size).toBe(3);
            const loaded = readOoxmlPackage(new Uint8Array(await editor.save()));
            if (!loaded.ok) throw new Error(loaded.reason);
            const mounted = propertyEditor(loaded.package);
            try {
              expect(groupIds(mounted)).toEqual([later.id, peer.id]);
            } finally {
              mounted.destroy();
            }
          }
          const final = await pair.savedDigests();
          expect(final[0]).toEqual(final[1]);
        } finally {
          pair.destroy();
        }
      }, 60_000);

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`rapid independent text proposals keep separate undo units through peer input; clients ${clients.join('/')}`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(groupFixture()),
      clientIds: clients,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
        tail = groupParagraph(bob, 'Keep tail');
      await pair.settle();
      pair.pause();
      proposeGroupText(alice, 'The cat sat.', 'The car sat.');
      const first = groups(alice)[0]!.id;
      proposeGroupText(alice, 'A dog ran.', 'A hen ran.');
      const second = groups(alice)[1]!.id;
      insert(bob, tail, 0, 'Peer: ');
      pair.reconnect();
      await pair.settle();
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(groupIds(editor)).toEqual([first]);
        expect(paragraphTextOf(editor.surface!.session.part(), tail)).toBe('Peer: Keep tail');
      }
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) expect(groupIds(editor)).toEqual([first, second]);
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  }, 60_000);
