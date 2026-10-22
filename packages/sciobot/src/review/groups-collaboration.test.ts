import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { Awareness } from 'y-protocols/awareness';
import { createDocumentCollaboration } from '../../../pro/src/collaboration/document-session';
import { readReviewGroupStore } from './group-record-store';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  groupFixture,
  groupParagraph,
  groups,
  insertLeadingGroup,
  proposeGroupText,
} from '../../test/support/review-groups';
import { executeProperties, propertyEditor } from '../../test/support/property-execution';
import {
  executeAtomic,
  proposedListening,
  proposedVisual,
} from '../../test/support/atomic-execution';
import { buildBlockIndex } from '../blocks/block-index';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import type { NativeReviewGroup } from './groups';
import { LINK_CONTEXT, linkDocument, linkPlan } from '../../test/support/link-review';
import { trackHyperlinkPlan } from './track-hyperlinks';

const ids = (editor: DocxEditorInstance) => groups(editor).map((group) => group.id);
function insert(editor: DocxEditorInstance, paragraphId: string, offset: number, text: string) {
  const surface = editor.surface!;
  surface.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
  surface.insertPlainText(text);
  surface.flushPendingInput();
}
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
  if (result.rejected) throw new Error(String(result.reason));
}

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const action of ['acceptRevision', 'rejectRevision'] as const)
      test(`nested and application group identities survive concurrency, ${action}, history, geometry, reconnect, and saved reopen; clients ${clients.join('/')}, receiver ${receiver}`, async () => {
        const pair = await createTwoEditorFixture({
          document: writeOoxmlPackage(groupFixture()),
          clientIds: clients,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
            catId = groupParagraph(alice, 'The cat sat.'),
            cellId = groupParagraph(alice, 'A dog ran.'),
            tailId = groupParagraph(alice, 'Keep tail');
          pair.pause();
          proposeGroupText(alice, 'The cat sat.', 'The car sat.');
          const property = executeProperties(alice);
          if (!property.success) throw new Error(property.error);
          proposeGroupText(bob, 'A dog ran.', 'A hen ran.');
          const atomic = executeAtomic(bob, [proposedVisual, proposedListening]);
          if (!atomic.success) throw new Error(atomic.error);
          insert(bob, tailId, 0, 'Bob: ');
          bob.surface!.setSelection({
            anchor: { paragraphId: cellId, offset: 2 },
            head: { paragraphId: cellId, offset: 2 },
          });
          pair.reconnect(receiver);
          await pair.settle();
          const pendingIds = ids(alice);
          expect(pendingIds).toHaveLength(5);
          for (const editor of pair.editors) {
            expect(ids(editor)).toEqual(pendingIds);
            expect(
              groups(editor).map((group) => group.properties.map((member) => member.blockId))
            ).toEqual([[], ['activity'], [], ['visual'], ['audio']]);
            expect(groups(editor).every((group) => !group.readOnly)).toBe(true);
            expect(paragraphTextOf(editor.surface!.session.part(), tailId)).toBe('Bob: Keep tail');
          }
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          );
          expect(remote?.head.nodeId).toBe(cellId);
          expect(remote!.head.offset).toBe(bob.surface!.state().selection!.head.offset);
          expect(
            caretAt(alice.surface!.layout(), { paragraphId: cellId, offset: remote!.head.offset })
          ).not.toBeNull();
          const pending = await pair.savedDigests();
          expect(pending[0]).toEqual(pending[1]);
          // Alice's last local unit is her property proposal. Bob's native groups stay intact.
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          const withoutActivity = pendingIds.filter((id) => id !== pendingIds[1]);
          for (const editor of pair.editors) expect(ids(editor)).toEqual(withoutActivity);
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(ids(editor)).toEqual(pendingIds);
          const target = groups(bob).find(
            (group) => group.anchor?.paragraphId === catId && group.properties.length === 0
          )!;
          decide(bob, target, action);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(ids(editor)).toEqual(pendingIds.filter((id) => id !== target.id));
            expect(paragraphTextOf(editor.surface!.session.part(), catId)).toBe(
              action === 'acceptRevision' ? 'The car sat.' : 'The cat sat.'
            );
            expect(paragraphTextOf(editor.surface!.session.part(), tailId)).toBe('Bob: Keep tail');
          }
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(ids(editor)).toEqual(pendingIds);
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          // A new earlier group changes every later path, while each later identity stays fixed.
          const beforeLeading = ids(alice);
          insertLeadingGroup(alice);
          await pair.settle();
          for (const editor of pair.editors) expect(ids(editor).slice(1)).toEqual(beforeLeading);
          const finalIds = ids(alice),
            final = await pair.savedDigests();
          expect(final[0]).toEqual(final[1]);
          for (const editor of pair.editors) {
            const read = readOoxmlPackage(new Uint8Array(await editor.save()));
            if (!read.ok) throw new Error(read.reason);
            const reopened = propertyEditor(read.package);
            try {
              expect(ids(reopened)).toEqual(finalIds);
            } finally {
              reopened.destroy();
            }
          }
        } finally {
          pair.destroy();
        }
      }, 30_000);

for (const receiver of [0, 1] as const)
  test(`deleting a nested owner during a peer replacement preserves its group identity on deletion undo; receiver ${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(groupFixture()),
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
        owner = buildBlockIndex(alice.surface!.session.currentPackage()).byId.get('activity')!,
        tailId = groupParagraph(alice, 'Keep tail');
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
      proposeGroupText(bob, 'A dog ran.', 'A hen ran.');
      const pendingId = ids(bob)[0]!;
      insert(bob, tailId, 0, 'Bob: ');
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(groups(editor)).toEqual([]);
        expect(paragraphTextOf(editor.surface!.session.part(), tailId)).toBe('Bob: Keep tail');
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(ids(editor)).toEqual([pendingId]);
        expect(groups(editor)[0]!.native).toHaveLength(2);
        expect(paragraphTextOf(editor.surface!.session.part(), tailId)).toBe('Bob: Keep tail');
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
  for (const action of ['acceptRevision', 'rejectRevision'] as const)
    test(`a retargeted link retains one persistent group through peer text and ${action}; clients ${clients.join('/')}`, async () => {
      const pair = await createTwoEditorFixture({
        document: linkDocument(),
        clientIds: clients,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          tailId = groupParagraph(alice, 'Tail');
        pair.pause();
        const proposed = trackHyperlinkPlan(
          alice,
          linkPlan(alice, { url: 'https://example.org/new' }),
          LINK_CONTEXT
        );
        if (!proposed.success) throw new Error(proposed.error);
        const pending = ids(alice);
        expect(pending).toHaveLength(1);
        insert(bob, tailId, 0, 'Bob: ');
        pair.reconnect(action === 'acceptRevision' ? 0 : 1);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(ids(editor)).toEqual(pending);
          const reopened = propertyEditor(editor.surface!.session.currentPackage());
          try {
            expect(ids(reopened)).toEqual(pending);
          } finally {
            reopened.destroy();
          }
        }
        decide(bob, groups(bob)[0]!, action);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(groups(editor)).toEqual([]);
          expect(paragraphTextOf(editor.surface!.session.part(), tailId)).toBe('Bob: Tail');
        }
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) expect(ids(editor)).toEqual(pending);
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
      } finally {
        pair.destroy();
      }
    }, 30_000);

for (const receiver of [0, 1] as const)
  test(`a deleted property owner retains a visible read-only group and recovers the same ID after undo; receiver ${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(groupFixture()),
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
        owner = buildBlockIndex(alice.surface!.session.currentPackage()).byId.get('activity')!;
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
      const proposed = executeProperties(bob);
      if (!proposed.success) throw new Error(proposed.error);
      const id = ids(bob)[0]!;
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(ids(editor)).toEqual([id]);
        expect(groups(editor)[0]!.readOnly).toBe(true);
        expect(groups(editor)[0]!.anchor).toBeNull();
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(ids(editor)).toEqual([id]);
        expect(groups(editor)[0]!.readOnly).toBe(false);
      }
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  }, 30_000);

// Existing room snapshots must pass through unchanged. A migration must prepare their store.
test('an unprepared joined room refuses tracked native groups atomically for both actors', async () => {
  const document = writeOoxmlPackage(groupFixture()),
    ydoc = new Y.Doc(),
    awareness = new Awareness(ydoc);
  let snapshot: Uint8Array;
  try {
    const room = await createDocumentCollaboration({
      ydoc,
      awareness,
      documentId: 'sciobot-parity',
      identity: { actorId: 'bootstrap', name: 'Bootstrap' },
      bootstrap: { kind: 'create', document },
    });
    try {
      snapshot = Y.encodeStateAsUpdate(ydoc);
    } finally {
      room.destroy();
    }
  } finally {
    awareness.destroy();
    ydoc.destroy();
  }
  const pair = await createTwoEditorFixture({
    document,
    roomSnapshot: snapshot,
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
  });
  try {
    for (const editor of pair.editors) {
      const session = editor.surface!.session,
        before = writeOoxmlPackage(session.currentPackage()),
        revision = session.packageRevision();
      expect(readReviewGroupStore(session.currentPackage()).status).toBe('absent');
      expect(() => proposeGroupText(editor, 'The cat sat.', 'The car sat.')).toThrow(
        'review-group-store-initialization-required'
      );
      expect(writeOoxmlPackage(session.currentPackage())).toEqual(before);
      expect(session.packageRevision()).toBe(revision);
      expect(session.canUndo()).toBe(false);
      expect(groups(editor)).toEqual([]);
    }
    await pair.settle();
    const saved = await pair.savedDigests();
    expect(saved[0]).toEqual(saved[1]);
  } finally {
    pair.destroy();
  }
}, 60_000);
