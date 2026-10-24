// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
} from '@docx-editor.dev/core/store';
import { caretAt, selectionRects } from '@docx-editor.dev/core/layout';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { tableIndentDocument, indentTable } from '../../test/support/table-indent';
import { savedTableFingerprint } from '../../test/support/manual-table-controls';
import { setManualTableIndent } from './table-indent';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const merged of [false, true])
    test(`table indent/reset preserves peer edits, independent history, reconnect, presence and table deletion ${clients}/${merged}`, async () => {
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        document: tableIndentDocument(0, merged),
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!,
          id = a.session.paragraphIds().at(-2)!,
          point = { paragraphId: id, offset: 2 };
        a.setSelection({ anchor: point, head: point });
        b.setSelection({ anchor: point, head: point });
        const before = a.session.part(),
          authored = [before];
        pair.pause();
        for (const pixels of [24, 0, 24]) {
          expect(setManualTableIndent(alice!, pixels).ok).toBe(true);
          authored.push(a.session.part());
        }
        b.insertPlainText('!');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        const peer = (part: typeof before, text = '!') => {
          const result = applyTreeOp(part, { op: 'insertText', paragraphId: id, offset: 2, text });
          if (!result.ok) throw new Error(result.reason);
          return result.part;
        };
        const expected = authored.map((p) => canonicalOoxmlFingerprint(peer(p)));
        for (const e of pair.editors)
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(expected.at(-1)!);
        const selection = {
          anchor: { paragraphId: id, offset: 0 },
          head: { paragraphId: id, offset: 4 },
        };
        b.setSelection(selection);
        await pair.settle();
        expect(
          pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob')?.head.nodeId
        ).toBe(id);
        for (const e of pair.editors) {
          expect(caretAt(e.surface!.layout(), selection.head)).not.toBeNull();
          expect(
            selectionRects(e.surface!.layout(), selection, e.surface!.session.paragraphIds()).length
          ).toBeGreaterThan(0);
          expect(await savedTableFingerprint(e)).toEqual(expected.at(-1)!);
        }
        for (let i = expected.length - 2; i >= 0; i--) {
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const e of pair.editors)
            expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(expected[i]!);
        }
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(
          canonicalOoxmlFingerprint(before)
        );
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (let i = 1; i < expected.length; i++) {
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(expected[i]!);
        }
        a.setSelection({ anchor: point, head: point });
        const savedBeforeDelete = a.session.part();
        pair.pause();
        expect(alice!.exec({ type: 'deleteTable' }).ok).toBe(true);
        const edit = { paragraphId: id, offset: 3 };
        b.setSelection({ anchor: edit, head: edit });
        b.insertPlainText('?');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 1 : 0);
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        const recovered = applyTreeOp(savedBeforeDelete, {
          op: 'insertText',
          paragraphId: id,
          offset: 3,
          text: '?',
        });
        if (!recovered.ok) throw new Error(recovered.reason);
        for (const e of pair.editors) {
          expect(indentTable(e.surface!.session.part()).id).toBe(indentTable(before).id);
          expect(paragraphTextOf(e.surface!.session.part(), id)).toContain('!?');
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(
            canonicalOoxmlFingerprint(recovered.part)
          );
          expect(await savedTableFingerprint(e)).toEqual(canonicalOoxmlFingerprint(recovered.part));
        }
      } finally {
        pair.destroy();
      }
    });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const decision of ['accept', 'reject'] as const)
    test(`tracked table indentation ${decision} retains offline peer text and exact decision history ${clients}`, async () => {
      const { SCIOBOT_NATIVE_REVIEW_MODULE } = await import('../review/native-module');
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        document: tableIndentDocument(0),
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!,
          before = a.session.part(),
          id = a.session.paragraphIds().at(-2)!,
          point = { paragraphId: id, offset: 2 };
        a.setSelection({ anchor: point, head: point });
        b.setSelection({ anchor: point, head: point });
        expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
        pair.pause();
        expect(setManualTableIndent(alice!, 24).ok).toBe(true);
        const authored = a.session.part();
        b.insertPlainText('!');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        const pendingPeer = applyTreeOp(authored, {
          op: 'insertText',
          paragraphId: id,
          offset: 2,
          text: '!',
        });
        if (!pendingPeer.ok) throw new Error(pendingPeer.reason);
        const pending = canonicalOoxmlFingerprint(pendingPeer.part);
        for (const e of pair.editors) {
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(pending);
          expect(await savedTableFingerprint(e)).toEqual(pending);
          expect(caretAt(e.surface!.layout(), point)).not.toBeNull();
        }
        expect(
          alice!.exec({ type: 'resolveAllReviewChanges', scope: 'document', action: decision }).ok
        ).toBe(true);
        await pair.settle();
        const direct =
          decision === 'reject'
            ? { ok: true as const, part: before }
            : applyTreeOp(before, {
                op: 'setTableProperties',
                tableId: indentTable(before).id,
                indentTwips: 360,
              });
        if (!direct.ok) throw new Error(direct.reason);
        const peer = applyTreeOp(direct.part, {
          op: 'insertText',
          paragraphId: id,
          offset: 2,
          text: '!',
        });
        if (!peer.ok) throw new Error(peer.reason);
        const expected = canonicalOoxmlFingerprint(peer.part);
        for (const e of pair.editors) {
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(expected);
          expect(await savedTableFingerprint(e)).toEqual(expected);
        }
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const e of pair.editors)
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(pending);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const e of pair.editors)
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(expected);
      } finally {
        pair.destroy();
      }
    });
