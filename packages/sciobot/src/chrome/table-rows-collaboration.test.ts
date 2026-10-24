// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { caretAt, selectionRects } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { manualTableDocument, recoverFirstTableCellSource } from '../../test/support/manual-table';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { changeManualTableRow } from './table-rows';
import { nativeTableAnchor } from './tables';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const action of ['above', 'below', 'delete'] as const)
    test(`row ${action} retains offline edits in the affected row, independent history and exact Word restoration ${clients}`, async () => {
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        document: manualTableDocument(),
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!;
        const id = a.session.paragraphIds()[2]!,
          point = { paragraphId: id, offset: 3 };
        a.setSelection({ anchor: point, head: point });
        const before = a.session.part(),
          text = paragraphTextOf(before, id)!;
        const peer = applyTreeOp(before, {
          op: 'insertText',
          paragraphId: id,
          offset: text.length,
          text: ' peer',
        });
        if (!peer.ok) throw new Error(peer.reason);
        const expectedBefore = canonicalOoxmlFingerprint(peer.part),
          anchor = nativeTableAnchor(alice!)!;
        b.setSelection({
          anchor: { paragraphId: id, offset: text.length },
          head: { paragraphId: id, offset: text.length },
        });
        pair.pause();
        expect(changeManualTableRow(alice!, action).ok).toBe(true);
        if (action !== 'delete') expect(nativeTableAnchor(alice!)?.tableId).toBe(anchor.tableId);
        b.insertPlainText(' peer');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(expectedBefore);
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(
          canonicalOoxmlFingerprint(before)
        );
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(expectedBefore);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        const selection = {
          anchor: { paragraphId: id, offset: 0 },
          head: { paragraphId: id, offset: text.length + 5 },
        };
        b.setSelection(selection);
        await pair.settle();
        expect(
          pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob')?.head.nodeId
        ).toBe(id);
        for (const editor of pair.editors) {
          const s = editor.surface!;
          expect(caretAt(s.layout(), { paragraphId: id, offset: 2 })).not.toBeNull();
          expect(
            selectionRects(s.layout(), selection, s.session.paragraphIds()).length
          ).toBeGreaterThan(0);
          const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!saved.ok) throw new Error(saved.reason);
          expect(
            canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
          ).toEqual(expectedBefore);
        }
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
      } finally {
        pair.destroy();
      }
    });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`recovered peer source in cell one supports row insertion, native typing, exact undo and preserved header behavior ${clients}`, async () => {
    const pair = await createTwoEditorFixture({
      clientIds: clients,
      document: manualTableDocument(),
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors,
        a = alice!.surface!,
        b = bob!.surface!;
      const recovered = await recoverFirstTableCellSource(pair),
        point = { paragraphId: recovered, offset: 2 };
      a.setSelection({ anchor: point, head: point });
      expect(nativeTableAnchor(alice!)).not.toBeNull();
      const before = a.session.part();
      const changedBefore = applyTreeOp(before, {
        op: 'insertText',
        paragraphId: recovered,
        offset: 2,
        text: '!',
      });
      if (!changedBefore.ok) throw new Error(changedBefore.reason);
      expect(changeManualTableRow(alice!, 'below').ok).toBe(true);
      const insertedParagraph = a.state().selection.head.paragraphId;
      expect(insertedParagraph).not.toBe(recovered);
      expect(nativeTableAnchor(alice!)).not.toBeNull();
      expect(paragraphTextOf(a.session.part(), insertedParagraph)).toBe('');
      const after = a.session.part();
      const changedAfter = applyTreeOp(after, {
        op: 'insertText',
        paragraphId: recovered,
        offset: 2,
        text: '!',
      });
      if (!changedAfter.ok) throw new Error(changedAfter.reason);
      b.setSelection({ anchor: point, head: point });
      b.insertPlainText('!');
      b.flushPendingInput();
      await pair.settle();
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(
          canonicalOoxmlFingerprint(changedBefore.part)
        );
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(
          canonicalOoxmlFingerprint(changedAfter.part)
        );
      a.setSelection({
        anchor: { paragraphId: insertedParagraph, offset: 0 },
        head: { paragraphId: insertedParagraph, offset: 0 },
      });
      a.insertPlainText('Live');
      a.flushPendingInput();
      await pair.settle();
      for (const editor of pair.editors) {
        expect(paragraphTextOf(editor.surface!.session.part(), insertedParagraph)).toBe('Live');
        expect(paragraphTextOf(editor.surface!.session.part(), recovered)).toBe('Pe!er');
      }
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(
        canonicalOoxmlFingerprint(changedAfter.part)
      );
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
