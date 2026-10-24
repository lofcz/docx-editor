// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { caretAt, paragraphsInCells, selectionRects } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import {
  insertManualTable,
  nativeTableAnchor,
  nativeTableCellIds,
  navigateTableCell,
} from './tables';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`2x3 table shares insertion, peer history, editing during deletion, geometry and save/reopen ${clients} receiver ${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        document: docxFixture(
          '<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
        ),
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!;
        const source = a.session.paragraphIds()[0]!,
          before = canonicalOoxmlFingerprint(a.session.part());
        const at = { paragraphId: source, offset: 5 };
        a.setSelection({ anchor: at, head: at });
        b.setSelection({
          anchor: { paragraphId: source, offset: 10 },
          head: { paragraphId: source, offset: 10 },
        });
        pair.pause();
        expect(insertManualTable(alice!, 2, 3).ok).toBe(true);
        const anchor = nativeTableAnchor(alice!)!;
        b.insertPlainText(' peer');
        b.flushPendingInput();
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(nativeTableCellIds(editor, anchor.tableId)).toHaveLength(6);
          expect(paragraphTextOf(editor.surface!.session.part(), source)).toBe('Alpha beta peer');
        }
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(nativeTableCellIds(editor, anchor.tableId)).toHaveLength(0);
        expect(paragraphTextOf(a.session.part(), source)).toBe('Alpha beta peer');
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        const cell = nativeTableCellIds(alice!, anchor.tableId)[0]!,
          id = paragraphsInCells(a.layout(), [cell])[0]!;
        b.setSelection({
          anchor: { paragraphId: id, offset: 0 },
          head: { paragraphId: id, offset: 0 },
        });
        b.insertPlainText('Cell');
        b.flushPendingInput();
        await pair.settle();
        for (const s of [a, b]) {
          expect(caretAt(s.layout(), { paragraphId: id, offset: 2 })).not.toBeNull();
          expect(
            selectionRects(
              s.layout(),
              { anchor: { paragraphId: id, offset: 0 }, head: { paragraphId: id, offset: 4 } },
              s.session.paragraphIds()
            ).length
          ).toBeGreaterThan(0);
        }
        expect(
          pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob')?.head.nodeId
        ).toBe(id);
        a.setSelection({
          anchor: { paragraphId: id, offset: 0 },
          head: { paragraphId: id, offset: 0 },
        });
        pair.pause();
        expect(alice!.exec({ type: 'deleteTable' }).ok).toBe(true);
        b.setSelection({
          anchor: { paragraphId: id, offset: 4 },
          head: { paragraphId: id, offset: 4 },
        });
        b.insertPlainText('!');
        b.flushPendingInput();
        pair.reconnect(receiver === 0 ? 1 : 0);
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(nativeTableCellIds(editor, anchor.tableId)).toHaveLength(6);
          expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Cell!');
          expect(paragraphTextOf(editor.surface!.session.part(), source)).toBe('Alpha beta peer');
          const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!saved.ok) throw new Error(saved.reason);
          expect(
            canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
          ).toEqual(canonicalOoxmlFingerprint(editor.surface!.session.part()));
        }
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(paragraphTextOf(a.session.part(), id)).toBe('Cell');
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(paragraphTextOf(a.session.part(), id)).toBe('Cell!');
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
        expect(canonicalOoxmlFingerprint(a.session.part())).not.toBe(before);
      } finally {
        pair.destroy();
      }
    });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`undoing a retained row birth preserves earlier cell typing and redo restores only the row ${clients}`, async () => {
    const pair = await createTwoEditorFixture({
      clientIds: clients,
      document: docxFixture('<w:p><w:r><w:t>Keep</w:t></w:r></w:p>'),
    });
    try {
      const editor = pair.editors[0]!,
        s = editor.surface!;
      expect(insertManualTable(editor, 2, 3).ok).toBe(true);
      const type = (text: string) => {
        const event = new InputEvent('beforeinput', {
          inputType: 'insertText',
          data: text,
          bubbles: true,
          cancelable: true,
        });
        s.hostElement().querySelector('[contenteditable="true"]')!.dispatchEvent(event);
        expect(event.defaultPrevented).toBe(true);
        s.flushPendingInput();
      };
      type('First');
      expect(navigateTableCell(editor, 'next').ok).toBe(true);
      type('Second');
      expect(navigateTableCell(editor, 'previous').ok).toBe(true);
      const anchor = nativeTableAnchor(editor)!,
        cells = nativeTableCellIds(editor, anchor.tableId);
      const last = paragraphsInCells(s.layout(), [cells[5]!])[0]!,
        point = { paragraphId: last, offset: 0 };
      s.setSelection({ anchor: point, head: point });
      const before = canonicalOoxmlFingerprint(s.session.part());
      expect(navigateTableCell(editor, 'next').ok).toBe(true);
      const after = canonicalOoxmlFingerprint(s.session.part());
      expect(nativeTableCellIds(editor, anchor.tableId)).toHaveLength(9);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const peer of pair.editors)
        expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const peer of pair.editors)
        expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(after);
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
