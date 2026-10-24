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
import { manualColumnDocument, columnShape } from '../../test/support/manual-table-columns';
import { changeManualTableColumn } from './table-columns';
import { nativeTableAnchor } from './tables';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const merge of ['horizontal', 'vertical'] as const)
    for (const action of ['left', 'right', 'delete'] as const)
      test(`column ${action} with ${merge} merge preserves concurrent cell text, independent history, geometry and Word convergence ${clients}`, async () => {
        const pair = await createTwoEditorFixture({
          clientIds: clients,
          document: manualColumnDocument(merge),
        });
        try {
          const [alice, bob] = pair.editors,
            a = alice!.surface!,
            b = bob!.surface!;
          const shape = columnShape(a.session.part()),
            id = shape.rows[1]!.cells[1]!.paragraphs[0]!.id;
          const text = paragraphTextOf(a.session.part(), id)!,
            point = { paragraphId: id, offset: text.length };
          a.setSelection({ anchor: point, head: point });
          b.setSelection({ anchor: point, head: point });
          const before = a.session.part();
          const peer = applyTreeOp(before, {
            op: 'insertText',
            paragraphId: id,
            offset: text.length,
            text: ' peer',
          });
          if (!peer.ok) throw new Error(peer.reason);
          const expectedBefore = canonicalOoxmlFingerprint(peer.part);
          pair.pause();
          expect(changeManualTableColumn(alice!, action).ok).toBe(true);
          const insertedParagraph = a.state().selection.head.paragraphId;
          if (action !== 'delete') {
            expect(insertedParagraph).not.toBe(id);
            expect(paragraphTextOf(a.session.part(), insertedParagraph)).toBe('');
            expect(nativeTableAnchor(alice!)).not.toBeNull();
          }
          const after = a.session.part();
          b.insertPlainText(' peer');
          b.flushPendingInput();
          pair.reconnect(clients[0] === 41 ? 0 : 1);
          await pair.settle();
          const converged = canonicalOoxmlFingerprint(a.session.part());
          expect(canonicalOoxmlFingerprint(b.session.part())).toEqual(converged);
          // Check geometry on the changed grid before history restores the source grid.
          const placed = columnShape(a.session.part()).rows[0]!.cells[0]!.paragraphs[0]!;
          const visibleSelection = {
            anchor: { paragraphId: placed.id, offset: 0 },
            head: { paragraphId: placed.id, offset: 2 },
          };
          b.setSelection(visibleSelection);
          await pair.settle();
          expect(
            pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob')?.head.nodeId
          ).toBe(placed.id);
          for (const editor of pair.editors) {
            const surface = editor.surface!;
            expect(caretAt(surface.layout(), visibleSelection.head)).not.toBeNull();
            expect(
              selectionRects(surface.layout(), visibleSelection, surface.session.paragraphIds())
                .length
            ).toBeGreaterThan(0);
          }

          if (action !== 'delete') {
            const expected = applyTreeOp(after, {
              op: 'insertText',
              paragraphId: id,
              offset: text.length,
              text: ' peer',
            });
            if (!expected.ok) throw new Error(expected.reason);
            expect(converged).toEqual(canonicalOoxmlFingerprint(expected.part));
            a.insertPlainText('Live');
            a.flushPendingInput();
            await pair.settle();
            for (const editor of pair.editors)
              expect(paragraphTextOf(editor.surface!.session.part(), insertedParagraph)).toBe(
                'Live'
              );
            const freshSelection = {
              anchor: { paragraphId: insertedParagraph, offset: 0 },
              head: { paragraphId: insertedParagraph, offset: 4 },
            };
            a.setSelection(freshSelection);
            await pair.settle();
            expect(
              pair.rooms[1]!.session.remoteSelections().find((s) => s.actorId === 'alice')?.head
                .nodeId
            ).toBe(insertedParagraph);
            expect(
              selectionRects(b.layout(), freshSelection, b.session.paragraphIds()).length
            ).toBeGreaterThan(0);
            expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(converged);
          }
          for (const editor of pair.editors) {
            const read = readOoxmlPackage(new Uint8Array(await editor.save()));
            if (!read.ok) throw new Error(read.reason);
            expect(
              canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!)
            ).toEqual(converged);
          }
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(
              expectedBefore
            );
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
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(converged);
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
            const read = readOoxmlPackage(new Uint8Array(await editor.save()));
            if (!read.ok) throw new Error(read.reason);
            expect(
              canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!)
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
  for (const action of ['left', 'right'] as const)
    test(`shared ${action} column insertion synthesizes a missing grid and adopts its live cell ${clients}`, async () => {
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        document: manualColumnDocument('none', false),
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!;
        const source = columnShape(a.session.part()),
          point = { paragraphId: source.rows[1]!.cells[1]!.paragraphs[0]!.id, offset: 2 };
        a.setSelection({ anchor: point, head: point });
        const before = canonicalOoxmlFingerprint(a.session.part());
        expect(source.widths).toEqual([]);
        expect(changeManualTableColumn(alice!, action).ok).toBe(true);
        const live = a.state().selection.head.paragraphId;
        expect(live).not.toBe(point.paragraphId);
        expect(paragraphTextOf(a.session.part(), live)).toBe('');
        expect(nativeTableAnchor(alice!)).not.toBeNull();
        const shape = columnShape(a.session.part());
        expect(shape.widths).toEqual([1200, 1800, 1800, 2400]);
        a.insertPlainText('Live');
        a.flushPendingInput();
        await pair.settle();
        expect(paragraphTextOf(bob!.surface!.session.part(), live)).toBe('Live');
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(before);
        expect(canonicalOoxmlFingerprint(bob!.surface!.session.part())).toEqual(before);
      } finally {
        pair.destroy();
      }
    });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const action of ['left', 'right'] as const)
    test(`gridless ${action} column undo retains peer text with its complete grid, then restores exact absence after peer undo ${clients}`, async () => {
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        document: manualColumnDocument('none', false),
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!;
        const original = columnShape(a.session.part()),
          point = { paragraphId: original.rows[1]!.cells[1]!.paragraphs[0]!.id, offset: 2 };
        const before = canonicalOoxmlFingerprint(a.session.part());
        a.setSelection({ anchor: point, head: point });
        expect(changeManualTableColumn(alice!, action).ok).toBe(true);
        const live = a.state().selection.head.paragraphId;
        await pair.settle();
        const at = { paragraphId: live, offset: 0 };
        b.setSelection({ anchor: at, head: at });
        b.insertPlainText('Peer');
        b.flushPendingInput();
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        const retained = canonicalOoxmlFingerprint(a.session.part());
        for (const editor of pair.editors) {
          const surface = editor.surface!,
            shape = columnShape(surface.session.part());
          expect(shape.widths).toEqual([1200, 1800, 1800, 2400]);
          expect(paragraphTextOf(surface.session.part(), live)).toBe('Peer');
          for (const row of original.rows)
            for (const cell of row.cells)
              for (const p of cell.paragraphs)
                expect(paragraphTextOf(surface.session.part(), p.id)).toBe(p.text);
          for (const row of shape.rows) expect(row.cells.reduce((n, c) => n + c.span, 0)).toBe(4);
          const read = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!read.ok) throw new Error(read.reason);
          expect(
            canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!)
          ).toEqual(retained);
          expect(caretAt(surface.layout(), { paragraphId: live, offset: 2 })).not.toBeNull();
        }
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(before);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(retained);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(retained);
      } finally {
        pair.destroy();
      }
    });
