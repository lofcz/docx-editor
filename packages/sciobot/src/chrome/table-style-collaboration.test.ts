// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
} from '@docx-editor.dev/core/store';
import { caretAt, selectionRects } from '@docx-editor.dev/core/layout';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { manualColumnDocument, columnShape } from '../../test/support/manual-table-columns';
import { selectTableCells, savedTableFingerprint } from '../../test/support/manual-table-controls';
import { mergeManualTableCells, unmergeManualTableCell } from './table-merges';
import { changeManualTableAppearance, type TableAppearanceChange } from './table-appearance';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const direction of ['horizontal', 'vertical', 'rectangle'] as const)
    test(`${direction} controls preserve peer text, moved identities, original-grid split, history, deletion, reconnect and geometry ${clients}`, async () => {
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        document: manualColumnDocument('none'),
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!,
          source = columnShape(a.session.part());
        const rowEnd = direction === 'horizontal' ? 1 : 2,
          colEnd = direction === 'vertical' ? 0 : 1;
        const moved = source.rows[rowEnd]!.cells[colEnd]!.paragraphs[0]!,
          owner = source.rows[1]!.cells[0]!.paragraphs[0]!;
        selectTableCells(alice!, [1, rowEnd], [0, colEnd]);
        const before = a.session.part();
        const peer = applyTreeOp(before, {
          op: 'insertText',
          paragraphId: moved.id,
          offset: 2,
          text: '!',
        });
        if (!peer.ok) throw new Error(peer.reason);
        pair.pause();
        expect(mergeManualTableCells(alice!).ok).toBe(true);
        const point = { paragraphId: moved.id, offset: 2 };
        b.setSelection({ anchor: point, head: point });
        b.insertPlainText('!');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        for (const e of pair.editors)
          expect(paragraphTextOf(e.surface!.session.part(), moved.id)).toBe(
            moved.text!.slice(0, 2) + '!' + moved.text!.slice(2)
          );
        const merged = canonicalOoxmlFingerprint(a.session.part());
        expect(canonicalOoxmlFingerprint(b.session.part())).toEqual(merged);
        const selection = {
          anchor: { paragraphId: moved.id, offset: 0 },
          head: { paragraphId: moved.id, offset: 3 },
        };
        b.setSelection(selection);
        await pair.settle();
        expect(
          pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob')?.head.nodeId
        ).toBe(moved.id);
        for (const e of pair.editors) {
          const s = e.surface!;
          expect(caretAt(s.layout(), selection.head)).not.toBeNull();
          expect(
            selectionRects(s.layout(), selection, s.session.paragraphIds()).length
          ).toBeGreaterThan(0);
          expect(await savedTableFingerprint(e)).toEqual(merged);
        }
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const e of pair.editors)
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(
            canonicalOoxmlFingerprint(peer.part)
          );
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(
          canonicalOoxmlFingerprint(before)
        );
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(merged);
        const at = { paragraphId: owner.id, offset: 0 };
        a.setCellSelection(null);
        a.setSelection({ anchor: at, head: at });
        expect(unmergeManualTableCell(alice!).ok).toBe(true);
        await pair.settle();
        const splitPart = a.session.part(),
          split = canonicalOoxmlFingerprint(splitPart);
        for (const e of pair.editors) {
          expect(columnShape(e.surface!.session.part()).widths).toEqual([1200, 1800, 2400]);
          expect(await savedTableFingerprint(e)).toEqual(split);
          expect(caretAt(e.surface!.layout(), { paragraphId: moved.id, offset: 2 })).not.toBeNull();
        }
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(merged);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(split);
        // Delete the containing table while the peer edits moved content, then recover that text.
        a.setSelection({ anchor: at, head: at });
        pair.pause();
        expect(alice!.exec({ type: 'deleteTable' }).ok).toBe(true);
        const peerAt = { paragraphId: moved.id, offset: 3 };
        b.setSelection({ anchor: peerAt, head: peerAt });
        b.insertPlainText('?');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 1 : 0);
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        const expected = applyTreeOp(splitPart, {
          op: 'insertText',
          paragraphId: moved.id,
          offset: 3,
          text: '?',
        });
        if (!expected.ok) throw new Error(expected.reason);
        for (const e of pair.editors) {
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(
            canonicalOoxmlFingerprint(expected.part)
          );
          expect(await savedTableFingerprint(e)).toEqual(canonicalOoxmlFingerprint(expected.part));
        }
      } finally {
        pair.destroy();
      }
    });

const changes: TableAppearanceChange[] = [
  { kind: 'fill', color: '1A2B3C' },
  { kind: 'column-width', twips: 2100 },
  { kind: 'row-height', twips: 720 },
  { kind: 'borders', color: '123456', size: 8 },
  { kind: 'headers', count: 2 },
];
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const merge of ['none', 'horizontal', 'vertical'] as const)
    test(`${merge} appearance controls keep peer edits, individual undo, geometry, deletion and exact Word output ${clients}`, async () => {
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        document: manualColumnDocument(merge),
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!,
          shape = columnShape(a.session.part()),
          id = shape.rows[1]!.cells[1]!.paragraphs[0]!.id;
        const point = { paragraphId: id, offset: 2 };
        a.setSelection({ anchor: point, head: point });
        b.setSelection({ anchor: point, head: point });
        const source = a.session.part(),
          snapshots = [source];
        pair.pause();
        for (const change of changes) {
          expect(changeManualTableAppearance(alice!, change).ok).toBe(true);
          snapshots.push(a.session.part());
        }
        b.insertPlainText('!');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        const peerParts = snapshots.map((part) => {
          const changed = applyTreeOp(part, {
            op: 'insertText',
            paragraphId: id,
            offset: 2,
            text: '!',
          });
          if (!changed.ok) throw new Error(changed.reason);
          return canonicalOoxmlFingerprint(changed.part);
        });
        for (const e of pair.editors)
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(peerParts.at(-1)!);
        const selection = {
          anchor: { paragraphId: id, offset: 0 },
          head: { paragraphId: id, offset: 3 },
        };
        b.setSelection(selection);
        await pair.settle();
        expect(
          pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob')?.head.nodeId
        ).toBe(id);
        for (const e of pair.editors) {
          const s = e.surface!;
          expect(caretAt(s.layout(), selection.head)).not.toBeNull();
          expect(
            selectionRects(s.layout(), selection, s.session.paragraphIds()).length
          ).toBeGreaterThan(0);
          expect(await savedTableFingerprint(e)).toEqual(peerParts.at(-1)!);
        }
        for (let i = changes.length - 1; i >= 0; i--) {
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const e of pair.editors)
            expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(peerParts[i]!);
        }
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(
          canonicalOoxmlFingerprint(source)
        );
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (let i = 1; i < peerParts.length; i++) {
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const e of pair.editors)
            expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(peerParts[i]!);
        }
        a.setSelection({ anchor: point, head: point });
        pair.pause();
        expect(alice!.exec({ type: 'deleteRow' }).ok).toBe(true);
        b.setSelection({
          anchor: { paragraphId: id, offset: 3 },
          head: { paragraphId: id, offset: 3 },
        });
        b.insertPlainText('?');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 1 : 0);
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        const extra = applyTreeOp(snapshots.at(-1)!, {
          op: 'insertText',
          paragraphId: id,
          offset: 2,
          text: '!?',
        });
        if (!extra.ok) throw new Error(extra.reason);
        for (const e of pair.editors) {
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(
            canonicalOoxmlFingerprint(extra.part)
          );
          expect(await savedTableFingerprint(e)).toEqual(canonicalOoxmlFingerprint(extra.part));
        }
      } finally {
        pair.destroy();
      }
    });
