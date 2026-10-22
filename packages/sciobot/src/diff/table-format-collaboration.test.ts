import { expect, test } from 'bun:test';
import { caretAt, selectionRects } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { nodeChildren, tableRows } from '../../test/support/native-table-oracle';
import {
  formatAttr as attr,
  formatBody as body,
  formatCells as cells,
  formatDescendants as descendants,
  formatNode as node,
  formatProperty as property,
  formatTables as tables,
} from '../../test/support/table-format-fixture';
import { nativeGridColumns } from './inserted-column-native';

for (const receiver of [0, 1] as const)
  test(`native table properties preserve metadata and peer edits through reconnect/history/deletion; receiver ${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: docxFixture(body(true)),
      documentId: `property-fidelity-${receiver}`,
      clientIds: [73, 41],
    });
    try {
      const [alice, bob] = pair.editors,
        first = alice!.surface!,
        second = bob!.surface!,
        initial = first.session.part(),
        tableId = tables(initial)[0]!.id,
        grid = nativeGridColumns(initial, tableId),
        originalCells = cells(initial, tableId),
        lastRow = tableRows(initial, tableId)[2]!,
        lastCell = originalCells[2]![1]!;
      const paragraph = descendants(originalCells[0]![0]!).find(
          (entry) => entry.kind === 'paragraph'
        )!,
        deletedParagraph = descendants(lastCell).find((entry) => entry.kind === 'paragraph')!,
        unrelated = tables(initial)[1]!;
      const originalCaret = caretAt(first.layout(), {
        paragraphId: deletedParagraph.id,
        offset: 2,
      });
      expect(originalCaret).not.toBeNull();
      const metadata = descendants(node(initial, tableId)).filter((entry) =>
        attr(entry, 'value', 'urn:keep')
      );
      const assertMetadata = (part: OoxmlPart, skipDeleted = false) => {
        for (const old of metadata) {
          if (skipDeleted && descendants(lastRow).some((entry) => entry.id === old.id)) continue;
          const next = node(part, old.id);
          expect(attr(next, 'value', 'urn:keep')).toBe(attr(old, 'value', 'urn:keep'));
          // All extension children remain unchanged. Cell/text children can contain peer text.
          if (
            old.kind !== 'textValue' &&
            ['tcW', 'tblW', 'tblLayout', 'trHeight', 'vAlign'].includes(old.localName)
          )
            expect(nodeChildren(next)).toEqual(nodeChildren(old));
        }
        expect(canonicalOoxmlFingerprint(node(part, unrelated.id))).toBe(
          canonicalOoxmlFingerprint(unrelated)
        );
      };
      const assertResize = () => {
        for (const editor of pair.editors) {
          const part = editor.surface!.session.part();
          expect(nativeGridColumns(part, tableId).map((entry) => attr(entry, 'w'))).toEqual([
            '1800',
            '1800',
            '4200',
          ]);
          expect(attr(property(node(part, tableId), 'tblPr', 'tblW'), 'w')).toBe('7800');
          expect(attr(property(node(part, lastRow.id), 'trPr', 'trHeight'), 'val')).toBe('600');
          expect(attr(property(node(part, lastRow.id), 'trPr', 'trHeight'), 'hRule')).toBe('exact');
          expect(
            attr(property(node(part, originalCells[0]![0]!.id), 'tcPr', 'gridSpan'), 'val')
          ).toBe('2');
          expect(paragraphTextOf(part, paragraph.id)).toBe('AB1 Bob');
          expect(
            attr(property(node(part, originalCells[0]![0]!.id), 'tcPr', 'vAlign'), 'val')
          ).toBe('center');
          assertMetadata(part);
        }
      };
      pair.pause();

      expect(
        first.applyAutomationOps(() => [
          {
            op: 'setTableColumnWidths',
            tableId,
            leftGridColumnId: grid[0]!.id,
            rightGridColumnId: grid[1]!.id,
            leftWidthTwips: 1800,
            rightWidthTwips: 1800,
          },
          {
            op: 'setTableRightEdgeWidth',
            tableId,
            gridColumnId: grid[2]!.id,
            columnWidthTwips: 4200,
            tableWidthTwips: 7800,
          },
          { op: 'setTableRowHeight', tableId, rowId: lastRow.id, heightTwips: 600 },
          {
            op: 'setTableCellVerticalAlignment',
            tableId,
            cellIds: [originalCells[0]![0]!.id, originalCells[1]![0]!.id],
            alignment: 'center',
          },
        ]).rejected
      ).toBe(false);

      expect(
        second.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: paragraph.id, offset: 3, text: ' Bob' },
        ]).rejected
      ).toBe(false);
      pair.reconnect(receiver);
      await pair.settle();
      assertResize();
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const part = editor.surface!.session.part();
        expect(nativeGridColumns(part, tableId).map((entry) => attr(entry, 'w'))).toEqual([
          '1200',
          '2400',
          '3600',
        ]);
        expect(attr(property(node(part, lastRow.id), 'trPr', 'trHeight'), 'hRule')).toBe('atLeast');
        expect(paragraphTextOf(part, paragraph.id)).toBe('AB1 Bob');
        assertMetadata(part);
      }
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      assertResize();
      second.setSelection({
        anchor: { paragraphId: deletedParagraph.id, offset: 0 },
        head: { paragraphId: deletedParagraph.id, offset: 2 },
      });
      await pair.settle();
      const remote = pair.rooms[0]!.session.remoteSelections().find(
        (entry) => entry.actorId === 'bob'
      )!;
      expect(remote.head.nodeId).toBe(deletedParagraph.id);
      expect(remote.anchor.nodeId).toBe(deletedParagraph.id);
      expect(remote.anchor.offset).toBe(0);
      const local = caretAt(second.layout(), { paragraphId: deletedParagraph.id, offset: 2 }),
        other = caretAt(first.layout(), {
          paragraphId: remote.head.nodeId,
          offset: remote.head.offset,
        });
      expect(local).not.toBeNull();
      expect(other).toEqual(local);
      expect(local!.x).toBeGreaterThan(originalCaret!.x);
      const remoteRange = {
        anchor: { paragraphId: remote.anchor.nodeId, offset: remote.anchor.offset },
        head: { paragraphId: remote.head.nodeId, offset: remote.head.offset },
      };
      const rects = selectionRects(first.layout(), remoteRange, first.session.paragraphIds());
      expect(rects.length).toBeGreaterThan(0);
      expect(rects).toEqual(
        selectionRects(second.layout(), remoteRange, second.session.paragraphIds())
      );
      const resizedDigests = await pair.savedDigests();
      expect(resizedDigests[0]).toEqual(resizedDigests[1]);
      pair.pause();

      expect(
        first.applyAutomationOps(() => [{ op: 'deleteTableRow', tableId, rowId: lastRow.id }])
          .rejected
      ).toBe(false);

      expect(
        second.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: deletedParagraph.id, offset: 2, text: ' peer' },
        ]).rejected
      ).toBe(false);
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        const part = editor.surface!.session.part();
        expect(findNode(part, lastRow.id)).toBeNull();
        expect(paragraphTextOf(part, paragraph.id)).toBe('AB1 Bob');
        assertMetadata(part, true);
      }
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      assertResize();
      for (const editor of pair.editors)
        expect(paragraphTextOf(editor.surface!.session.part(), deletedParagraph.id)).toBe(
          'B3 peer'
        );
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(findNode(editor.surface!.session.part(), lastRow.id)).toBeNull();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
