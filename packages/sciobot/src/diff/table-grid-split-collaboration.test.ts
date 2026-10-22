import { expect, test } from 'bun:test';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { namedChild, nodeChildren, rowCells } from '../../test/support/native-table-oracle';
import { buildBlockIndex } from '../blocks/block-index';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { nativeGridColumns } from './inserted-column-native';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
function fixture() {
  const loaded = readOoxmlPackage(
    docxFixture(
      p('Outside.') +
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="2400"/></w:tblGrid>' +
        '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr>' +
        p('Merged.') +
        '<w:sdt><w:sdtPr><w:tag w:val="scio:1:owned"/></w:sdtPr><w:sdtContent>' +
        p('Owned.') +
        '</w:sdtContent></w:sdt>' +
        '</w:tc></w:tr><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>' +
        p('Between.') +
        '<w:tbl><w:tblGrid><w:gridCol w:w="3600"/></w:tblGrid><w:tr><w:tc>' +
        p('Unrelated.') +
        '</w:tc></w:tr></w:tbl>' +
        p('Tail.')
    )
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const stored = withPayloadRecord(
    loaded.package,
    {
      id: 'owned',
      type: 'lesson_activity',
      version: 1,
      attributes: { name: 'Owned', duration: 5, keep: { asset: 'saved' } },
      unknownRoot: ['keep'],
    },
    'Owned'
  );
  if (!stored.ok) throw new Error(stored.reason);
  return writeOoxmlPackage(stored.pkg);
}
const tables = (part: OoxmlPart) =>
  nodeChildren(bodyStoryRoot(part)!).filter((node) => node.kind === 'table');
const rows = (part: OoxmlPart) =>
  nodeChildren(tables(part)[0]!).filter((node) => node.kind === 'tableRow');
const firstParagraph = (node: OoxmlNode) =>
  nodeChildren(node).find((child) => child.kind === 'paragraph')!;
const attr = (node: OoxmlNode, key: string) =>
  node.kind === 'textValue'
    ? undefined
    : node.attributes.find((attribute) => attribute.localName === key)?.value;
const widths = (part: OoxmlPart) =>
  nativeGridColumns(part, tables(part)[0]!.id).map((node) => attr(node, 'w'));

for (const receiver of [0, 1] as const)
  test(`existing-grid split retains payloads and peer text through reconnect, undo/redo, deletion, and geometry; receiver ${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: fixture(),
      documentId: `grid-split-${receiver}`,
      clientIds: [73, 41],
    });
    try {
      const [alice, bob] = pair.editors,
        first = alice!.surface!,
        second = bob!.surface!;
      const initial = first.session.part(),
        tableId = tables(initial)[0]!.id,
        owner = rowCells(rows(initial)[0]!)[0]!;
      const firstId = firstParagraph(owner).id,
        unrelated = tables(initial)[1]!;
      const block = buildBlockIndex(first.session.currentPackage()).byId.get('owned')!;
      const control = findNode(initial, block.controlId)!,
        content = namedChild(control, 'sdtContent')!,
        ownedId = firstParagraph(content).id;
      const gridIds = nativeGridColumns(initial, tableId).map((node) => node.id);
      pair.pause();
      expect(
        first.applyAutomationOps(() => [
          {
            op: 'splitTableCell',
            tableId,
            cellId: owner.id,
            rows: 2,
            columns: 2,
            columnDistribution: 'existing-grid',
          },
        ]).rejected
      ).toBe(false);
      expect(
        second.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: ownedId, offset: 6, text: ' Bob' },
        ]).rejected
      ).toBe(false);
      pair.reconnect(receiver);
      await pair.settle();
      const assertSplit = () => {
        for (const editor of pair.editors) {
          const part = editor.surface!.session.part(),
            pkg = editor.surface!.session.currentPackage();
          expect(rows(part).map((row) => rowCells(row).length)).toEqual([2, 2]);
          expect(widths(part)).toEqual(['1200', '2400']);
          expect(nativeGridColumns(part, tableId).map((node) => node.id)).toEqual(gridIds);
          expect(paragraphTextOf(part, firstId)).toBe('Merged.');
          expect(paragraphTextOf(part, ownedId)).toBe('Owned. Bob');
          expect(rowCells(rows(part)[0]!)[0]!.id).toBe(owner.id);
          expect(buildBlockIndex(pkg).byId.get('owned')?.payload).toEqual(block.payload);
          expect(buildBlockIndex(pkg).byId.get('owned')?.controlId).toBe(block.controlId);
          expect(canonicalOoxmlFingerprint(findNode(part, unrelated.id)!)).toBe(
            canonicalOoxmlFingerprint(unrelated)
          );
        }
      };
      assertSplit();
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const part = editor.surface!.session.part();
        expect(rows(part).map((row) => rowCells(row).length)).toEqual([1, 1]);
        expect(paragraphTextOf(part, ownedId)).toBe('Owned. Bob');
        expect(widths(part)).toEqual(['1200', '2400']);
      }
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      assertSplit();
      second.setSelection({
        anchor: { paragraphId: ownedId, offset: 3 },
        head: { paragraphId: ownedId, offset: 3 },
      });
      await pair.settle();
      const remote = pair.rooms[0]!.session.remoteSelections().find(
        (selection) => selection.actorId === 'bob'
      )!;
      expect(Boolean(remote)).toBe(true);
      expect(remote.head.nodeId).toBe(ownedId);
      const remoteCaret = caretAt(first.layout(), {
        paragraphId: remote.head.nodeId,
        offset: remote.head.offset,
      });
      const localCaret = caretAt(second.layout(), { paragraphId: ownedId, offset: 3 });
      expect(remoteCaret).not.toBeNull();
      expect(remoteCaret).toEqual(localCaret);
      const freshCell = rowCells(rows(first.session.part())[0]!)[1]!,
        freshParagraph = firstParagraph(freshCell).id;
      second.setSelection({
        anchor: { paragraphId: freshParagraph, offset: 0 },
        head: { paragraphId: freshParagraph, offset: 0 },
      });
      await pair.settle();
      const freshRemote = pair.rooms[0]!.session.remoteSelections().find(
        (selection) => selection.actorId === 'bob'
      )!;
      expect(freshRemote.head.nodeId).toBe(freshParagraph);
      const freshCaret = caretAt(first.layout(), { paragraphId: freshParagraph, offset: 0 });
      expect(freshCaret).not.toBeNull();
      expect(freshCaret).toEqual(
        caretAt(second.layout(), { paragraphId: freshParagraph, offset: 0 })
      );
      expect(freshCaret!.x).toBeGreaterThan(
        caretAt(first.layout(), { paragraphId: firstId, offset: 0 })!.x
      );
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
      const removedCell = rowCells(rows(first.session.part())[0]!)[1]!,
        removedParagraph = firstParagraph(removedCell).id;
      pair.pause();
      expect(
        first.applyAutomationOps(() => [
          { op: 'deleteTableColumn', tableId, gridColumnId: gridIds[1]! },
        ]).rejected
      ).toBe(false);
      expect(
        second.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: removedParagraph, offset: 0, text: 'Peer deleted cell' },
        ]).rejected
      ).toBe(false);
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        const part = editor.surface!.session.part();
        expect(rows(part).map((row) => rowCells(row).length)).toEqual([1, 1]);
        expect(findNode(part, removedCell.id)).toBeNull();
        expect(paragraphTextOf(part, ownedId)).toBe('Owned. Bob');
      }
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const part = editor.surface!.session.part();
        expect(rows(part).map((row) => rowCells(row).length)).toEqual([2, 2]);
        expect(paragraphTextOf(part, removedParagraph)).toBe('Peer deleted cell');
      }
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(findNode(editor.surface!.session.part(), removedCell.id)).toBeNull();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
