import { expect, test } from 'bun:test';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  parentNodeOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import {
  namedChild,
  nodeChildren,
  rowCells,
  tableRows,
} from '../../test/support/native-table-oracle';
import { buildBlockIndex } from '../blocks/block-index';
import { withPayloadRecord } from '../blocks/payload-opc-codec';

const p = (value: string) => `<w:p><w:r><w:t>${value}</w:t></w:r></w:p>`;
const cell = (body: string) => `<w:tc><w:tcPr/>${body}</w:tc>`;
function fixture() {
  const loaded = readOoxmlPackage(
    docxFixture(
      p('Outside.') +
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="2400"/></w:tblGrid>' +
        `<w:tr>${cell(p('Upper.'))}${cell(p('Top right.'))}</w:tr>` +
        `<w:tr>${cell(p('Lower.') + '<w:sdt><w:sdtPr><w:tag w:val="scio:1:moved"/></w:sdtPr><w:sdtContent>' + p('Owned.') + '</w:sdtContent></w:sdt>')}${cell(p('Lower right.'))}</w:tr>` +
        `<w:tr>${cell(p('Keep.'))}${cell(p('Keep right.'))}</w:tr></w:tbl>` +
        p('Between.') +
        `<w:tbl><w:tblGrid><w:gridCol w:w="3600"/></w:tblGrid><w:tr>${cell(p('Unrelated.'))}</w:tr></w:tbl>` +
        p('Tail.')
    )
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const stored = withPayloadRecord(
    loaded.package,
    {
      id: 'moved',
      type: 'lesson_activity',
      version: 1,
      attributes: { name: 'Moved', duration: 5, asset: { keep: true } },
      unknownRoot: ['retain'],
    },
    'Moved'
  );
  if (!stored.ok) throw new Error(stored.reason);
  return writeOoxmlPackage(stored.pkg);
}
const tables = (part: OoxmlPart) =>
  nodeChildren(bodyStoryRoot(part)!).filter((node) => node.kind === 'table');
const paragraph = (node: OoxmlNode) =>
  nodeChildren(node).find((node) => node.kind === 'paragraph')!;
function assertIndex(part: OoxmlPart) {
  const ids = new Set<string>();
  const visit = (node: OoxmlNode, parent: OoxmlNode | null) => {
    expect(ids.has(node.id)).toBe(false);
    ids.add(node.id);
    expect(findNode(part, node.id) === node).toBe(true);
    expect(parentNodeOf(part, node.id)?.id ?? null).toBe(parent?.id ?? null);
    for (const child of nodeChildren(node)) visit(child, node);
  };
  visit(part.root, null);
}

for (const receiver of [0, 1] as const)
  test(`merged subtree IDs stay editable and indexed through peer edits, history, reconnect, geometry, and deletion; receiver ${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: fixture(),
      documentId: `merge-index-${receiver}`,
      clientIds: [73, 41],
    });
    try {
      const [alice, bob] = pair.editors,
        first = alice!.surface!,
        second = bob!.surface!;
      const original = first.session.part(),
        table = tables(original)[0]!,
        other = tables(original)[1]!;
      const originalCells = tableRows(original, table.id).map(rowCells),
        owner = originalCells[0]![0]!;
      const lowerId = paragraph(originalCells[1]![0]!).id,
        rightId = paragraph(originalCells[0]![1]!).id;
      const record = buildBlockIndex(first.session.currentPackage()).byId.get('moved')!;
      const ownedId = paragraph(
        namedChild(findNode(original, record.controlId)!, 'sdtContent')!
      ).id;
      const merge = {
        op: 'mergeTableCells',
        tableId: table.id,
        cellIds: [owner.id, originalCells[1]![1]!.id],
      } as const;
      pair.pause();
      expect(first.applyAutomationOps(() => [merge]).rejected).toBe(false);
      expect(
        second.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: ownedId, offset: 6, text: ' Bob' },
        ]).rejected
      ).toBe(false);
      pair.reconnect(receiver);
      await pair.settle();
      const assertMerged = (lowerText = 'Lower.') => {
        for (const editor of pair.editors) {
          const surface = editor.surface!,
            part = surface.session.part();
          expect(tableRows(part, table.id).map((row) => rowCells(row).length)).toEqual([1, 1, 2]);
          expect(paragraphTextOf(part, lowerId)).toBe(lowerText);
          expect(paragraphTextOf(part, ownedId)).toBe('Owned. Bob');
          expect(parentNodeOf(part, lowerId)?.id).toBe(owner.id);
          expect(parentNodeOf(part, record.controlId)?.id).toBe(owner.id);
          expect(
            buildBlockIndex(surface.session.currentPackage()).byId.get('moved')?.payload
          ).toEqual(record.payload);
          expect(canonicalOoxmlFingerprint(findNode(part, other.id)!)).toBe(
            canonicalOoxmlFingerprint(other)
          );
          assertIndex(part);
        }
      };
      assertMerged();
      expect(
        second.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: lowerId, offset: 6, text: ' Later' },
        ]).rejected
      ).toBe(false);
      await pair.settle();
      for (const editor of pair.editors)
        expect(paragraphTextOf(editor.surface!.session.part(), lowerId)).toBe('Lower. Later');
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const part = editor.surface!.session.part();
        expect(tableRows(part, table.id).map((row) => rowCells(row).length)).toEqual([2, 2, 2]);
        expect(parentNodeOf(part, lowerId)?.id).toBe(originalCells[1]![0]!.id);
        expect(paragraphTextOf(part, ownedId)).toBe('Owned. Bob');
        expect(paragraphTextOf(part, lowerId)).toBe('Lower. Later');
        assertIndex(part);
      }
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      assertMerged('Lower. Later');
      second.setSelection({
        anchor: { paragraphId: lowerId, offset: 3 },
        head: { paragraphId: lowerId, offset: 3 },
      });
      await pair.settle();
      const remote = pair.rooms[0]!.session.remoteSelections().find(
        (selection) => selection.actorId === 'bob'
      )!;
      expect(remote.head.nodeId).toBe(lowerId);
      const caret = caretAt(first.layout(), { paragraphId: lowerId, offset: 3 });
      expect(caret).not.toBeNull();
      expect(caret).toEqual(caretAt(second.layout(), { paragraphId: lowerId, offset: 3 }));
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      pair.pause();
      expect(
        first.applyAutomationOps(() => [
          { op: 'joinParagraphs', firstId: rightId, secondId: lowerId },
        ]).rejected
      ).toBe(false);
      expect(
        second.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: lowerId, offset: 12, text: ' Peer' },
        ]).rejected
      ).toBe(false);
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        const part = editor.surface!.session.part();
        expect(findNode(part, lowerId)).toBeNull();
        expect(paragraphTextOf(part, rightId)).toBe('Top right.Lower. Later Peer');
        expect(paragraphTextOf(part, ownedId)).toBe('Owned. Bob');
        assertIndex(part);
      }
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const part = editor.surface!.session.part();
        expect(paragraphTextOf(part, rightId)).toBe('Top right.');
        expect(paragraphTextOf(part, lowerId)).toBe('Lower. Later Peer');
        assertIndex(part);
      }
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const final = await pair.savedDigests();
      expect(final[0]).toEqual(final[1]);
    } finally {
      pair.destroy();
    }
  });
