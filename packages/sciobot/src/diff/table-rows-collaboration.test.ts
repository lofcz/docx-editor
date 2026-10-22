import { expect, test } from 'bun:test';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  replaceNode,
  type OoxmlNode,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';

const p = (text: string) =>
  `<w:p><w:pPr><w:rPr><w:i/></w:rPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`;
const row = (a: string, b: string) =>
  `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="360"/></w:trPr>${cell(a)}${cell(b)}</w:tr>`;
const document = () =>
  docxFixture(
    p('Outside.') +
      '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
      row('A', 'B') +
      row('One', 'Two') +
      '</w:tbl>' +
      p('Tail.')
  );
const children = (node: OoxmlNode): readonly OoxmlNode[] =>
  node.kind === 'textValue' ? [] : node.children;
const child = (node: OoxmlNode, name: string) =>
  children(node).find((child) => child.kind !== 'textValue' && child.localName === name);
function table(part: OoxmlPart) {
  return children(bodyStoryRoot(part)!).find((node) => node.kind === 'table')!;
}
function rows(part: OoxmlPart) {
  return children(table(part)).filter((node) => node.kind === 'tableRow');
}
function cells(row: OoxmlNode) {
  return children(row).filter((node) => node.kind === 'tableCell');
}
function paragraph(cell: OoxmlNode) {
  return children(cell).find((node) => node.kind === 'paragraph')!;
}
function values(part: OoxmlPart) {
  return rows(part).map((row) =>
    cells(row).map((cell) => paragraphTextOf(part, paragraph(cell).id))
  );
}
function native(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const changed = applyTreeOp(part, op);
  if (!changed.ok) throw new Error(changed.reason);
  return changed.part;
}
/** Redo allocates fresh Word paragraph IDs. Align only these identities in the independent direct oracle. */
function withLandedRowIdentities(expected: OoxmlPart, pending: OoxmlPart): OoxmlPart {
  const w14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
  const identity = (attribute: { namespaceUri: string; localName: string }) =>
    attribute.namespaceUri === w14 && ['paraId', 'textId'].includes(attribute.localName);
  const targetCells = cells(rows(expected)[2]!),
    actualCells = cells(rows(pending)[2]!);
  let next = expected;
  targetCells.forEach((cell, index) => {
    const target = paragraph(cell),
      actual = paragraph(actualCells[index]!);
    const changed = replaceNode(next, target.id, {
      ...target,
      attributes: [
        ...target.attributes.filter((attribute) => !identity(attribute)),
        ...actual.attributes.filter(identity),
      ],
    });
    if (!changed.ok) throw new Error('Direct identity oracle failed');
    next = changed.part;
  });
  return next;
}

for (const receiver of [0, 1] as const)
  test(`fresh tracked row content converges with inherited formatting and peer edits; receiver ${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: document(),
      documentId: `fresh-row-${receiver}`,
      clientIds: [73, 41],
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors,
        first = alice!.surface!,
        second = bob!.surface!;
      const initial = first.session.part(),
        body = bodyStoryRoot(initial)!,
        outside = children(body)[0]!,
        anchor = rows(initial)[1]!;
      const oldParagraph = paragraph(cells(anchor)[1]!).id;
      second.setSelection({
        anchor: { paragraphId: oldParagraph, offset: 3 },
        head: { paragraphId: oldParagraph, offset: 3 },
      });
      await pair.settle();
      const action = {
        kind: 'existing',
        tableId: table(initial).id,
        mutation: {
          kind: 'insertRows',
          rowId: anchor.id,
          location: 'after',
          count: 1,
          values: [['Three', 'Four']],
        },
      } as const;
      const peerOp = {
        op: 'insertText',
        paragraphId: outside.id,
        offset: 8,
        text: ' Bob',
      } as const;
      let expected = native(native(initial, { op: 'authorTable', action }), peerOp);
      pair.pause();
      expect(
        first.applyAutomationOps(() => [
          {
            op: 'authorTable',
            action,
            revision: { author: 'Alice', date: '2026-10-19T12:00:00Z' },
          },
        ]).rejected
      ).toBe(false);
      expect(second.applyAutomationOps(() => [peerOp]).rejected).toBe(false);
      pair.reconnect(receiver);
      await pair.settle();
      const assertPending = () => {
        for (const editor of pair.editors) {
          const part = editor.surface!.session.part();
          expect(values(part)).toEqual([
            ['A', 'B'],
            ['One', 'Two'],
            ['Three', 'Four'],
          ]);
          expect(paragraphTextOf(part, outside.id)).toBe('Outside. Bob');
          expect(editor.getReviewItems().some((item) => item.kind === 'revision')).toBe(true);
          for (const cell of cells(rows(part)[2]!)) {
            const body = paragraph(cell),
              runs = children(body).filter((node) => node.kind === 'run');
            expect(runs).toHaveLength(1);
            expect(child(child(runs[0]!, 'rPr')!, 'i')).toBeDefined();
          }
        }
      };
      assertPending();
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(rows(editor.surface!.session.part())).toHaveLength(2);
        expect(paragraphTextOf(editor.surface!.session.part(), outside.id)).toBe('Outside. Bob');
      }
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      assertPending();
      const current = second.session.part(),
        newParagraph = paragraph(cells(rows(current)[2]!)[0]!).id;
      second.setSelection({
        anchor: { paragraphId: newParagraph, offset: 5 },
        head: { paragraphId: newParagraph, offset: 5 },
      });
      await pair.settle();
      const remote = pair.rooms[0]!.session.remoteSelections().find(
        (selection) => selection.actorId === 'bob'
      )!;
      expect(Boolean(remote)).toBe(true);
      expect(remote.head.nodeId).toBe(newParagraph);
      const remoteCaret = caretAt(first.layout(), {
          paragraphId: remote.head.nodeId,
          offset: remote.head.offset,
        }),
        peerCaret = caretAt(second.layout(), { paragraphId: newParagraph, offset: 5 });
      expect(remoteCaret !== null && peerCaret !== null).toBe(true);
      expect({ x: remoteCaret!.x, y: remoteCaret!.y }).toEqual({
        x: peerCaret!.x,
        y: peerCaret!.y,
      });
      const [pendingA, pendingB] = await pair.savedDigests();
      expect(pendingA).toEqual(pendingB);
      expected = withLandedRowIdentities(expected, first.session.part());
      expect(bob!.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
          canonicalOoxmlFingerprint(expected)
        );
        expect(editor.getReviewItems()).toHaveLength(0);
      }
      expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      assertPending();
      expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const [acceptedA, acceptedB] = await pair.savedDigests();
      expect(acceptedA).toEqual(acceptedB);
      const removable = rows(first.session.part())[2]!,
        peerParagraph = paragraph(cells(rows(second.session.part())[2]!)[0]!).id;
      pair.pause();
      expect(
        second.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: peerParagraph, offset: 5, text: ' peer' },
        ]).rejected
      ).toBe(false);
      expect(
        first.applyAutomationOps(() => [
          { op: 'deleteTableRow', tableId: table(first.session.part()).id, rowId: removable.id },
        ]).rejected
      ).toBe(false);
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        const part = editor.surface!.session.part();
        expect(values(part)).toEqual([
          ['A', 'B'],
          ['One', 'Two'],
        ]);
        expect(paragraphTextOf(part, outside.id)).toBe('Outside. Bob');
        expect(findNode(part, removable.id)).toBeNull();
      }
      const [deletedA, deletedB] = await pair.savedDigests();
      expect(deletedA).toEqual(deletedB);
    } finally {
      pair.destroy();
    }
  }, 30_000);
