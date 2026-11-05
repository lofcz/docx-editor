// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, type OoxmlNode } from '@docx-editor.dev/core/store';
import { selectCellRectangle } from '../../../core/src/editor/__tests__/paginated-surface-fixtures';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';

const children = (node: OoxmlNode) => (node.kind === 'textValue' ? [] : node.children);

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`own proposed table reshaping retains peer edits and history (${clients})`, async () => {
    const pair = await createTwoEditorFixture({
      document: docxFixture('<w:p><w:r><w:t>Tail</w:t></w:r></w:p>'),
      clientIds: clients,
    });
    try {
      const [alice, bob] = pair.editors;
      const a = alice!.surface!,
        b = bob!.surface!;
      a.setEditingMode('suggest');
      expect(alice!.exec({ type: 'insertTable', rows: 2, cols: 2 }).ok).toBe(true);
      await pair.settle();
      const table = children(a.session.part().root)
        .flatMap(children)
        .find((n) => n.kind === 'table')!;
      const rows = () =>
        children(
          children(a.session.part().root)
            .flatMap(children)
            .find((n) => n.id === table.id)!
        ).filter((n) => n.kind === 'tableRow');
      const peerCell = children(rows()[1]!).filter((n) => n.kind === 'tableCell')[0]!;
      const peerParagraph = children(peerCell).find((n) => n.kind === 'paragraph')!.id;
      const at = { paragraphId: peerParagraph, offset: 0 };
      b.setSelection({ anchor: at, head: at });
      pair.pause();
      selectCellRectangle(a, { row: 0, column: 0 }, { row: 0, column: 1 });
      expect(alice!.exec({ type: 'mergeCells' }).ok).toBe(true);
      b.insertPlainText('peer');
      b.flushPendingInput();
      pair.reconnect(clients[0] === 41 ? 0 : 1);
      await pair.settle();
      expect(children(rows()[0]!).filter((n) => n.kind === 'tableCell')).toHaveLength(1);
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(children(rows()[0]!).filter((n) => n.kind === 'tableCell')).toHaveLength(2);
      expect(paragraphTextOf(a.session.part(), peerParagraph)).toBe('peer');
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();

      pair.pause();
      selectCellRectangle(a, { row: 1, column: 0 }, { row: 1, column: 0 });
      expect(alice!.exec({ type: 'splitCell', rows: 2, cols: 1 }).ok).toBe(true);
      expect(
        b.applyAutomationOps(() => [
          {
            op: 'deleteText',
            paragraphId: peerParagraph,
            start: 0,
            end: 1,
          },
        ]).rejected
      ).toBe(false);
      pair.reconnect(clients[0] === 41 ? 1 : 0);
      await pair.settle();
      expect(rows()).toHaveLength(3);
      expect(paragraphTextOf(a.session.part(), peerParagraph)).toBe('eer');
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(rows()).toHaveLength(2);
      expect(paragraphTextOf(a.session.part(), peerParagraph)).toBe('eer');
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(paragraphTextOf(a.session.part(), peerParagraph)).toBe('peer');
      expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const surface = editor.surface!;
        expect(paragraphTextOf(surface.session.part(), peerParagraph)).toBe('eer');
        expect(caretAt(surface.layout(), { paragraphId: peerParagraph, offset: 3 })).not.toBeNull();
      }
      expect(
        pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob')?.head.nodeId
      ).toBe(peerParagraph);
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      const reopened = await createTwoEditorFixture({
        document: new Uint8Array(await alice!.save()),
      });
      try {
        expect(reopened.editors[0]!.surface!.session.bodyText()).toContain('eer');
        const roundtrip = await reopened.savedDigests();
        expect(roundtrip[0]).toEqual(saved[0]);
        expect(roundtrip[1]).toEqual(saved[1]);
      } finally {
        reopened.destroy();
      }
    } finally {
      pair.destroy();
    }
  });
