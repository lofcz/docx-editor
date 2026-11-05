import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  parentNodeOf,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { moveBlockSelection } from './selection-move';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const table = (body: string) =>
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${body}</w:tc></w:tr></w:tbl>`;
const fixture = () =>
  docxFixture(
    p('Before') + table(p('Source first') + table(p('Nested cell')) + p('Source last')) + p('After')
  );
const paragraph = (editor: DocxEditorInstance, text: string) =>
  editor
    .surface!.session.paragraphIds()
    .find((id) => paragraphTextOf(editor.surface!.session.part(), id) === text)!;
const sourceCell = (editor: DocxEditorInstance) => {
  const part = editor.surface!.session.part();
  const id = editor
    .surface!.session.paragraphIds()
    .find((id) => paragraphTextOf(part, id)?.startsWith('Source first'))!;
  const cell = parentNodeOf(part, id)!;
  if (cell.kind !== 'tableCell') throw Error('Missing source cell');
  return cell;
};
const text = (editor: DocxEditorInstance) =>
  editor
    .surface!.session.paragraphIds()
    .map((id) => paragraphTextOf(editor.surface!.session.part(), id));
const moveLast = (editor: DocxEditorInstance) =>
  moveBlockSelection(editor, {
    nodeIds: [paragraph(editor, 'Source last')],
    parentId: bodyStoryRoot(editor.surface!.session.part())!.id,
    beforeId: paragraph(editor, 'After'),
  });
const expectTerminalParagraph = (editor: DocxEditorInstance) => {
  const children = sourceCell(editor).children.filter((node) =>
    ['paragraph', 'table', 'contentControl'].includes(node.kind)
  );
  expect(children.map((node) => node.kind)).toEqual(['paragraph', 'table', 'paragraph']);
  expect(paragraphTextOf(editor.surface!.session.part(), children.at(-1)!.id)).toBe('');
};

test('moving the final paragraph from a nested-table cell retains an editable terminator and exact undo', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: fixture(),
  });
  try {
    const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
    const movedId = paragraph(editor, 'Source last');
    expect(moveLast(editor)).toMatchObject({ success: true, changed: true });
    expectTerminalParagraph(editor);
    expect(findNode(editor.surface!.session.part(), movedId)?.kind).toBe('paragraph');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const reopened = createDocxEditor({
      container: document.createElement('div'),
      document: new Uint8Array(await editor.save()),
    });
    try {
      expectTerminalParagraph(reopened);
      expect(text(reopened)).toEqual(text(editor));
    } finally {
      reopened.destroy();
    }
  } finally {
    editor.destroy();
  }
});

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`source-cell terminators preserve concurrent text, history, and reopen (${clientIds})`, async () => {
    const pair = await createTwoEditorFixture({ document: fixture(), clientIds });
    try {
      const [alice, bob] = pair.editors;
      const first = paragraph(bob!, 'Source first'),
        last = paragraph(bob!, 'Source last');
      pair.pause();
      expect(moveLast(alice!)).toMatchObject({ success: true, changed: true });
      expect(
        bob!.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: first, offset: 'Source first'.length, text: ' peer' },
          { op: 'insertText', paragraphId: last, offset: 'Source last'.length, text: ' peer' },
        ]).committed
      ).toBe(true);
      pair.reconnect(clientIds[0] === 41 ? 0 : 1);
      await pair.settle();
      for (const editor of pair.editors) {
        expectTerminalParagraph(editor!);
        expect(text(editor!)).toEqual([
          'Before',
          'Source first peer',
          'Nested cell',
          '',
          'Source last peer',
          'After',
        ]);
      }
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const end = sourceCell(editor!).children.at(-1)!;
        expect(paragraphTextOf(editor!.surface!.session.part(), end.id)).toBe('Source last peer');
      }
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      const reopened = createDocxEditor({
        container: document.createElement('div'),
        document: new Uint8Array(await alice!.save()),
      });
      try {
        expectTerminalParagraph(reopened);
        expect(text(reopened)).toEqual(text(alice!));
      } finally {
        reopened.destroy();
      }
    } finally {
      pair.destroy();
    }
  });
