import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  parentNodeOf,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { targetDocumentInput } from '../../test/support/target-document';
import { duplicateBlockSelection } from './duplicate';
import { removeBlockSelection } from './remove';
import { resolveExactBlockSelection } from './selection-exact';
import { moveBlockSelection } from './selection-move';

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const ordinary = ['A', 'B', 'C', 'D', 'E'].map(paragraph).join('');
async function bytes(body = ordinary) {
  return writeOoxmlPackage((await targetDocumentInput(body)).pkg);
}
function mount(documentBytes: Uint8Array) {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: documentBytes });
  return {
    editor,
    destroy: () => {
      editor.destroy();
      container.remove();
    },
  };
}
const part = (editor: DocxEditorInstance) => editor.surface!.session.part();
const texts = (editor: DocxEditorInstance) =>
  editor.surface!.session.paragraphIds().map((id) => paragraphTextOf(part(editor), id));
const fingerprint = (editor: DocxEditorInstance) => canonicalOoxmlFingerprint(part(editor).root);
const idFor = (editor: DocxEditorInstance, text: string) =>
  editor.surface!.session.paragraphIds().find((id) => paragraphTextOf(part(editor), id) === text)!;
const selected = (editor: DocxEditorInstance) => ['A', 'D'].map((text) => idFor(editor, text));
const unusedId = () => {
  throw Error('Ordinary paragraphs do not allocate payload IDs.');
};
const perform = (
  editor: DocxEditorInstance,
  action: 'copy' | 'move' | 'delete',
  nodeIds = selected(editor)
) => {
  if (action === 'copy')
    return duplicateBlockSelection(editor, { nodeIds, allocateBlockId: unusedId });
  if (action === 'delete') return removeBlockSelection(editor, { nodeIds });
  return moveBlockSelection(editor, {
    nodeIds,
    parentId: bodyStoryRoot(part(editor))!.id,
    beforeId: null,
  });
};

for (const action of ['copy', 'move', 'delete'] as const)
  test(`disjoint ${action} retains every unselected gap and uses one history entry`, async () => {
    const { editor, destroy } = mount(await bytes());
    try {
      const before = fingerprint(editor),
        originalIds = selected(editor);
      const untouched = ['B', 'C', 'E'].map((text) => findNode(part(editor), idFor(editor, text)));
      const result = perform(editor, action, [...originalIds].reverse());
      if (!result.success) throw Error(result.error);
      const expected =
        action === 'copy'
          ? ['A', 'B', 'C', 'D', 'A', 'D', 'E']
          : action === 'move'
            ? ['B', 'C', 'E', 'A', 'D']
            : ['B', 'C', 'E'];
      expect(texts(editor)).toEqual(expected);
      for (const node of untouched) expect(findNode(part(editor), node!.id)).toBe(node);
      if (action === 'move')
        expect(editor.surface!.session.paragraphIds().slice(-2)).toEqual(originalIds);
      if ('nodeIds' in result) {
        expect(result.nodeIds).toHaveLength(2);
        expect(result.nodeIds.some((id) => originalIds.includes(id))).toBe(false);
      }
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(editor)).toBe(before);
      expect(editor.surface!.session.canUndo()).toBe(false);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      const reopened = mount(new Uint8Array(await editor.save()));
      try {
        expect(texts(reopened.editor)).toEqual(expected);
      } finally {
        reopened.destroy();
      }
    } finally {
      destroy();
    }
  });

test('exact selection normalizes ancestor overlap and document order without filling gaps', async () => {
  const body = `${paragraph('A')}<w:sdt><w:sdtPr/><w:sdtContent>${paragraph('Inside')}</w:sdtContent></w:sdt>${paragraph('C')}${paragraph('D')}`;
  const { editor, destroy } = mount(await bytes(body));
  try {
    const child = idFor(editor, 'Inside');
    const control = parentNodeOf(part(editor), parentNodeOf(part(editor), child)!.id)!;
    const result = resolveExactBlockSelection(part(editor), [
      idFor(editor, 'D'),
      child,
      control.id,
      child,
    ]);
    expect(result.success && result.nodes.map((node) => node.id)).toEqual([
      control.id,
      idFor(editor, 'D'),
    ]);
    const before = fingerprint(editor);
    expect(
      removeBlockSelection(editor, { nodeIds: [idFor(editor, 'D'), control.id, child] }).success
    ).toBe(true);
    expect(texts(editor)).toEqual(['A', 'C']);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(editor)).toBe(before);
    expect(
      resolveExactBlockSelection(
        part(editor),
        Array.from({ length: 4097 }, () => child)
      ).success
    ).toBe(false);
    expect(resolveExactBlockSelection(part(editor), [child, 'missing']).success).toBe(false);
    expect(fingerprint(editor)).toBe(before);
  } finally {
    destroy();
  }
});

for (const action of ['copy', 'move', 'delete'] as const)
  test(`disjoint ${action} handles distinct native owners and retains editable source cells`, async () => {
    const body = `${paragraph('A')}<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${paragraph('Left')}</w:tc><w:tc><w:tcPr/>${paragraph('Right')}</w:tc></w:tr></w:tbl>${paragraph('E')}`;
    const { editor, destroy } = mount(await bytes(body));
    try {
      const before = fingerprint(editor);
      const nodeIds = [idFor(editor, 'A'), idFor(editor, 'Right')];
      const cellId = parentNodeOf(part(editor), nodeIds[1]!)!.id;
      const result = perform(editor, action, nodeIds);
      if (!result.success) throw Error(result.error);
      expect(texts(editor)).toEqual(
        action === 'copy'
          ? ['A', 'Left', 'Right', 'A', 'Right', 'E']
          : action === 'move'
            ? ['Left', '', 'E', 'A', 'Right']
            : ['Left', '', 'E']
      );
      const cell = findNode(part(editor), cellId)!;
      expect(storyParagraphs(cell)).toHaveLength(action === 'copy' ? 3 : 1);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(editor)).toBe(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      const reopened = mount(new Uint8Array(await editor.save()));
      try {
        expect(texts(reopened.editor)).toEqual(texts(editor));
      } finally {
        reopened.destroy();
      }
    } finally {
      destroy();
    }
  });

test('disjoint removal retains a terminal paragraph after an unselected nested table', async () => {
  const table = (body: string) =>
    `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${body}</w:tc></w:tr></w:tbl>`;
  const { editor, destroy } = mount(
    await bytes(
      `${paragraph('A')}${table(`${paragraph('Keep')}${table(paragraph('Nested'))}${paragraph('Remove')}`)}${paragraph('E')}`
    )
  );
  try {
    const before = fingerprint(editor),
      cellId = parentNodeOf(part(editor), idFor(editor, 'Remove'))!.id;
    const result = removeBlockSelection(editor, {
      nodeIds: [idFor(editor, 'A'), idFor(editor, 'Remove')],
    });
    if (!result.success) throw Error(result.error);
    const cell = findNode(part(editor), cellId)!;
    expect(cell.kind !== 'textValue' && cell.children.at(-1)?.kind).toBe('paragraph');
    expect(texts(editor)).toEqual(['Keep', 'Nested', '', 'E']);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(editor)).toBe(before);
  } finally {
    destroy();
  }
});

test('a protected later owner refuses disjoint commands before any history or content changes', async () => {
  const body = `${paragraph('A')}${paragraph('Gap')}<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>${paragraph('Locked')}</w:sdtContent></w:sdt>${paragraph('E')}`;
  const { editor, destroy } = mount(await bytes(body));
  try {
    const before = fingerprint(editor);
    for (const action of ['copy', 'move', 'delete'] as const) {
      expect(perform(editor, action, [idFor(editor, 'A'), idFor(editor, 'Locked')]).success).toBe(
        false
      );
      expect(fingerprint(editor)).toBe(before);
      expect(editor.surface!.session.canUndo()).toBe(false);
    }
  } finally {
    destroy();
  }
});

test('deleting all selected text keeps only required cell paragraphs without adding an outer gap', async () => {
  const body = `${paragraph('A')}<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${paragraph('Cell')}</w:tc></w:tr></w:tbl>`;
  const { editor, destroy } = mount(await bytes(body));
  try {
    const before = fingerprint(editor);
    expect(
      removeBlockSelection(editor, { nodeIds: [idFor(editor, 'A'), idFor(editor, 'Cell')] }).success
    ).toBe(true);
    expect(texts(editor)).toEqual(['']);
    const root = bodyStoryRoot(part(editor))!;
    expect(
      root.kind !== 'textValue' && root.children.filter((node) => node.kind === 'paragraph')
    ).toHaveLength(0);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(editor)).toBe(before);
  } finally {
    destroy();
  }
});

test('deleting the last text retains an editable position beside an empty native control', async () => {
  const { editor, destroy } = mount(
    await bytes(`${paragraph('A')}<w:sdt><w:sdtPr/><w:sdtContent/></w:sdt>`)
  );
  try {
    const before = fingerprint(editor);
    const removed = removeBlockSelection(editor, { nodeIds: [idFor(editor, 'A')] });
    if (!removed.success) throw Error(removed.error);
    expect(texts(editor)).toEqual(['']);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(editor)).toBe(before);
  } finally {
    destroy();
  }
});

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const action of ['copy', 'move', 'delete'] as const)
    test(`disjoint ${action} converges with peer text, selection, history, deletion, and reopen (${clientIds})`, async () => {
      const pair = await createTwoEditorFixture({ document: await bytes(), clientIds });
      try {
        const [alice, bob] = pair.editors;
        const originalIds = selected(alice!),
          gap = idFor(bob!, 'B');
        const point = { paragraphId: originalIds[1]!, offset: 1 };
        bob!.surface!.setSelection({ anchor: point, head: point });
        await pair.settle();
        pair.pause();
        const result = perform(alice!, action);
        if (!result.success) throw Error(result.error);
        expect(
          bob!.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: originalIds[1]!, offset: 1, text: ' peer' },
            { op: 'insertText', paragraphId: gap, offset: 1, text: ' gap' },
          ]).committed
        ).toBe(true);
        pair.reconnect(clientIds[0] === 41 ? 0 : 1);
        await pair.settle();
        const expected =
          action === 'copy'
            ? ['A', 'B gap', 'C', 'D peer', 'A', 'D', 'E']
            : action === 'move'
              ? ['B gap', 'C', 'E', 'A', 'D peer']
              : ['B gap', 'C', 'E'];
        for (const editor of pair.editors) expect(texts(editor)).toEqual(expected);
        if (action !== 'delete') {
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (entry) => entry.actorId === 'bob'
          );
          expect(remote?.head.nodeId).toBe(originalIds[1]);
        }
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(texts(editor)).toEqual(['A', 'B gap', 'C', 'D peer', 'E']);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) expect(texts(editor)).toEqual(expected);
        if ('nodeIds' in result) {
          expect(result.nodeIds.every((id) => !!findNode(part(bob!), id))).toBe(true);
          expect(removeBlockSelection(bob!, { nodeIds: [result.nodeIds[0]!] }).success).toBe(true);
          await pair.settle();
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(texts(editor)).toEqual(expected);
        }
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
        for (const editor of pair.editors) {
          const reopened = mount(new Uint8Array(await editor.save()));
          try {
            expect(texts(reopened.editor)).toEqual(expected);
          } finally {
            reopened.destroy();
          }
        }
      } finally {
        pair.destroy();
      }
    });
