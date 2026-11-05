import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  contentControlContentNodeOf,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { applyDirectPlan, previewDirectPlan } from '../ai/apply-direct';
import { planColumnWidths } from '../diff/column-widths';
import { readEditorRevision } from '../projection/revision-cache';
import { planTargetDocument } from '../diff/plan-target';
import { buildBlockIndex } from './block-index';
import { columnWidthsTwips, readColumnGeometry } from './column-geometry';
import { setColumnCount } from './set-column-count';
import { enterColumn } from './enter-column';

async function fixture(count: number) {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Tail</w:t></w:r></w:p>');
  let id = 0;
  const plan = planTargetDocument(
    input,
    `<column_group>\n${Array.from({ length: count }, (_, i) => `<column width="${100 / count}%">\n\nColumn ${i + 1}\n\n</column>`).join('\n')}\n</column_group>\n\nTail`,
    { allocateBlockId: () => `count-${++id}` }
  );
  if (!plan.success) throw Error(plan.error);
  const applied = previewDirectPlan(input.pkg, plan.plan);
  if (!applied.success) throw Error(applied.error);
  return writeOoxmlPackage(applied.pkg);
}
const geometry = (editor: DocxEditorInstance) => {
  const pkg = editor.surface!.session.currentPackage();
  const group = [...buildBlockIndex(pkg).byId.values()].find(
    (block) => block.payload.type === 'column_group'
  )!;
  return readColumnGeometry(pkg, group)!;
};
const change = (editor: DocxEditorInstance, count: number) => {
  const result = setColumnCount(editor, {
    groupId: geometry(editor).group.id,
    count,
    expectedRevision: editor.surface!.session.packageRevision(),
  });
  if (!result.success) throw Error(result.error);
};
const text = (editor: DocxEditorInstance) =>
  editor
    .surface!.session.paragraphIds()
    .map((id) => paragraphTextOf(editor.surface!.session.part(), id));
function coherent(editor: DocxEditorInstance) {
  const grid = geometry(editor);
  expect(grid).not.toBeNull();
  const widths = grid.columns.map((column) => column.width);
  expect(
    columnWidthsTwips(
      grid.columns.map((column) => column.block.payload.attributes.width),
      widths.reduce((a, b) => a + b, 0)
    )
  ).toEqual(widths);
  const painted = editor
    .surface!.layout()
    .pages[0]!.fragments.find((block) => block.kind === 'table' && block.tableId === grid.tableId)!;
  expect(painted.box.width).toBeCloseTo(widths.reduce((a, b) => a + b, 0) / 20);
}

for (const [from, to] of [
  [2, 3],
  [4, 3],
  [3, 1],
  [3, 3],
] as const)
  test(`column count ${from} to ${to} retains native paragraphs, metadata, undo, and saved content`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: await fixture(from),
    });
    try {
      const before = semanticDigest(editor.surface!.session.currentPackage().parts.values());
      const paragraphs = editor.surface!.session.paragraphIds().map((id) => ({
        id,
        value: canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), id)!),
      }));
      change(editor, to);
      expect(geometry(editor).columns).toHaveLength(to);
      coherent(editor);
      for (const { id, value } of paragraphs)
        expect(canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), id)!)).toBe(
          value
        );
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(semanticDigest(editor.surface!.session.currentPackage().parts.values())).toEqual(
        before
      );
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!reopened.ok) throw Error(reopened.reason);
      expect(semanticDigest(reopened.package.parts.values())).toEqual(
        semanticDigest(editor.surface!.session.currentPackage().parts.values())
      );
      expect((await projectSourcePackage(reopened.package)).text).toContain(`Column ${from}`);
    } finally {
      editor.destroy();
    }
  });

for (const first of [0, 1] as const)
  for (const to of [3, 1])
    test(`shared count change to ${to} preserves peer typing and history; receiver ${first}`, async () => {
      const pair = await createTwoEditorFixture({ document: await fixture(2) });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const paragraph = bob
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Column 2')!;
        pair.pause();
        change(alice, to);
        bob.surface!.setSelection({
          anchor: { paragraphId: paragraph, offset: 8 },
          head: { paragraphId: paragraph, offset: 8 },
        });
        bob.surface!.type(' Bob');
        pair.reconnect(first);
        await pair.settle();
        for (const editor of pair.editors) {
          coherent(editor);
          expect(text(editor)).toContain('Column 2 Bob');
        }
        expect(
          pair.rooms[0]!.session.remoteSelections().find((selection) => selection.actorId === 'bob')
            ?.head.nodeId
        ).toBe(paragraph);
        expect(
          caretAt(alice.surface!.layout(), { paragraphId: paragraph, offset: 12 })
        ).not.toBeNull();
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          coherent(editor);
          expect(text(editor)).toContain('Column 2 Bob');
          expect(geometry(editor).columns).toHaveLength(2);
        }
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
      } finally {
        pair.destroy();
      }
    });

for (const first of [0, 1] as const)
  for (const counts of [
    [3, 3],
    [3, 4],
    [1, 1],
    [1, 3],
    [3, 7],
  ] as const)
    test(`concurrent count changes ${counts.join('/')} preserve a coherent column model; receiver ${first}`, async () => {
      const pair = await createTwoEditorFixture({ document: await fixture(2) });
      try {
        pair.pause();
        for (const [index, editor] of pair.editors.entries()) change(editor, counts[index]!);
        pair.reconnect(first);
        await pair.settle();
        for (const editor of pair.editors) coherent(editor);
        const alice = pair.editors[0]!;
        const current = geometry(alice);
        if (current.columns.length > 1) {
          const widths = current.columns.map((column) => column.width);
          widths[0]! += 120;
          widths[1]! -= 120;
          const context = { documentId: 'count-resize', epoch: 'mount' };
          const resize = planColumnWidths(
            {
              pkg: alice.surface!.session.currentPackage(),
              groupId: current.group.id,
              revision: readEditorRevision(alice, context.documentId, context.epoch)!,
            },
            widths
          );
          if (!resize.success) throw Error(resize.error);
          const applied = applyDirectPlan(alice, resize.plan, context);
          if (!applied.success) throw Error(applied.error);
          await pair.settle();
          for (const editor of pair.editors) coherent(editor);
        }
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
      } finally {
        pair.destroy();
      }
    });

test('count changes refuse stale requests and unsupported editing modes without changing content', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: await fixture(2),
  });
  try {
    const surface = editor.surface!,
      before = surface.session.currentPackage();
    const request = {
      groupId: geometry(editor).group.id,
      count: 3,
      expectedRevision: surface.session.packageRevision(),
    };
    expect(setColumnCount(editor, { ...request, expectedRevision: -1 })).toMatchObject({
      success: false,
      code: 'stale-revision',
    });
    for (const count of [0, 1.5, 64, Number.NaN])
      expect(setColumnCount(editor, { ...request, count })).toMatchObject({
        success: false,
        code: 'invalid-input',
      });
    for (const mode of ['suggest', 'view'] as const) {
      surface.setEditingMode(mode);
      expect(setColumnCount(editor, request)).toMatchObject({
        success: false,
        code: 'unsupported-operation',
      });
      expect(surface.editingMode()).toBe(mode);
      expect(surface.session.currentPackage()).toBe(before);
    }
    surface.setEditingMode('edit');
    change(editor, 3);
    const revision = surface.session.packageRevision();
    change(editor, 3);
    expect(surface.session.packageRevision()).toBe(revision);
  } finally {
    editor.destroy();
  }
});

for (const first of [0, 1] as const)
  test(`column reduction retains a concurrent deletion through undo and redo; receiver ${first}`, async () => {
    const pair = await createTwoEditorFixture({ document: await fixture(3) });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const id = bob
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Column 3')!;
      pair.pause();
      change(alice, 2);
      expect(
        bob.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: id }]).committed
      ).toBe(true);
      pair.reconnect(first);
      await pair.settle();
      for (const editor of pair.editors) {
        coherent(editor);
        expect(text(editor)).not.toContain('Column 3');
        const part = editor.surface!.session.part();
        for (const column of geometry(editor).columns) {
          const content = contentControlContentNodeOf(findNode(part, column.block.controlId)!)!;
          expect(storyParagraphs(content).length).toBeGreaterThan(0);
        }
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        coherent(editor);
        expect(text(editor)).not.toContain('Column 3');
      }
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      // A peer deletion remains deleted. Entering the restored empty column creates a
      // fresh native paragraph only on explicit input, without resurrecting peer content.
      const empty = geometry(alice).columns.find(
        (column) =>
          !storyParagraphs(
            contentControlContentNodeOf(
              findNode(alice.surface!.session.part(), column.block.controlId)!
            )!
          ).length
      )!;
      expect(empty).toBeDefined();
      const entered = enterColumn(alice, {
        columnId: empty.block.id,
        expectedRevision: alice.surface!.session.packageRevision(),
      });
      if (!entered.success) throw Error(entered.error);
      alice.surface!.type('Restored column');
      await pair.settle();
      for (const editor of pair.editors) {
        expect(text(editor)).not.toContain('Column 3');
        expect(text(editor)).toContain('Restored column');
        expect(
          caretAt(editor.surface!.layout(), { paragraphId: entered.paragraphId, offset: 0 })
        ).not.toBeNull();
      }
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
