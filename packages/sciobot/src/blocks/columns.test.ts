import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  paragraphTextOf,
  findNode,
  applyTreeOp,
  withPart,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { strFromU8 } from 'fflate';
import { targetDocumentInput } from '../../test/support/target-document';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { previewDirectPlan, applyDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from '../diff/plan-target';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { buildBlockIndex } from './block-index';
import { columnWidthsTwips, readColumnGeometry } from './column-geometry';

// Each column keeps its native control inside an owned borderless table cell.
const column = (width: string, body: string) => `<column width="${width}">\n\n${body}\n\n</column>`;
const group = (...columns: string[]) =>
  `<column_group>\n\n${columns.join('\n\n')}\n\n</column_group>`;
const TABLE = '| A | B |\n| --- | --- |\n| 1 | 2 |';
const TWO = group(column('50%', `Left text.\n\n${TABLE}`), column('50%', 'Right text.'));
const THREE = group(
  column('33%', 'First column.'),
  column('34%', `Second column.\n\n- Item one\n- Item two`),
  column('33%', 'Third column.')
);

/** Body text in document order: the order a Word reader sees. */
const wordOrder = (pkg: OoxmlPackage) =>
  [
    ...strFromU8(pkg.partBytes.get(pkg.mainDocumentPart)!).matchAll(
      /<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g
    ),
  ]
    .map((match) => match[1])
    .filter((text) => text !== 'Seed.');

let ids = 0;
async function land(target: string) {
  return (await landWith(target)).pkg;
}
async function landWith(target: string) {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Seed.</w:t></w:r></w:p>');
  const planned = planTargetDocument(input, `${target}\n\nSeed.`, {
    allocateBlockId: () => `block-${++ids}`,
  });
  if (!planned.success) throw Error(planned.error);
  const applied = previewDirectPlan(input.pkg, planned.plan);
  if (!applied.success) throw Error(applied.error);
  return { input, pkg: applied.pkg };
}

for (const [name, target, widths, order] of [
  [
    'empty',
    group(column('33.33%', ''), column('33.33%', ''), column('33.33%', '')),
    ['33.33%', '33.33%', '33.33%'],
    [],
  ],
  ['two', TWO, ['50%', '50%'], ['Left text.', 'A', 'B', '1', '2', 'Right text.']],
  [
    'three',
    THREE,
    ['33%', '34%', '33%'],
    ['First column.', 'Second column.', 'Item one', 'Item two', 'Third column.'],
  ],
] as const)
  test(`${name} columns keep nested content, widths, and Word order through save and reopen`, async () => {
    const { input, pkg } = await landWith(target);
    const blocks = [...buildBlockIndex(pkg).byId.values()];
    expect(blocks.filter((block) => block.payload.type === 'column_group')).toHaveLength(1);
    expect(
      blocks
        .filter((block) => block.payload.type === 'column')
        .map((block) => block.payload.attributes.width)
    ).toEqual([...widths]);
    const owner = blocks.find((block) => block.payload.type === 'column_group')!;
    const geometry = readColumnGeometry(pkg, owner)!;
    expect(geometry.columns.map((column) => column.block.payload.attributes.width)).toEqual([
      ...widths,
    ]);
    expect(findNode(pkg.parts.get(pkg.mainDocumentPart)!, geometry.tableId)?.kind).toBe('table');
    expect(geometry.columns.map((column) => column.width)).toEqual([
      ...columnWidthsTwips(
        widths,
        geometry.columns.reduce((sum, column) => sum + column.width, 0)
      ),
    ]);
    const projection = await projectSourcePackage(pkg);
    expect(projection.text).toContain('<column_group>');
    for (const width of widths) expect(projection.text).toContain(`<column width="${width}">`);
    const reopened = readOoxmlPackage(writeOoxmlPackage(pkg));
    if (!reopened.ok) throw Error(reopened.reason);
    expect((await projectSourcePackage(reopened.package)).text).toBe(projection.text);
    // Word reads cells in column order and retains the nested table.
    expect(wordOrder(reopened.package)).toEqual([...order]);
    // The reopened document replans its own target to no change.
    const again = planTargetDocument(
      { ...input, pkg: reopened.package, projection: await projectSourcePackage(reopened.package) },
      `${target}\n\nSeed.`
    );
    if (!again.success) throw Error(again.error);
    expect(again.plan.operations).toEqual([]);
  });

const context = { documentId: 'columns-doc', epoch: 'mount-1' };
const token = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, context.documentId, context.epoch)!;
async function plan(editor: DocxEditorInstance, target: string) {
  const view = await readWorkingView(editor.surface!.session, token(editor));
  if (!view.success) throw Error(view.error);
  const result = planTargetDocument(
    {
      revision: token(editor),
      pkg: editor.surface!.session.currentPackage(),
      projection: view.view.projection,
    },
    target
  );
  if (!result.success) throw Error(`${result.code}: ${result.error}`);
  return result.plan;
}
const texts = (editor: DocxEditorInstance) =>
  editor
    .surface!.session.paragraphIds()
    .map((id) => paragraphTextOf(editor.surface!.session.part(), id) ?? '');
const geometryOf = (editor: DocxEditorInstance) => {
  const pkg = editor.surface!.session.currentPackage();
  const block = [...buildBlockIndex(pkg).byId.values()].find(
    (block) => block.payload.type === 'column_group'
  )!;
  return readColumnGeometry(pkg, block)!;
};

test('column reading order stays complete across page fragments', async () => {
  const left = Array.from({ length: 80 }, (_, index) => `Left paragraph ${index}.`).join('\n\n');
  const right = Array.from({ length: 65 }, (_, index) => `Right paragraph ${index}.`).join('\n\n');
  const pkg = await land(group(column('50%', left), column('50%', right)));
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(pkg),
  });
  try {
    expect(editor.surface!.layout().pages.length).toBeGreaterThan(1);
  } finally {
    editor.destroy();
  }
  const projected = await projectSourcePackage(pkg);
  expect(projected.text).toContain('Left paragraph 79');
  expect(projected.text).toContain('Right paragraph 0');
  expect(projected.text.indexOf('Left paragraph 79')).toBeLessThan(
    projected.text.indexOf('Right paragraph 0')
  );
  expect(
    [...projected.text.matchAll(/Left paragraph (\d+)/g)].map((match) => Number(match[1]))
  ).toEqual(Array.from({ length: 80 }, (_, index) => index));
  expect(
    [...projected.text.matchAll(/Right paragraph (\d+)/g)].map((match) => Number(match[1]))
  ).toEqual(Array.from({ length: 65 }, (_, index) => index));
  const reopened = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!reopened.ok) throw Error(reopened.reason);
  expect((await projectSourcePackage(reopened.package)).text).toBe(projected.text);
});

test('column widths reject unsupported values without rounding away the grid total', () => {
  expect(columnWidthsTwips(['33.33%', '33.33%', '33.33%'], 9360)).toEqual([3120, 3120, 3120]);
  expect(columnWidthsTwips(['30%', '70%'], 9361).reduce((sum, width) => sum + width, 0)).toBe(9361);
  for (const widths of [
    [],
    ['0%', '100%'],
    ['200px'],
    ['Infinity%'],
    ['0.001%', '99.999%'],
    Array(65).fill('1%'),
  ])
    expect(() => columnWidthsTwips(widths, 9360)).toThrow();
});

test('column ownership refuses a separator that contains user content', async () => {
  const pkg = await land(TWO);
  const group = [...buildBlockIndex(pkg).byId.values()].find(
    (block) => block.payload.type === 'column_group'
  )!;
  const geometry = readColumnGeometry(pkg, group)!;
  const changed = applyTreeOp(pkg.parts.get(pkg.mainDocumentPart)!, {
    op: 'insertText',
    paragraphId: [...geometry.separatorIds][0]!,
    offset: 0,
    text: 'Keep this content',
  });
  if (!changed.ok) throw Error(changed.reason);
  expect(readColumnGeometry(withPart(pkg, changed.part), group)).toBeNull();
  expect(readColumnGeometry(pkg, group)?.tableId).toBe(geometry.tableId);
});

test('a column edit is one undo step', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: writeOoxmlPackage(await land(TWO)),
    author: 'Teacher',
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
  try {
    const before = texts(editor);
    const target = group(column('50%', `Left text.\n\n${TABLE}`), column('50%', 'Right changed.'));
    const applied = applyDirectPlan(editor, await plan(editor, `${target}\n\nSeed.`), context);
    if (!applied.success) throw Error(applied.error);
    expect(texts(editor)).toContain('Right changed.');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(editor)).toEqual(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(texts(editor)).toContain('Right changed.');
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw Error(reopened.reason);
    expect(wordOrder(reopened.package)).toEqual([
      'Left text.',
      'A',
      'B',
      '1',
      '2',
      'Right changed.',
    ]);
  } finally {
    editor.destroy();
    container.remove();
  }
});

for (const first of [0, 1] as const)
  test(`two sessions converge on column resize and editing beside peer typing; receiver ${first}`, async () => {
    const pair = await createTwoEditorFixture({ document: writeOoxmlPackage(await land(TWO)) });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const leftId = bob
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Left text.')!;
      const target = group(
        column('30%', `Left text.\n\n${TABLE}`),
        column('70%', 'Right changed.')
      );
      const beforeGeometry = geometryOf(alice);
      const originalWidths = beforeGeometry.columns.map((column) => column.width);
      const originalRightId = alice
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(alice.surface!.session.part(), id) === 'Right text.')!;
      const rightBefore = caretAt(alice.surface!.layout(), {
        paragraphId: originalRightId,
        offset: 0,
      })!;
      const leftBefore = caretAt(alice.surface!.layout(), { paragraphId: leftId, offset: 0 })!;
      expect(rightBefore.y).toBeCloseTo(leftBefore.y, 2);
      expect(rightBefore.x).toBeGreaterThan(leftBefore.x);
      const planned = await plan(alice, `${target}\n\nSeed.`);
      pair.pause();
      const applied = applyDirectPlan(alice, planned, context);
      if (!applied.success) throw Error(applied.error);
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: leftId, offset: 10, text: ' Bob' },
        ]).committed
      ).toBe(true);
      pair.reconnect(first);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(texts(editor)).toContain('Left text. Bob');
        expect(texts(editor)).toContain('Right changed.');
        const geometry = geometryOf(editor);
        expect(geometry.columns.map((column) => column.block.payload.attributes.width)).toEqual([
          '30%',
          '70%',
        ]);
        expect(geometry.columns.map((column) => column.width)).toEqual([
          ...columnWidthsTwips(
            ['30%', '70%'],
            originalWidths.reduce((sum, width) => sum + width, 0)
          ),
        ]);
        expect(geometry.columns.map((column) => column.block.controlId)).toEqual(
          beforeGeometry.columns.map((column) => column.block.controlId)
        );
        expect(
          caretAt(editor.surface!.layout(), { paragraphId: originalRightId, offset: 0 })!.x
        ).toBeLessThan(rightBefore.x);
      }
      expect(texts(alice)).toEqual(texts(bob));
      // Bob's caret inside the first column is visible to Alice at the right place.
      bob.surface!.setSelection({
        anchor: { paragraphId: leftId, offset: 14 },
        head: { paragraphId: leftId, offset: 14 },
      });
      await pair.settle();
      const remote = pair.rooms[0]!.session.remoteSelections().find(
        (selection) => selection.actorId === 'bob'
      )!;
      expect(remote.head.nodeId).toBe(leftId);
      const caret = caretAt(alice.surface!.layout(), { paragraphId: leftId, offset: 14 }),
        start = caretAt(alice.surface!.layout(), { paragraphId: leftId, offset: 0 });
      expect(caret!.x).toBeGreaterThan(start!.x);
      // Alice's undo keeps Bob's text; redo restores her edit.
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(texts(editor)).toContain('Right text.');
        expect(texts(editor)).toContain('Left text. Bob');
        expect(geometryOf(editor).columns.map((column) => column.width)).toEqual(originalWidths);
        expect(
          geometryOf(editor).columns.map((column) => column.block.payload.attributes.width)
        ).toEqual(['50%', '50%']);
      }
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) expect(texts(editor)).toContain('Right changed.');
      // Bob deletes the edited column's paragraph text while Alice types elsewhere.
      pair.pause();
      const rightId = bob
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Right changed.')!;
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'deleteText', paragraphId: rightId, start: 0, end: 6 },
        ]).committed
      ).toBe(true);
      expect(
        alice.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: leftId, offset: 0, text: 'A ' },
        ]).committed
      ).toBe(true);
      pair.reconnect(first);
      await pair.settle();
      expect(texts(alice)).toEqual(texts(bob));
      expect(texts(alice)).toContain('changed.');
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      const reopened = readOoxmlPackage(new Uint8Array(await alice.save()));
      if (!reopened.ok) throw Error(reopened.reason);
      expect(
        [...buildBlockIndex(reopened.package).byId.values()].map((b) => b.payload.type)
      ).toEqual(expect.arrayContaining(['column_group', 'column', 'column']));
    } finally {
      pair.destroy();
    }
  });
