import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, type OoxmlNode, type OoxmlPart } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { CELL_CONTEXT, cellViews } from '../../test/support/table-review';
import { planInsertedTableColumn } from '../diff/inserted-table-columns';
import { planTableColumn } from '../diff/table-columns';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTableColumnPlan } from './track-table-columns';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, width: number, extra = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr>${p(text)}</w:tc>`;
const mergedDocument = () =>
  docxFixture(
    '<w:tbl><w:tblPr><w:tblW w:w="9600" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/></w:tblGrid>' +
      `<w:tr>${cell('A', 3600, '<w:gridSpan w:val="2"/>')}${cell('B', 2400)}${cell('C', 3600, '<w:vMerge w:val="restart"/>')}</w:tr>` +
      `<w:tr>${cell('D', 1200)}${cell('E', 2400)}${cell('F', 2400)}${cell('', 3600, '<w:vMerge/>')}</w:tr>` +
      '</w:tbl>' +
      p('Tail')
  );
const ORIGINAL = '| A |  | B | C |\n| --- | --- | --- | --- |\n| D | E | F |  |\n\nTail';
const fixture = (clientIds: readonly [number, number]) =>
  createTwoEditorFixture({
    document: mergedDocument(),
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    ownedAttributes: true,
    clientIds,
  });
const children = (node: OoxmlNode | null | undefined) =>
  node && node.kind !== 'textValue' ? node.children : [];
const tableOf = (part: OoxmlPart) =>
  children(part.root.children[0]).find((n) => n.kind === 'table')!;
const rowsOf = (part: OoxmlPart) => children(tableOf(part)).filter((n) => n.kind === 'tableRow');
const gridOf = (part: OoxmlPart) =>
  children(
    children(tableOf(part)).find((n) => n.kind !== 'textValue' && n.localName === 'tblGrid')
  ).filter((n) => n.kind !== 'textValue' && n.localName === 'gridCol');
const revisionOf = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!;
const paragraphNamed = (editor: DocxEditorInstance, text: string) => {
  const part = editor.surface!.session.part();
  return editor.surface!.session.paragraphIds().find((id) => paragraphTextOf(part, id) === text)!;
};
function blocks(text: string): TargetAstNode[] {
  const parsed = parseTargetMarkdown(text);
  if (!parsed.success) throw new Error(parsed.error);
  return parsed.target.ast.children as TargetAstNode[];
}
function proposeColumn(editor: DocxEditorInstance) {
  const session = editor.surface!.session,
    part = session.part(),
    [first, second] = rowsOf(part);
  const plan = planInsertedTableColumn(
    {
      revision: revisionOf(editor),
      pkg: session.currentPackage(),
      partName: part.name,
      tableId: tableOf(part).id,
    },
    {
      anchor: { kind: 'grid-column', id: gridOf(part)[1]!.id },
      where: 'right',
      cells: [
        { rowId: first!.id, blocks: blocks('**New**') },
        { rowId: second!.id, blocks: blocks('Fresh') },
      ],
      properties: { widthTwips: 1800 },
    }
  );
  if (!plan.success) throw new Error(plan.error);
  const result = trackTableColumnPlan(editor, plan.plan, CELL_CONTEXT);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
}
function proposeDeletion(editor: DocxEditorInstance) {
  const part = editor.surface!.session.part();
  const plan = planTableColumn(
    { revision: revisionOf(editor), part, tableId: tableOf(part).id },
    { kind: 'delete', gridColumnId: gridOf(part)[2]!.id }
  );
  if (!plan.success) throw new Error(plan.error);
  const result = trackTableColumnPlan(editor, plan.plan, CELL_CONTEXT);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
}
const append = (editor: DocxEditorInstance, word: string, text: string) => {
  const id = paragraphNamed(editor, word);
  expect(
    editor.surface!.applyAutomationOps(() => [
      { op: 'insertText', paragraphId: id, offset: word.length, text },
    ]).rejected
  ).toBe(false);
};
async function same(pair: Awaited<ReturnType<typeof fixture>>) {
  const [a, b] = await Promise.all(pair.editors.map((editor) => cellViews(editor)));
  expect(a).toEqual(b);
  return a!;
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    for (const action of ['accept', 'reject'] as const)
      test(`offline column proposal converges with peer cell text, geometry, history, and ${action}; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
        const pair = await fixture(clientIds);
        try {
          const [alice, bob] = pair.editors;
          pair.pause();
          proposeColumn(alice!);
          append(bob!, 'E', ' Bob');
          append(bob!, 'Tail', ' Bob');
          pair.reconnect(first);
          await pair.settle();
          const original = ORIGINAL.replace('| E |', '| E Bob |').replace('Tail', 'Tail Bob');
          const [committed, working] = await same(pair);
          expect(committed).toBe(original);
          expect(working).toContain('**New**');
          expect(working).toContain('Fresh');
          expect(working).toContain('E Bob');
          const fresh = paragraphNamed(bob!, 'Fresh');
          expect(paragraphNamed(alice!, 'Fresh')).toBe(fresh);
          bob!.surface!.setSelection({
            anchor: { paragraphId: fresh, offset: 2 },
            head: { paragraphId: fresh, offset: 2 },
          });
          await pair.settle();
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          )!;
          expect(remote.head.nodeId).toBe(fresh);
          const carets = pair.editors.map((editor) =>
            caretAt(editor.surface!.layout(), { paragraphId: fresh, offset: 2 })
          );
          expect(carets.every((caret) => caret !== null)).toBe(true);
          expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({
            x: carets[1]!.x,
            y: carets[1]!.y,
          });
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(await same(pair)).toEqual([original, original]);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(await same(pair)).toEqual([committed, working]);
          const [pendingAlice, pendingBob] = await pair.savedDigests();
          expect(pendingAlice).toEqual(pendingBob);
          expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          await pair.settle();
          const resolved = action === 'accept' ? working : original;
          expect(await same(pair)).toEqual([resolved, resolved]);
          for (const editor of pair.editors) expect(editor.getReviewItems()).toHaveLength(0);
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(await same(pair)).toEqual([committed, working]);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(await same(pair)).toEqual([resolved, resolved]);
          const [savedAlice, savedBob] = await pair.savedDigests();
          expect(savedAlice).toEqual(savedBob);
          const reopened = await createTwoEditorFixture({
            document: new Uint8Array(await alice!.save()),
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
            ownedAttributes: true,
          });
          try {
            expect(await same(reopened)).toEqual([resolved, resolved]);
          } finally {
            reopened.destroy();
          }
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 60_000);

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    for (const action of ['accept', 'reject'] as const)
      test(`tracked column deletion while the peer edits a deleted cell ${action}s it; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
        const pair = await fixture(clientIds);
        try {
          const [alice, bob] = pair.editors;
          pair.pause();
          proposeDeletion(alice!);
          append(bob!, 'F', ' Bob');
          pair.reconnect(first);
          await pair.settle();
          const original = ORIGINAL.replace('| F |', '| F Bob |');
          const [committed, working] = await same(pair);
          expect(committed).toBe(original);
          expect(working).not.toContain('F Bob');
          expect(working).not.toContain('| B |');
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(await same(pair)).toEqual([original, original]);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          await pair.settle();
          const resolved = await same(pair);
          if (action === 'reject') expect(resolved).toEqual([original, original]);
          else {
            expect(resolved[0]).toBe(resolved[1]);
            expect(resolved[0]).not.toContain('F Bob');
          }
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(await same(pair)).toEqual([committed, working]);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(await same(pair)).toEqual(resolved);
          const [savedAlice, savedBob] = await pair.savedDigests();
          expect(savedAlice).toEqual(savedBob);
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 60_000);
