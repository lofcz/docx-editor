import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, type OoxmlNode, type OoxmlPart } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { CELL_CONTEXT, cellViews } from '../../test/support/table-review';
import { planTableMerge, type TableMergeTarget } from '../diff/table-merges';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTableMergePlan } from './track-table-merges';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, width: number) =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`;
const row = (a: string, b: string, c: string) =>
  `<w:tr>${cell(a, 1200)}${cell(b, 2400)}${cell(c, 3600)}</w:tr>`;
const fixture = (clientIds: readonly [number, number]) =>
  createTwoEditorFixture({
    document: docxFixture(
      '<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/></w:tblGrid>' +
        row('A', 'B', 'C') +
        row('D', 'E', 'F') +
        row('G', 'H', 'I') +
        '</w:tbl>' +
        p('Tail')
    ),
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    ownedAttributes: true,
    clientIds,
  });
const children = (node: OoxmlNode | null | undefined) =>
  node && node.kind !== 'textValue' ? node.children : [];
const tableOf = (part: OoxmlPart) =>
  children(part.root.children[0]).find((n) => n.kind === 'table')!;
const ids = (part: OoxmlPart) =>
  children(tableOf(part))
    .filter((r) => r.kind === 'tableRow')
    .map((r) =>
      children(r)
        .filter((n) => n.kind === 'tableCell')
        .map((n) => n.id)
    );
const paragraphNamed = (editor: DocxEditorInstance, text: string) => {
  const part = editor.surface!.session.part();
  return editor.surface!.session.paragraphIds().find((id) => paragraphTextOf(part, id) === text)!;
};
function propose(editor: DocxEditorInstance, target: (cells: string[][]) => TableMergeTarget) {
  const part = editor.surface!.session.part();
  const plan = planTableMerge(
    {
      revision: readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
      part,
      tableId: tableOf(part).id,
    },
    target(ids(part))
  );
  if (!plan.success) throw new Error(plan.error);
  const result = trackTableMergePlan(editor, plan.plan, CELL_CONTEXT);
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
const rectangle = (cells: string[][]): TableMergeTarget => ({
  kind: 'merge',
  cellIds: [cells[0]![0]!, cells[0]![1]!, cells[1]![0]!, cells[1]![1]!],
});

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    for (const action of ['accept', 'reject'] as const)
      test(`offline merge proposal keeps owner and outside peer text through history and ${action}; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
        const pair = await fixture(clientIds);
        try {
          const [alice, bob] = pair.editors;
          pair.pause();
          propose(alice!, rectangle);
          append(bob!, 'A', ' Bob');
          append(bob!, 'I', ' Bob');
          pair.reconnect(first);
          await pair.settle();
          const [committed, working] = await same(pair);
          expect(committed).toContain('| A Bob | B | C |');
          expect(working.replace(/<br>/g, ' ')).toContain('A Bob B D E');
          expect(working).toContain('I Bob');
          const owner = paragraphNamed(bob!, 'A Bob');
          bob!.surface!.setSelection({
            anchor: { paragraphId: owner, offset: 5 },
            head: { paragraphId: owner, offset: 5 },
          });
          await pair.settle();
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          )!;
          expect(remote.head.nodeId).toBe(owner);
          const carets = pair.editors.map((editor) =>
            caretAt(editor.surface!.layout(), { paragraphId: owner, offset: 5 })
          );
          expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({
            x: carets[1]!.x,
            y: carets[1]!.y,
          });
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(await same(pair)).toEqual([committed, committed]);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(await same(pair)).toEqual([committed, working]);
          expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          await pair.settle();
          const resolved = action === 'accept' ? working : committed;
          expect(await same(pair)).toEqual([resolved, resolved]);
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(await same(pair)).toEqual([committed, working]);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(await same(pair)).toEqual([resolved, resolved]);
          const [savedAlice, savedBob] = await pair.savedDigests();
          expect(savedAlice).toEqual(savedBob);
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 60_000);

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const action of ['accept', 'reject'] as const)
    test(`peer text typed into a merged source cell survives rejection; acceptance keeps the proposal content; clients ${clientIds.join('/')}, ${action}`, async () => {
      const pair = await fixture(clientIds);
      try {
        const [alice, bob] = pair.editors;
        pair.pause();
        propose(alice!, rectangle);
        append(bob!, 'E', ' Bob');
        pair.reconnect(0);
        await pair.settle();
        const [committed] = await same(pair);
        expect(committed).toContain('| D | E Bob | F |');
        expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        await pair.settle();
        const [resolved] = await same(pair);
        // A merge proposal gathers the content it saw; the source cell itself is deleted on accept.
        if (action === 'reject') expect(resolved).toBe(committed);
        else expect(resolved).not.toContain('E Bob');
        const [savedAlice, savedBob] = await pair.savedDigests();
        expect(savedAlice).toEqual(savedBob);
      } finally {
        pair.destroy();
      }
    }, 60_000);

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const action of ['accept', 'reject'] as const)
    test(`offline split proposal keeps peer text in the split cell through ${action}; clients ${clientIds.join('/')}`, async () => {
      const pair = await fixture(clientIds);
      try {
        const [alice, bob] = pair.editors;
        pair.pause();
        propose(alice!, (cells) => ({ kind: 'split', cellId: cells[1]![1]!, rows: 2, columns: 2 }));
        append(bob!, 'E', ' Bob');
        pair.reconnect(1);
        await pair.settle();
        const [committed, working] = await same(pair);
        expect(committed).toContain('| D | E Bob | F |');
        expect(working).toContain('E Bob');
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(await same(pair)).toEqual([committed, committed]);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        await pair.settle();
        const resolved = action === 'accept' ? working : committed;
        expect(await same(pair)).toEqual([resolved, resolved]);
        const [savedAlice, savedBob] = await pair.savedDigests();
        expect(savedAlice).toEqual(savedBob);
      } finally {
        pair.destroy();
      }
    }, 60_000);
