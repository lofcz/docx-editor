import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  findNode,
  paragraphTextOf,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { CELL_CONTEXT, CELL_ORIGINAL, cellViews } from '../../test/support/table-review';
import { planInsertedTableRow } from '../diff/inserted-table-rows';
import { planTableRow } from '../diff/table-rows';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTableRowPlan } from './track-table-rows';

const p = (text: string) =>
  `<w:p><w:r><w:rPr><w:color w:val="123456"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string) => `<w:tc><w:tcPr><w:shd w:fill="EEEEEE"/></w:tcPr>${p(text)}</w:tc>`;
const row = (a: string, b: string) =>
  `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell(a)}${cell(b)}</w:tr>`;
const rowDocument = () =>
  docxFixture(
    '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      row('A', 'B') +
      row('One', 'Two') +
      '</w:tbl>' +
      p('Tail')
  );
const fixture = (clientIds: readonly [number, number]) =>
  createTwoEditorFixture({
    document: rowDocument(),
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    ownedAttributes: true,
    clientIds,
  });
const children = (node: OoxmlNode | null | undefined) =>
  node && node.kind !== 'textValue' ? node.children : [];
const tableOf = (part: OoxmlPart) =>
  children(part.root.children[0]).find((n) => n.kind === 'table')!;
const rowsOf = (part: OoxmlPart) => children(tableOf(part)).filter((n) => n.kind === 'tableRow');
const paragraphNamed = (editor: DocxEditorInstance, text: string) => {
  const part = editor.surface!.session.part();
  return editor.surface!.session.paragraphIds().find((id) => paragraphTextOf(part, id) === text)!;
};
const revisionOf = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!;
function cells(text: string): TargetAstNode[] {
  const parsed = parseTargetMarkdown(text);
  if (!parsed.success) throw new Error(parsed.error);
  return parsed.target.ast.children as TargetAstNode[];
}
function proposeRow(editor: DocxEditorInstance) {
  const session = editor.surface!.session,
    part = session.part();
  const plan = planInsertedTableRow(
    {
      revision: revisionOf(editor),
      pkg: session.currentPackage(),
      partName: part.name,
      tableId: tableOf(part).id,
    },
    { rowId: rowsOf(part)[1]!.id, where: 'below', cells: [cells('Three'), cells('Four')] }
  );
  if (!plan.success) throw new Error(plan.error);
  const result = trackTableRowPlan(editor, plan.plan, CELL_CONTEXT);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result;
}
function proposeDeletion(editor: DocxEditorInstance) {
  const part = editor.surface!.session.part();
  const plan = planTableRow(
    { revision: revisionOf(editor), part, tableId: tableOf(part).id },
    { kind: 'delete', rowId: rowsOf(part)[1]!.id }
  );
  if (!plan.success) throw new Error(plan.error);
  const result = trackTableRowPlan(editor, plan.plan, CELL_CONTEXT);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result;
}
const append = (editor: DocxEditorInstance, word: string, text: string) => {
  const id = paragraphNamed(editor, word);
  expect(
    editor.surface!.applyAutomationOps(() => [
      { op: 'insertText', paragraphId: id, offset: word.length, text },
    ]).rejected
  ).toBe(false);
};
const withBob = CELL_ORIGINAL.replace('Tail', 'Tail Bob');
const withRow = withBob.replace('| One | Two |', '| One | Two |\n| Three | Four |');

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    for (const action of ['accept', 'reject'] as const)
      test(`offline row proposal converges with peer text, geometry, history, and ${action}; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
        const pair = await fixture(clientIds);
        try {
          const [alice, bob] = pair.editors;
          pair.pause();
          const proposal = proposeRow(alice!);
          expect(proposal.proposals).toHaveLength(1);
          append(bob!, 'Tail', ' Bob');
          pair.reconnect(first);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(await cellViews(editor)).toEqual([withBob, withRow]);
            expect(editor.getReviewItems()).toHaveLength(1);
          }
          // Bob's caret inside the proposed row resolves to the same native geometry on both sides.
          const three = paragraphNamed(bob!, 'Three');
          expect(paragraphNamed(alice!, 'Three')).toBe(three);
          bob!.surface!.setSelection({
            anchor: { paragraphId: three, offset: 1 },
            head: { paragraphId: three, offset: 5 },
          });
          await pair.settle();
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          )!;
          expect(remote.head.nodeId).toBe(three);
          expect(remote.head.offset).toBe(5);
          const carets = pair.editors.map((editor) =>
            caretAt(editor.surface!.layout(), { paragraphId: three, offset: 5 })
          );
          expect(carets.every((caret) => caret !== null)).toBe(true);
          expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({
            x: carets[1]!.x,
            y: carets[1]!.y,
          });
          expect(
            pair.containers[0]!.querySelectorAll('.docx-remote-caret, .docx-remote-selection-rect')
              .length
          ).toBeGreaterThan(0);
          // Proposal history keeps the peer's text.
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(await cellViews(editor)).toEqual([withBob, withBob]);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(await cellViews(editor)).toEqual([withBob, withRow]);
          const [pendingAlice, pendingBob] = await pair.savedDigests();
          expect(pendingAlice).toEqual(pendingBob);
          expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          await pair.settle();
          const resolved = action === 'accept' ? withRow : withBob;
          for (const editor of pair.editors) {
            expect(await cellViews(editor)).toEqual([resolved, resolved]);
            expect(editor.getReviewItems()).toHaveLength(0);
          }
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(await cellViews(editor)).toEqual([withBob, withRow]);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(await cellViews(editor)).toEqual([resolved, resolved]);
          const [savedAlice, savedBob] = await pair.savedDigests();
          expect(savedAlice).toEqual(savedBob);
          const reopened = await createTwoEditorFixture({
            document: new Uint8Array(await alice!.save()),
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
            ownedAttributes: true,
          });
          try {
            for (const editor of reopened.editors)
              expect(await cellViews(editor)).toEqual([resolved, resolved]);
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
      test(`tracked row deletion during peer editing keeps peer text pending and ${action}s it; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
        const pair = await fixture(clientIds);
        try {
          const [alice, bob] = pair.editors;
          const rowId = rowsOf(alice!.surface!.session.part())[1]!.id;
          pair.pause();
          proposeDeletion(alice!);
          append(bob!, 'Two', ' Bob');
          pair.reconnect(first);
          await pair.settle();
          const original = CELL_ORIGINAL.replace('Two', 'Two Bob'),
            removed = original.replace('\n| One | Two Bob |', '');
          for (const editor of pair.editors) {
            expect(await cellViews(editor)).toEqual([original, removed]);
            expect(editor.getReviewItems()).toHaveLength(1);
          }
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(await cellViews(editor)).toEqual([original, original]);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          await pair.settle();
          const resolved = action === 'accept' ? removed : original;
          for (const editor of pair.editors) {
            expect(await cellViews(editor)).toEqual([resolved, resolved]);
            expect(findNode(editor.surface!.session.part(), rowId) !== null).toBe(
              action === 'reject'
            );
          }
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(await cellViews(editor)).toEqual([original, removed]);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
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
  test(`rejecting a proposed row while its proposer types inside it converges without the row; clients ${clientIds.join('/')}`, async () => {
    const pair = await fixture(clientIds);
    try {
      const [alice, bob] = pair.editors;
      proposeRow(alice!);
      await pair.settle();
      pair.pause();
      append(alice!, 'Three', ' more');
      append(alice!, 'Tail', ' Alice');
      expect(bob!.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
      pair.reconnect(0);
      await pair.settle();
      const expected = CELL_ORIGINAL.replace('Tail', 'Tail Alice');
      for (const editor of pair.editors) {
        expect(await cellViews(editor)).toEqual([expected, expected]);
        expect(editor.getReviewItems()).toHaveLength(0);
      }
      const [savedAlice, savedBob] = await pair.savedDigests();
      expect(savedAlice).toEqual(savedBob);
    } finally {
      pair.destroy();
    }
  }, 60_000);
