import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  paragraphTextOf,
  serializeOoxmlPart,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { CELL_CONTEXT, cellViews } from '../../test/support/table-review';
import type { CanonicalEditPlan } from '../diff/plan';
import { planTableFormat, type TableFormatTarget } from '../diff/table-format';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTableFormatPlan } from './track-table-format';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, width: number, extra = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr>${p(text)}</w:tc>`;
const fixture = (clientIds: readonly [number, number]) =>
  createTwoEditorFixture({
    document: docxFixture(
      '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
        `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell('A', 2000, '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>')}${cell('B', 4000)}</w:tr>` +
        `<w:tr>${cell('One', 2000)}${cell('Two', 4000)}</w:tr>` +
        `<w:tr>${cell('Three', 2000)}${cell('Four', 4000)}</w:tr>` +
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
const rowsOf = (part: OoxmlPart) => children(tableOf(part)).filter((n) => n.kind === 'tableRow');
const cellIds = (part: OoxmlPart) =>
  rowsOf(part).map((row) =>
    children(row)
      .filter((n) => n.kind === 'tableCell')
      .map((n) => n.id)
  );
const paragraphNamed = (editor: DocxEditorInstance, text: string) => {
  const part = editor.surface!.session.part();
  return editor.surface!.session.paragraphIds().find((id) => paragraphTextOf(part, id) === text)!;
};
function plan(editor: DocxEditorInstance, target: (part: OoxmlPart) => TableFormatTarget) {
  const part = editor.surface!.session.part();
  const result = planTableFormat(
    {
      revision: readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
      part,
      tableId: tableOf(part).id,
    },
    target(part)
  );
  if (!result.success) throw new Error(result.error);
  return result.plan;
}
/** A background on the second row and wider first column: cell, grid, and table records. */
function propose(editor: DocxEditorInstance) {
  const fill = plan(editor, (part) => ({
    op: 'setTableCellFill',
    cellIds: [cellIds(part)[1]![1]!],
    color: { kind: 'hex', value: '00FF00' },
  }));
  const widths = plan(editor, () => ({
    op: 'setTableProperties',
    columnWidthsTwips: [3500, 2500],
  }));
  const combined = {
    ...fill,
    operations: [...fill.operations, ...widths.operations],
  } as CanonicalEditPlan;
  const result = trackTableFormatPlan(editor, combined, CELL_CONTEXT);
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
/** The table's property markup without paragraph content, as both peers must hold it. */
const formatOf = (editor: DocxEditorInstance) =>
  serializeOoxmlPart(editor.surface!.session.part())
    .replace(/^.*?<w:tbl>/, '<w:tbl>')
    .replace(/<\/w:tbl>.*$/, '</w:tbl>')
    .replace(/<w:p(?: [^>]*)?>.*?<\/w:p>|<w:p\/>/g, '');
async function same(pair: Awaited<ReturnType<typeof fixture>>) {
  const [a, b] = await Promise.all(pair.editors.map((editor) => cellViews(editor)));
  expect(a).toEqual(b);
  expect(formatOf(pair.editors[0]!)).toBe(formatOf(pair.editors[1]!));
  return { views: a!, format: formatOf(pair.editors[0]!) };
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    for (const action of ['accept', 'reject'] as const)
      test(`offline table format proposal keeps peer text through history and ${action}; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
        const pair = await fixture(clientIds);
        try {
          const [alice, bob] = pair.editors;
          const original = formatOf(alice!);
          pair.pause();
          propose(alice!);
          append(bob!, 'Two', ' Bob');
          append(bob!, 'A', ' Bob');
          pair.reconnect(first);
          await pair.settle();
          const pending = await same(pair);
          const [committed, working] = pending.views;
          expect(committed).toBe(working);
          expect(committed).toContain('| A Bob | B |');
          expect(committed).toContain('| One | Two Bob |');
          expect(pending.format).toContain('<w:tcPrChange ');
          expect(pending.format).toContain('<w:tblGridChange ');
          // Bob's caret in the reformatted cell is at the same native geometry in both sessions.
          const target = paragraphNamed(bob!, 'Two Bob');
          bob!.surface!.setSelection({
            anchor: { paragraphId: target, offset: 3 },
            head: { paragraphId: target, offset: 3 },
          });
          await pair.settle();
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          )!;
          expect(remote.head.nodeId).toBe(target);
          const carets = pair.editors.map((editor) =>
            caretAt(editor.surface!.layout(), { paragraphId: target, offset: 3 })
          );
          expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({
            x: carets[1]!.x,
            y: carets[1]!.y,
          });
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect((await same(pair)).format).toBe(original);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect((await same(pair)).format).toBe(pending.format);
          expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          await pair.settle();
          const resolved = await same(pair);
          expect(resolved.views).toEqual(pending.views);
          expect(resolved.format).not.toContain('Change ');
          if (action === 'reject') expect(resolved.format).toBe(original);
          else expect(resolved.format).toContain('w:fill="00FF00"');
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect((await same(pair)).format).toBe(pending.format);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect((await same(pair)).format).toBe(resolved.format);
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
  test(`a peer deleting the reformatted row while offline converges; clients ${clientIds.join('/')}`, async () => {
    const pair = await fixture(clientIds);
    try {
      const [alice, bob] = pair.editors;
      pair.pause();
      propose(alice!);
      const part = bob!.surface!.session.part();
      expect(
        bob!.surface!.applyAutomationOps(() => [
          { op: 'deleteTableRow', tableId: tableOf(part).id, rowId: rowsOf(part)[1]!.id },
        ]).rejected
      ).toBe(false);
      pair.reconnect(0);
      await pair.settle();
      const merged = await same(pair);
      expect(merged.views[0]).not.toContain('Two');
      expect(merged.views[0]).toContain('Three');
      // The surviving width records still decide together on both peers.
      expect(alice!.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
      await pair.settle();
      const rejected = await same(pair);
      expect(rejected.format).not.toContain('Change ');
      expect(rejected.format).toContain('<w:gridCol w:w="2000"/><w:gridCol w:w="4000"/>');
      const [savedAlice, savedBob] = await pair.savedDigests();
      expect(savedAlice).toEqual(savedBob);
    } finally {
      pair.destroy();
    }
  }, 60_000);
