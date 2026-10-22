import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  serializeOoxmlPart,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { CELL_CONTEXT, cellViews } from '../../test/support/table-review';
import { applyDirectPlan } from '../ai/apply-direct';
import type { NativeEditFailure } from '../contracts/document';
import { planInsertedTableColumn } from '../diff/inserted-table-columns';
import { createEditPlan, type CanonicalEditPlan } from '../diff/plan';
import { planTableColumn, type TableColumnTarget } from '../diff/table-columns';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTableColumnPlan } from './track-table-columns';
import type { TrackTextContext } from './tracked-inline';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, width: number, extra = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr>${p(text)}</w:tc>`;
/** Four columns with merged neighbours: a horizontal span left and a vertical merge right. */
const mergedDocument = () =>
  docxFixture(
    '<w:tbl><w:tblPr><w:tblW w:w="9600" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/></w:tblGrid>' +
      `<w:tr>${cell('A', 3600, '<w:gridSpan w:val="2"/>')}${cell('B', 2400)}${cell('C', 3600, '<w:vMerge w:val="restart"/>')}</w:tr>` +
      `<w:tr>${cell('D', 1200)}${cell('E', 2400)}${cell('F', 2400)}${cell('', 3600, '<w:vMerge/>')}</w:tr>` +
      '</w:tbl>' +
      p('Tail')
  );
const mount = (bytes = mergedDocument(), review = true): DocxEditorInstance =>
  createDocxEditor({
    container: document.createElement('div'),
    document: bytes,
    author: 'User',
    modules: review ? [SCIOBOT_NATIVE_REVIEW_MODULE] : [],
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
function blocks(text: string): TargetAstNode[] {
  const parsed = parseTargetMarkdown(text);
  if (!parsed.success) throw new Error(parsed.error);
  return parsed.target.ast.children as TargetAstNode[];
}
function richColumnPlan(editor: DocxEditorInstance, widthTwips?: number) {
  const session = editor.surface!.session,
    part = session.part(),
    [first, second] = rowsOf(part);
  const result = planInsertedTableColumn(
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
        { rowId: first!.id, blocks: blocks('**New** [one](https://example.com/n)') },
        { rowId: second!.id, blocks: blocks('Two  \nlines') },
      ],
      ...(widthTwips === undefined ? {} : { properties: { widthTwips } }),
    }
  );
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
function columnPlan(editor: DocxEditorInstance, target: TableColumnTarget) {
  const part = editor.surface!.session.part();
  const result = planTableColumn(
    { revision: revisionOf(editor), part, tableId: tableOf(part).id },
    target
  );
  if (!result.success) throw new Error(result.error);
  return result.plan;
}
/** The production direct executor (A01) applies the same plan to a second mount. */
function direct(plan: CanonicalEditPlan): OoxmlPart {
  const editor = mount(mergedDocument(), false);
  try {
    const result = applyDirectPlan(editor, plan, CELL_CONTEXT);
    if (!result.success) throw new Error(`${result.code}: ${result.error}`);
    return editor.surface!.session.part();
  } finally {
    editor.destroy();
  }
}
const comparable = (part: OoxmlPart) =>
  serializeOoxmlPart(part).replace(/ w14:(paraId|textId)="[0-9A-F]+"/g, '');
const cellTexts = (part: OoxmlPart) =>
  rowsOf(part).map((row) =>
    children(row)
      .filter((node) => node.kind === 'tableCell')
      .map((cell) =>
        children(cell)
          .filter((node) => node.kind === 'paragraph')
          .map((node) => paragraphTextOf(part, node.id))
          .join('|')
      )
  );
/** A valid grid: every row covers all grid columns, and cell widths sum to the grid. */
function validGrid(part: OoxmlPart) {
  const xml = serializeOoxmlPart(part);
  const gridWidths = [...xml.matchAll(/<w:gridCol w:w="(\d+)"\/>/g)].map((m) => Number(m[1]));
  const total = gridWidths.reduce((sum, width) => sum + width, 0);
  for (const row of xml.match(/<w:tr\b.*?<\/w:tr>/g)!) {
    const spans = [...row.matchAll(/<w:tc>(.*?)<\/w:tc>/g)].map((m) =>
      Number(/<w:gridSpan w:val="(\d+)"/.exec(m[1]!)?.[1] ?? 1)
    );
    expect(spans.reduce((sum, span) => sum + span, 0)).toBe(gridWidths.length);
    const widths = [...row.matchAll(/<w:tcW w:type="dxa" w:w="(\d+)"/g)].map((m) => Number(m[1]));
    expect(widths.reduce((sum, width) => sum + width, 0)).toBe(total);
  }
}

for (const action of ['accept', 'reject'] as const)
  test(`a rich middle column beside merged neighbours ${action}s to the direct output or the exact original`, async () => {
    const editor = mount();
    try {
      const session = editor.surface!.session,
        original = session.part(),
        originalPrint = canonicalOoxmlFingerprint(original);
      const plan = richColumnPlan(editor, 1800);
      const expected = direct(plan);
      const [committed] = await cellViews(editor);
      const result = trackTableColumnPlan(editor, plan, CELL_CONTEXT);
      if (!result.success) throw new Error(`${result.code}: ${result.error}`);
      expect(result.proposals).toHaveLength(1);
      expect(result.proposals[0]).toMatchObject({
        kind: 'columnInsert',
        tableId: tableOf(original).id,
      });
      expect(result.proposals[0]!.cellIds).toHaveLength(2);
      const xml = serializeOoxmlPart(session.part());
      expect(xml.match(/<w:cellIns /g)).toHaveLength(2);
      expect(xml).toContain('<w:tblGridChange');
      expect(xml).toContain('<w:tblPrChange');
      const views = await cellViews(editor);
      expect(views[0]).toBe(committed);
      expect(views[1]).toContain('**New** [one](https://example.com/n)');
      const pending = canonicalOoxmlFingerprint(session.part());
      expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
      expect(editor.getReviewItems()).toHaveLength(0);
      if (action === 'accept') {
        expect(comparable(session.part())).toBe(comparable(expected));
        expect(cellTexts(session.part())).toEqual([
          ['A', 'New one', 'B', 'C'],
          ['D', 'E', 'Two\nlines', 'F', ''],
        ]);
        validGrid(session.part());
      } else {
        expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
        validGrid(session.part());
      }
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(pending);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
      expect(session.canUndo()).toBe(false);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(pending);
    } finally {
      editor.destroy();
    }
  });

for (const action of ['accept', 'reject'] as const)
  test(`a middle column between merged neighbours deletes through ${action} to a valid grid`, async () => {
    const editor = mount();
    try {
      const session = editor.surface!.session,
        original = session.part(),
        originalPrint = canonicalOoxmlFingerprint(original);
      const plan = columnPlan(editor, { kind: 'delete', gridColumnId: gridOf(original)[2]!.id });
      const result = trackTableColumnPlan(editor, plan, CELL_CONTEXT);
      if (!result.success) throw new Error(`${result.code}: ${result.error}`);
      expect(result.proposals[0]).toMatchObject({ kind: 'columnDelete' });
      const textOf = (id: string) => {
        const cell = rowsOf(session.part())
          .flatMap((row) => children(row))
          .find((node) => node.id === id);
        return paragraphTextOf(
          session.part(),
          children(cell).find((n) => n.kind === 'paragraph')!.id
        );
      };
      expect(result.proposals[0]!.cellIds.map(textOf)).toEqual(['B', 'F']);
      const [committed, working] = await cellViews(editor);
      expect(committed).toContain('B');
      expect(working).not.toContain('| B |');
      expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
      validGrid(session.part());
      if (action === 'reject')
        expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
      else {
        expect(gridOf(session.part())).toHaveLength(3);
        expect(cellTexts(session.part())).toEqual([
          ['A', 'C'],
          ['D', 'E', ''],
        ]);
      }
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
    } finally {
      editor.destroy();
    }
  });

test('a pending column insertion survives save/reopen with its addresses and rejects exactly', async () => {
  const editor = mount();
  try {
    const originalPrint = canonicalOoxmlFingerprint(editor.surface!.session.part());
    const result = trackTableColumnPlan(editor, richColumnPlan(editor), CELL_CONTEXT);
    if (!result.success) throw new Error(result.error);
    const reopened = mount(new Uint8Array(await editor.save()));
    try {
      const addresses = reopened
        .surface!.session.reviewItems()
        .flatMap((item) => (item.kind === 'revision' ? item.addresses : []))
        .map((address) => address.id)
        .sort();
      expect(addresses).toEqual(result.proposals[0]!.addresses.map((address) => address.id).sort());
      expect(reopened.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(reopened.surface!.session.part())).toBe(originalPrint);
    } finally {
      reopened.destroy();
    }
  } finally {
    editor.destroy();
  }
});

test('unsupported and invalid column plans refuse atomically', () => {
  const editor = mount();
  try {
    const session = editor.surface!.session,
      part = session.part(),
      revision = session.packageRevision(),
      print = canonicalOoxmlFingerprint(part);
    const refuse = (
      plan: CanonicalEditPlan,
      code: NativeEditFailure['code'],
      context: TrackTextContext = CELL_CONTEXT
    ) => {
      const result = trackTableColumnPlan(editor, plan, context);
      expect(result.success).toBe(false);
      if (!result.success) expect(result.code).toBe(code);
      expect(session.packageRevision()).toBe(revision);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(print);
    };
    // The span over the first two grid columns cannot narrow through `w:cellDel`.
    refuse(
      columnPlan(editor, { kind: 'delete', gridColumnId: gridOf(part)[0]!.id }),
      'unsupported-operation'
    );
    refuse(richColumnPlan(editor), 'invalid-input', { ...CELL_CONTEXT, date: 'later' });
    const deletion = columnPlan(editor, { kind: 'delete', gridColumnId: gridOf(part)[2]!.id });
    refuse({ ...deletion, revision: { ...deletion.revision, revision: 999 } }, 'stale-revision');
    const two = createEditPlan(revisionOf(editor), [
      ...deletion.operations,
      ...columnPlan(editor, { kind: 'insert', where: 'left', gridColumnId: gridOf(part)[3]!.id })
        .operations,
    ]);
    if (!two.success) throw new Error(two.error);
    refuse(two.plan, 'unsupported-operation');
    expect(trackTableColumnPlan(editor, deletion, CELL_CONTEXT).success).toBe(true);
    // A second structure change waits for the first decision.
    const again = columnPlan(editor, {
      kind: 'insert',
      where: 'left',
      gridColumnId: gridOf(session.part())[3]!.id,
    });
    const pending = trackTableColumnPlan(editor, again, CELL_CONTEXT);
    expect(pending.success).toBe(false);
    if (!pending.success) expect(pending.code).toBe('unsupported-operation');
  } finally {
    editor.destroy();
  }
});

for (const mode of ['editing', 'suggesting'] as const)
  test(`column proposals preserve ${mode} mode and AI attribution`, () => {
    const editor = mount();
    try {
      expect(editor.setEditingMode(mode).ok).toBe(true);
      const result = trackTableColumnPlan(editor, richColumnPlan(editor, 1800), CELL_CONTEXT);
      if (!result.success) throw new Error(result.error);
      expect(
        result.proposals[0]!.addresses.every((entry) => ['ScioBot', ''].includes(entry.author))
      ).toBe(true);
      expect(editor.surface!.editingMode()).toBe(mode === 'editing' ? 'edit' : 'suggest');
    } finally {
      editor.destroy();
    }
  });
