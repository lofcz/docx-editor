import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  serializeOoxmlPart,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import reference from '../__tests__/fixtures/ai-pairs.json';
import { docxFixture } from '../../test/support/document';
import { materializeInsertedRow as oracleRow } from '../../test/support/inserted-row-oracle';
import { CELL_CONTEXT, CELL_ORIGINAL, cellViews } from '../../test/support/table-review';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { readEditorRevision } from '../projection/revision-cache';
import { planInsertedTableRow } from '../diff/inserted-table-rows';
import { insertedPayload } from '../diff/inserted-blocks';
import { MDX_BLOCK } from '../projection/mdx';
import type { InsertedRowPayloadBinding } from '../diff/inserted-row-types';
import { createEditPlan, type CanonicalEditPlan } from '../diff/plan';
import { planTableRow, type TableRowTarget } from '../diff/table-rows';
import { applyDirectPlan } from '../ai/apply-direct';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTableRowPlan } from './track-table-rows';
import type { TrackTextContext } from './tracked-inline';
import type { NativeEditFailure } from '../contracts/document';

const p = (text: string) =>
  `<w:p><w:r><w:rPr><w:color w:val="123456"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, props = true) =>
  `<w:tc>${props ? '<w:tcPr><w:shd w:fill="EEEEEE"/></w:tcPr>' : ''}${p(text)}</w:tc>`;
const row = (a: string, b: string, props = true) =>
  `<w:tr>${props ? '<w:trPr><w:cantSplit/></w:trPr>' : ''}${cell(a, props)}${cell(b, props)}</w:tr>`;
/** The C01 table shape, with an imported last row that has no property holders. */
const rowDocument = (bare = false) =>
  docxFixture(
    '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      row('A', 'B') +
      row('One', 'Two', !bare) +
      '</w:tbl>' +
      p('Tail')
  );
const mount = (bytes = rowDocument(), review = true): DocxEditorInstance =>
  createDocxEditor({
    container: document.createElement('div'),
    document: bytes,
    author: 'User',
    modules: review ? [SCIOBOT_NATIVE_REVIEW_MODULE] : [],
  });
const C01_ROW = CELL_ORIGINAL.replace('| One | Two |', '| One | Two |\n| Three | Four |');
const children = (node: OoxmlNode | null | undefined) =>
  node && node.kind !== 'textValue' ? node.children : [];
const tableOf = (part: OoxmlPart) =>
  children(part.root.children[0]).find((n) => n.kind === 'table')!;
const rowsOf = (part: OoxmlPart) => children(tableOf(part)).filter((n) => n.kind === 'tableRow');
const revisionOf = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!;
function paragraphs(nodes: readonly string[]): TargetAstNode[] {
  const parsed = parseTargetMarkdown(nodes.join('\n\n'));
  if (!parsed.success) throw new Error(parsed.error);
  return parsed.target.ast.children as TargetAstNode[];
}
function rowPlan(
  editor: DocxEditorInstance,
  cells: readonly (readonly TargetAstNode[])[],
  where: 'above' | 'below' = 'below',
  anchor = 1,
  heightTwips?: number,
  extra: { readonly payloads?: readonly InsertedRowPayloadBinding[] } = {}
): CanonicalEditPlan {
  const session = editor.surface!.session,
    part = session.part();
  const result = planInsertedTableRow(
    {
      revision: revisionOf(editor),
      pkg: session.currentPackage(),
      partName: part.name,
      tableId: tableOf(part).id,
    },
    {
      rowId: rowsOf(part)[anchor]!.id,
      where,
      cells,
      ...extra,
      ...(heightTwips === undefined ? {} : { properties: { heightTwips } }),
    }
  );
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
function nativeRowPlan(editor: DocxEditorInstance, target: TableRowTarget) {
  const part = editor.surface!.session.part();
  const result = planTableRow(
    { revision: revisionOf(editor), part, tableId: tableOf(part).id },
    target
  );
  if (!result.success) throw new Error(result.error);
  return result.plan;
}
/** The production direct executor (A01) applies the same plan to a second mount of the same file. */
function direct(bytes: Uint8Array, plan: CanonicalEditPlan): OoxmlPart {
  const editor = mount(bytes, false);
  try {
    const result = applyDirectPlan(editor, plan, CELL_CONTEXT);
    if (!result.success) throw new Error(`${result.code}: ${result.error}`);
    return editor.surface!.session.part();
  } finally {
    editor.destroy();
  }
}
/** Fresh Word paragraph IDs differ per allocation; compare everything else exactly. */
const rowTexts = (part: OoxmlPart) =>
  rowsOf(part).map((row) =>
    children(row)
      .filter((node) => node.kind === 'tableCell')
      .map((cell) =>
        children(cell)
          .filter((node) => node.kind === 'paragraph')
          .map((node) => paragraphTextOf(part, node.id))
      )
  );
const comparable = (part: OoxmlPart) =>
  serializeOoxmlPart(part).replace(/ w14:(paraId|textId)="[0-9A-F]+"/g, '');

for (const action of ['accept', 'reject'] as const)
  test(`C01 added row ${action} matches direct output or the exact original, with one undo entry`, async () => {
    const captured = reference.cases.find((item) => item.name === 'tables and links')!;
    expect(captured.afterMarkdown).toContain('| Three | Four |');
    const editor = mount();
    try {
      const session = editor.surface!.session,
        original = session.part(),
        originalPrint = canonicalOoxmlFingerprint(original);
      const plan = rowPlan(editor, [paragraphs(['Three']), paragraphs(['Four'])]);
      const expected = direct(rowDocument(), plan);
      // The independent D16 fixture codec lands the same cell content.
      const independent = oracleRow(session.currentPackage(), plan).pkg.parts.get(original.name)!;
      expect(rowTexts(expected)).toEqual(rowTexts(independent));
      const result = trackTableRowPlan(editor, plan, CELL_CONTEXT);
      if (!result.success) throw new Error(`${result.code}: ${result.error}`);
      expect(result.proposals).toHaveLength(1);
      expect(result.proposals[0]).toMatchObject({ kind: 'rowInsert', partName: original.name });
      expect(
        result.proposals[0]!.addresses.every(
          (address) => address.author === 'ScioBot' && address.date === '2026-10-19T09:00:00.000Z'
        )
      ).toBe(true);
      expect(editor.getReviewItems()).toHaveLength(1);
      expect(editor.getReviewItems()[0]).toMatchObject({
        item: { structuralChanges: expect.arrayContaining(['rowInsert']) },
      });
      expect(await cellViews(editor)).toEqual([CELL_ORIGINAL, C01_ROW]);
      // Existing rows, grid, and table properties keep their objects.
      const pending = session.part();
      for (const node of rowsOf(original)) expect(findNode(pending, node.id)).toBe(node);
      const pendingPrint = canonicalOoxmlFingerprint(pending);
      expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
      if (action === 'accept') expect(comparable(session.part())).toBe(comparable(expected));
      else expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
      expect(editor.getReviewItems()).toHaveLength(0);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(pendingPrint);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
      expect(session.canUndo()).toBe(false);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(pendingPrint);
    } finally {
      editor.destroy();
    }
  });

for (const action of ['accept', 'reject'] as const)
  test(`imported row deletion without row/cell properties ${action}s exactly`, async () => {
    const editor = mount(rowDocument(true));
    try {
      const session = editor.surface!.session,
        original = session.part(),
        originalPrint = canonicalOoxmlFingerprint(original),
        target = rowsOf(original)[1]!;
      expect(serializeOoxmlPart(original)).not.toMatch(/<w:tr><w:trPr\/>|<w:tcPr\/>/);
      const plan = nativeRowPlan(editor, { kind: 'delete', rowId: target.id });
      const expected = direct(rowDocument(true), plan);
      const result = trackTableRowPlan(editor, plan, CELL_CONTEXT);
      if (!result.success) throw new Error(result.error);
      expect(result.proposals).toEqual([
        expect.objectContaining({ kind: 'rowDelete', rowIds: [target.id] }),
      ]);
      expect(editor.getReviewItems()[0]).toMatchObject({
        item: { structuralChanges: expect.arrayContaining(['rowDelete']) },
      });
      const [committed, working] = await cellViews(editor);
      expect(committed).toContain('| One | Two |');
      expect(working).not.toContain('| One | Two |');
      expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
      if (action === 'accept')
        expect(canonicalOoxmlFingerprint(session.part())).toBe(canonicalOoxmlFingerprint(expected));
      else {
        expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
        const restored = findNode(session.part(), target.id)!;
        expect(
          children(restored).some((n) => n.kind !== 'textValue' && n.localName === 'trPr')
        ).toBe(false);
      }
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
    } finally {
      editor.destroy();
    }
  });

test('a deleted row with imported properties keeps its holders through rejection', () => {
  const editor = mount();
  try {
    const session = editor.surface!.session,
      originalPrint = canonicalOoxmlFingerprint(session.part());
    const plan = nativeRowPlan(editor, { kind: 'delete', rowId: rowsOf(session.part())[1]!.id });
    const result = trackTableRowPlan(editor, plan, CELL_CONTEXT);
    if (!result.success) throw new Error(result.error);
    expect(serializeOoxmlPart(session.part())).toContain('<w:cantSplit/><w:del ');
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
  } finally {
    editor.destroy();
  }
});

test('native value rows and blank rows match direct native output when accepted', () => {
  for (const values of [['Three', '😀\tČtyři'], undefined] as const) {
    const editor = mount();
    try {
      const session = editor.surface!.session,
        original = session.part();
      const plan = nativeRowPlan(editor, {
        kind: 'insert',
        rowId: rowsOf(original)[1]!.id,
        where: 'above',
        ...(values ? { values } : {}),
      });
      const expected = direct(rowDocument(), plan);
      const result = trackTableRowPlan(editor, plan, CELL_CONTEXT);
      if (!result.success) throw new Error(result.error);
      expect(result.proposals).toHaveLength(1);
      expect(result.proposals[0]!.kind).toBe('rowInsert');
      expect(rowsOf(session.part())[1]!.id).toBe(result.proposals[0]!.rowIds[0]!);
      expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
      expect(comparable(session.part())).toBe(comparable(expected));
    } finally {
      editor.destroy();
    }
  }
});

test('a rich new row keeps marks, links, breaks, paragraphs, headings, and height through save/reopen', async () => {
  const editor = mount();
  try {
    const session = editor.surface!.session,
      original = session.part(),
      originalPrint = canonicalOoxmlFingerprint(original);
    const plan = rowPlan(
      editor,
      [
        paragraphs(['**Bold** and [link](https://example.com/a "Tip")', 'Second  \nline']),
        paragraphs(['## Heading']),
      ],
      'above',
      1,
      500
    );
    const expected = direct(rowDocument(), plan);
    const result = trackTableRowPlan(editor, plan, CELL_CONTEXT);
    if (!result.success) throw new Error(`${result.code}: ${result.error}`);
    const [newRowId] = result.proposals[0]!.rowIds;
    const xml = serializeOoxmlPart(session.part());
    expect(xml).toContain('<w:trHeight w:hRule="exact" w:val="500"/>');
    expect(xml).toMatch(/<w:hyperlink [^>]*w:tooltip="Tip"/);
    expect(xml).toContain('<w:br/>');
    expect(xml).toContain('<w:outlineLvl w:val="1"/>');
    const fresh = findNode(session.part(), newRowId!)!;
    const texts = children(fresh)
      .filter((n) => n.kind === 'tableCell')
      .map((cell) =>
        children(cell)
          .filter((n) => n.kind === 'paragraph')
          .map((n) => paragraphTextOf(session.part(), n.id))
      );
    expect(texts).toEqual([['Bold and link', 'Second\nline'], ['Heading']]);
    // Pending save/reopen keeps the row revision and content.
    const reopened = mount(new Uint8Array(await editor.save()));
    try {
      expect(reopened.surface!.session.reviewItems()).toEqual([
        expect.objectContaining({
          structuralChanges: expect.arrayContaining(['rowInsert']),
          author: 'ScioBot',
        }),
      ]);
      expect(
        reopened
          .surface!.session.reviewItems()
          .flatMap((item) => (item.kind === 'revision' ? item.addresses : []))
          .map((a) => [a.id, a.author, a.date])
      ).toEqual(result.proposals[0]!.addresses.map((a) => [a.id, a.author, a.date]));
      expect(reopened.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(reopened.surface!.session.part())).toBe(originalPrint);
    } finally {
      reopened.destroy();
    }
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
    expect(comparable(session.part())).toBe(comparable(expected));
  } finally {
    editor.destroy();
  }
});

test('row insertions and deletions in one plan form separate decisions in one undo entry', async () => {
  const editor = mount();
  try {
    const session = editor.surface!.session,
      part = session.part(),
      [first, second] = rowsOf(part);
    const plan = createEditPlan(revisionOf(editor), [
      ...nativeRowPlan(editor, { kind: 'delete', rowId: second!.id }).operations,
      ...rowPlan(editor, [paragraphs(['X']), paragraphs(['Y'])], 'above', 0).operations,
    ]);
    if (!plan.success) throw new Error(plan.error);
    const originalPrint = canonicalOoxmlFingerprint(part);
    const result = trackTableRowPlan(editor, plan.plan, CELL_CONTEXT);
    if (!result.success) throw new Error(result.error);
    expect(result.proposals.map((item) => item.kind)).toEqual(['rowDelete', 'rowInsert']);
    expect(editor.getReviewItems()).toHaveLength(2);
    expect(rowsOf(session.part()).map((node) => node.id)).toEqual([
      result.proposals[1]!.rowIds[0],
      first!.id,
      second!.id,
    ]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
    expect(session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});

test('unsupported, stale, and invalid row plans refuse atomically', () => {
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
      const result = trackTableRowPlan(editor, plan, context);
      expect(result.success).toBe(false);
      if (!result.success) expect(result.code).toBe(code);
      expect(session.packageRevision()).toBe(revision);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(print);
    };
    const plain = rowPlan(editor, [paragraphs(['A']), paragraphs(['B'])]);
    refuse(plain, 'invalid-input', { ...CELL_CONTEXT, date: 'yesterday' });
    refuse(plain, 'invalid-input', { ...CELL_CONTEXT, author: ' ' });
    const twice = createEditPlan(revisionOf(editor), [...plain.operations, ...plain.operations]);
    if (!twice.success) throw new Error(twice.error);
    refuse(twice.plan, 'ambiguous-target');
    const deletion = nativeRowPlan(editor, { kind: 'delete', rowId: rowsOf(part)[1]!.id });
    refuse({ ...deletion, revision: { ...deletion.revision, revision: 999 } }, 'stale-revision');
    const text = createEditPlan(revisionOf(editor), [
      {
        kind: 'text',
        partName: part.name,
        paragraphId: session.paragraphIds().at(-1)!,
        start: 0,
        end: 4,
        before: 'Tail',
        after: 'End',
      },
    ]);
    if (!text.success) throw new Error(text.error);
    refuse(text.plan, 'unsupported-operation');
    // A pending row proposal needs ownership reconciliation before another change.
    expect(trackTableRowPlan(editor, deletion, CELL_CONTEXT).success).toBe(true);
    const again = nativeRowPlan(editor, { kind: 'delete', rowId: rowsOf(part)[1]!.id });
    const pending = trackTableRowPlan(editor, again, CELL_CONTEXT);
    expect(pending.success).toBe(false);
    if (!pending.success) expect(pending.code).toBe('unsupported-operation');
  } finally {
    editor.destroy();
  }
  const bare = mount(rowDocument(), false);
  try {
    const part = bare.surface!.session.part();
    const plan = nativeRowPlan(bare, { kind: 'delete', rowId: rowsOf(part)[1]!.id });
    const result = trackTableRowPlan(bare, plan, CELL_CONTEXT);
    expect(result.success).toBe(false);
    if (!result.success) expect(result.code).toBe('unsupported-operation');
    expect(bare.surface!.session.part()).toBe(part);
  } finally {
    bare.destroy();
  }
});

for (const mode of ['editing', 'suggesting'] as const)
  test(`row proposals preserve ${mode} mode and AI attribution`, () => {
    const editor = mount();
    try {
      expect(editor.setEditingMode(mode).ok).toBe(true);
      const result = trackTableRowPlan(
        editor,
        rowPlan(editor, [paragraphs(['Three']), paragraphs(['Four'])]),
        CELL_CONTEXT
      );
      if (!result.success) throw new Error(result.error);
      expect(result.proposals[0]!.addresses.every((entry) => entry.author === 'ScioBot')).toBe(
        true
      );
      expect(editor.surface!.editingMode()).toBe(mode === 'editing' ? 'edit' : 'suggest');
    } finally {
      editor.destroy();
    }
  });

test('tracked rich row fragments keep code, nested tables, and application payloads under the row revision', () => {
  const editor = mount();
  try {
    const codeParsed = parseTargetMarkdown('```ts\nconst value = 1;\n```');
    const tableParsed = parseTargetMarkdown('| Nested |\n| --- |\n| cell |');
    const widgetParsed = parseTargetMarkdown('<Widget state={{ mode: "safe" }} />');
    if (!codeParsed.success || !tableParsed.success || !widgetParsed.success)
      throw new Error('Rich row fixture did not parse.');
    const code = codeParsed.target.ast.children[0]!;
    const nested = tableParsed.target.ast.children[0]!;
    const widget = widgetParsed.target.ast.children[0]!;
    const codePayload = insertedPayload(code, 'tracked-code');
    if (!codePayload) throw new Error('Code payload did not compile.');
    const payloads: InsertedRowPayloadBinding[] = [
      { path: [0, 0], label: 'Code', payload: codePayload },
      {
        path: [1, 0],
        label: 'Widget',
        payload: {
          id: 'tracked-widget',
          type: MDX_BLOCK,
          version: 1,
          attributes: { host: 'react' },
          mdast: widget,
        },
      },
    ];
    const plan = rowPlan(editor, [[code, nested], [widget]], 'below', 1, undefined, { payloads });
    const original = canonicalOoxmlFingerprint(editor.surface!.session.part());
    const result = trackTableRowPlan(editor, plan, CELL_CONTEXT);
    if (!result.success) throw new Error(result.error);
    expect(result.proposals).toHaveLength(1);
    expect(editor.getReviewItems()).toHaveLength(1);
    const pending = serializeOoxmlPart(editor.surface!.session.part());
    expect(pending).toContain('tracked-code');
    expect(pending).toContain('tracked-widget');
    expect(pending).toContain('const value = 1;');
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(original);
  } finally {
    editor.destroy();
  }
});
