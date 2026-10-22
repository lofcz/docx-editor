import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  storyParagraphs,
  writeOoxmlPackage,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import {
  namedChild,
  nodeChildren,
  rowCells,
  tableRows,
} from '../../test/support/native-table-oracle';
import { cell, row, table } from '../../test/support/target-table-fixture';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { applyDirectPlan } from '../ai/apply-direct';
import type { CanonicalEdit } from './plan';
import { planTargetDocument } from './plan-target';

const context = { documentId: 'merges', epoch: 'mount' };
const TAIL = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
const HEADER = '<w:hdr><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>';

function mount(document_: Uint8Array) {
  const container = document.createElement('div');
  document.body.append(container);
  return createDocxEditor({ container, document: document_ });
}
async function source(editor: DocxEditorInstance) {
  const token = readEditorRevision(editor, context.documentId, context.epoch)!,
    view = await readWorkingView(editor.surface!.session, token);
  if (!view.success) throw new Error(view.error);
  return {
    revision: token,
    pkg: editor.surface!.session.currentPackage(),
    projection: view.view.projection,
  };
}
const part = (editor: DocxEditorInstance) => editor.surface!.session.part();
const nativeTable = (value: OoxmlPart) =>
  nodeChildren(bodyStoryRoot(value)!).find((node) => node.kind === 'table')!;
const wordVal = (node: OoxmlNode | undefined) =>
  node && node.kind !== 'textValue'
    ? (node.attributes.find((attribute) => attribute.localName === 'val')?.value ?? 'continue')
    : null;
/** Native visual layout: each physical cell's text, grid span, and vertical merge state. */
function layout(value: OoxmlPart) {
  return tableRows(value, nativeTable(value).id).map((tableRow) =>
    rowCells(tableRow).map((tableCell) => {
      const properties = namedChild(tableCell, 'tcPr');
      return {
        text: storyParagraphs(tableCell).map((node) => paragraphTextOf(value, node.id)),
        span: Number(wordVal(properties && namedChild(properties, 'gridSpan')) ?? 1),
        merge: wordVal(properties && namedChild(properties, 'vMerge')),
      };
    })
  );
}

/** Plan, apply directly, replan empty, undo in one step, redo, and replan empty after reopen. */
async function land(
  editor: DocxEditorInstance,
  target: string,
  kinds: readonly CanonicalEdit['kind'][]
) {
  const planned = planTargetDocument(await source(editor), target);
  if (!planned.success) throw new Error(`${planned.code}: ${planned.error}`);
  expect(planned.plan.operations.map((edit) => edit.kind)).toEqual([...kinds]);
  expect(JSON.stringify(planned.plan)).not.toContain('replaceStoryBlocks');
  const before = canonicalOoxmlFingerprint(nativeTable(part(editor)));
  const applied = applyDirectPlan(editor, planned.plan, context);
  if (!applied.success) throw new Error(`${applied.code}: ${applied.error}`);
  const after = layout(part(editor));
  const replanned = planTargetDocument(await source(editor), target);
  if (!replanned.success) throw new Error(`replan ${replanned.code}: ${replanned.error}`);
  expect(replanned.plan.operations).toEqual([]);
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(editor.surface!.session.canUndo()).toBe(false);
  expect(canonicalOoxmlFingerprint(nativeTable(part(editor)))).toBe(before);
  expect(editor.exec({ type: 'redo' }).ok).toBe(true);
  expect(layout(part(editor))).toEqual(after);
  const saved = new Uint8Array(await editor.save()),
    again = mount(saved);
  try {
    expect(layout(part(again))).toEqual(after);
    const check = planTargetDocument(await source(again), target);
    if (!check.success) throw new Error(`reopen ${check.code}: ${check.error}`);
    expect(check.plan.operations).toEqual([]);
  } finally {
    again.destroy();
  }
  return { plan: planned.plan, after };
}
const html = (rows: readonly (readonly string[])[]) =>
  '<table>\n' +
  rows.map((cells) => '<tr>\n' + cells.join('\n') + '\n</tr>').join('\n') +
  '\n</table>';
const td = (text: string, attributes = '') => `<td${attributes}>\n\n${text}\n\n</td>`;
const merged = (text: string, span: number, extra = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${2400 * span}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}${extra}</w:tcPr><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;
const BLANK =
  '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/></w:pPr></w:p></w:tc>';
const blankRow = (cells: readonly (string | null)[]) =>
  '<w:tr>' + cells.map((text) => (text === null ? BLANK : cell(text))).join('') + '</w:tr>';
const plainTable = (rows: readonly (readonly (string | null)[])[]) =>
  `<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="2400"/>'.repeat(rows[0]!.length)}</w:tblGrid>${rows.map(blankRow).join('')}</w:tbl>`;
const grid = (columns: number, rows: string) =>
  `<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="2400"/>'.repeat(columns)}</w:tblGrid>${rows}</w:tbl>`;

test('a colspan target merges two cells and gives the head its new text', async () => {
  const editor = mount(
    docxFixture(
      table([
        ['A', 'B', 'C'],
        ['One', 'Two', 'Three'],
      ]) + TAIL
    )
  );
  try {
    const cells = rowCells(tableRows(part(editor), nativeTable(part(editor)).id)[0]!);
    const { after } = await land(
      editor,
      html([
        [td('Title', ' colspan="2"'), td('C')],
        [td('One'), td('Two'), td('Three')],
      ]),
      ['text', 'text', 'table']
    );
    expect(after[0]).toEqual([
      { text: ['Title'], span: 2, merge: null },
      { text: ['C'], span: 1, merge: null },
    ]);
    // The head and the untouched neighbor keep their native identities.
    expect(findNode(part(editor), cells[0]!.id)?.kind).toBe('tableCell');
    expect(findNode(part(editor), cells[2]!.id)).toBe(cells[2]);
    expect(findNode(part(editor), cells[1]!.id)).toBeNull();
  } finally {
    editor.destroy();
  }
}, 60_000);

test('a merge into a blank head keeps the covered paragraph intact', async () => {
  const editor = mount(
    docxFixture(
      plainTable([
        ['A', 'B'],
        [null, 'Two'],
      ]) + TAIL
    )
  );
  try {
    const value = part(editor),
      cells = rowCells(tableRows(value, nativeTable(value).id)[1]!),
      paragraph = storyParagraphs(cells[1]!)[0]!;
    const { after } = await land(editor, html([[td('A'), td('B')], [td('Two', ' colspan="2"')]]), [
      'table',
    ]);
    expect(after[1]).toEqual([{ text: ['Two'], span: 2, merge: null }]);
    expect(findNode(part(editor), paragraph.id)).toBe(paragraph);
  } finally {
    editor.destroy();
  }
}, 60_000);

test('an emptied cell keeps one blank native paragraph', async () => {
  const editor = mount(
    docxFixture(
      table([
        ['A', 'B'],
        ['One', 'Two'],
      ]) + TAIL
    )
  );
  try {
    const value = part(editor),
      kept = storyParagraphs(rowCells(tableRows(value, nativeTable(value).id)[1]!)[1]!)[0]!;
    const { after } = await land(editor, '| A | B |\n| --- | --- |\n| One |  |', ['text']);
    expect(after[1]!.map((entry) => entry.text)).toEqual([['One'], ['']]);
    expect(findNode(part(editor), kept.id)?.kind).toBe('paragraph');
  } finally {
    editor.destroy();
  }
}, 60_000);

test('a rowspan target merges a column of cells, dropping blank covered cells', async () => {
  const editor = mount(
    docxFixture(
      plainTable([
        ['Group', 'B1'],
        [null, 'B2'],
        ['A3', 'B3'],
      ]) + TAIL
    )
  );
  try {
    const { after } = await land(
      editor,
      html([[td('Group', ' rowspan="2"'), td('B1')], [td('B2')], [td('A3'), td('B3')]]),
      ['table']
    );
    expect(after.map((cells) => cells.map((entry) => [entry.text, entry.merge]))).toEqual([
      [
        [['Group'], 'restart'],
        [['B1'], null],
      ],
      [
        [[''], 'continue'],
        [['B2'], null],
      ],
      [
        [['A3'], null],
        [['B3'], null],
      ],
    ]);
  } finally {
    editor.destroy();
  }
}, 60_000);

test('content written into merged placeholders splits the native cell into its grid', async () => {
  const editor = mount(
    docxFixture(
      grid(
        3,
        '<w:tr>' +
          merged('AB', 2, '<w:vMerge w:val="restart"/>') +
          merged('C1', 1) +
          '</w:tr><w:tr>' +
          merged('', 2, '<w:vMerge/>') +
          merged('C2', 1) +
          '</w:tr>' +
          row(['A3', 'B3', 'C3'])
      ) + TAIL
    )
  );
  try {
    const value = part(editor),
      head = rowCells(tableRows(value, nativeTable(value).id)[0]!)[0]!;
    const { plan, after } = await land(
      editor,
      '| A | B | C1 |\n| --- | --- | --- |\n| X | Y | C2 |\n| A3 | B3 | C3 |',
      ['text', 'split-cell']
    );
    const split = plan.operations.find((edit) => edit.kind === 'split-cell')!;
    expect(split.kind === 'split-cell' && [split.intent.rows, split.intent.columns]).toEqual([
      2, 2,
    ]);
    expect(after.map((cells) => cells.map((entry) => entry.text[0]))).toEqual([
      ['A', 'B', 'C1'],
      ['X', 'Y', 'C2'],
      ['A3', 'B3', 'C3'],
    ]);
    expect(after.flat().every((entry) => entry.span === 1 && entry.merge === null)).toBe(true);
    // The first resulting cell keeps the native cell.
    expect(findNode(part(editor), head.id)?.kind).toBe('tableCell');
  } finally {
    editor.destroy();
  }
}, 60_000);

test('explicit rowspans split a two-by-two cell into two vertical bands', async () => {
  const editor = mount(
    docxFixture(
      grid(
        3,
        '<w:tr>' +
          merged('AB', 2, '<w:vMerge w:val="restart"/>') +
          merged('C1', 1) +
          '</w:tr><w:tr>' +
          merged('', 2, '<w:vMerge/>') +
          merged('C2', 1) +
          '</w:tr>'
      ) + TAIL
    )
  );
  try {
    const { after } = await land(
      editor,
      html([[td('Left', ' rowspan="2"'), td('Right', ' rowspan="2"'), td('C1')], [td('C2')]]),
      ['text', 'split-cell']
    );
    expect(after.map((cells) => cells.map((entry) => [entry.text[0], entry.merge]))).toEqual([
      [
        ['Left', 'restart'],
        ['Right', 'restart'],
        ['C1', null],
      ],
      [
        ['', 'continue'],
        ['', 'continue'],
        ['C2', null],
      ],
    ]);
  } finally {
    editor.destroy();
  }
}, 60_000);

test('an emptied merge placeholder keeps the native merge', async () => {
  const editor = mount(
    docxFixture(
      grid(3, '<w:tr>' + merged('AB', 2) + merged('C1', 1) + '</w:tr>' + row(['A2', 'B2', 'C2'])) +
        TAIL
    )
  );
  try {
    const { after } = await land(
      editor,
      '| AB changed |  | C1 |\n| --- | --- | --- |\n| A2 | B2 | C2 |',
      ['text']
    );
    expect(after[0]!.map((entry) => [entry.text[0], entry.span])).toEqual([
      ['AB changed', 2],
      ['C1', 1],
    ]);
  } finally {
    editor.destroy();
  }
}, 60_000);

test('crossing or ragged span targets refuse without writes', async () => {
  const editor = mount(
    docxFixture(
      grid(3, '<w:tr>' + merged('AB', 2) + merged('C1', 1) + '</w:tr>' + row(['A2', 'B2', 'C2'])) +
        TAIL
    )
  );
  try {
    const input = await source(editor),
      bytes = writeOoxmlPackage(input.pkg);
    for (const target of [
      // Merges B..C across the native AB merge boundary.
      html([
        [td('A'), td('BC', ' colspan="2"')],
        [td('A2'), td('B2'), td('C2')],
      ]),
      // A span that leaves the table.
      html([
        [td('AB', ' colspan="2"'), td('C1', ' rowspan="3"')],
        [td('A2'), td('B2')],
      ]),
      // Unequal split parts.
      html([
        [td('AB', ' colspan="2"'), td('C1')],
        [td('A2', ' colspan="2"'), td('B2'), td('C2')],
      ]),
    ]) {
      const result = planTargetDocument(input, target);
      expect(result.success).toBe(false);
      expect('plan' in result).toBe(false);
    }
    expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
  } finally {
    editor.destroy();
  }
}, 60_000);

void HEADER;
