import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  paragraphTextOf,
  withPart,
  directParagraphProperties,
  readOoxmlPackage,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { table, nativeTable, part } from '../../test/support/target-table-fixture';
import { rowCells, tableRows, namedChild } from '../../test/support/native-table-oracle';
import { formatAttr } from '../../test/support/table-format-fixture';
import { planTargetDocument } from './plan-target';
import { previewDirectPlan } from '../ai/apply-direct';
import { stageTrackedPlan } from '../review/tracked-stage';
import { isNewCellPropertyBindings } from './new-cell-properties';
import { createEditPlan, decodeEditPlan, serializeEditPlan } from './plan';

const style =
  'style="background-color: #123456; border: 1pt solid red; vertical-align: bottom; text-align: center"';
const html = (rows: readonly (readonly string[])[], attrs: (r: number, c: number) => string) =>
  '<table>\n' +
  rows
    .map(
      (row, r) =>
        '<tr>\n' +
        row.map((text, c) => `<td ${attrs(r, c)}>\n\n${text}\n\n</td>`).join('\n') +
        '\n</tr>'
    )
    .join('\n') +
  '\n</table>';
const splitXml =
  '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
for (const mode of ['direct', 'tracked'] as const)
  for (const kind of ['row', 'column', 'split'] as const)
    test(`${mode} deferred ${kind} properties survive native allocation and reopen`, async () => {
      const initial = kind === 'row' ? [['A', 'B']] : [['A']];
      const input = await targetDocumentInput(kind === 'split' ? splitXml : table(initial));
      const target = html(kind === 'row' ? [...initial, ['C', 'D']] : [['A', 'B']], (r, c) =>
        (kind === 'row' ? r === 1 : c === 1) ? style : ''
      );
      const bytes = writeOoxmlPackage(input.pkg);
      const planned = planTargetDocument(input, target);
      if (!planned.success) throw new Error(planned.error);
      const operation = planned.plan.operations.find((edit) =>
        ['insert-row', 'insert-column', 'split-cell'].includes(edit.kind)
      );
      if (!operation || !('intent' in operation) || !('cellProperties' in operation.intent))
        throw new Error('Missing deferred property bindings');
      expect(operation.intent.cellProperties).toEqual(
        (kind === 'row' ? [0, 1] : [0]).map((cellIndex) => ({
          cellIndex,
          properties: {
            fill: { kind: 'hex', value: '123456' },
            vertical: 'bottom',
            horizontal: 'center',
            borders: {
              outside: {
                clear: false,
                spec: { style: 'single', size: 8, color: { kind: 'hex', value: 'FF0000' } },
              },
            },
          },
        }))
      );
      expect(decodeEditPlan(serializeEditPlan(planned.plan)).success).toBe(true);
      expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
      let pkg = input.pkg;
      if (mode === 'direct') {
        const result = previewDirectPlan(pkg, planned.plan);
        if (!result.success) throw new Error(result.error);
        pkg = result.pkg;
      } else {
        stageTrackedPlan(
          {
            currentPackage: () => pkg,
            apply: (name, op) => {
              const result = applyTreeOp(pkg.parts.get(name)!, op);
              if (!result.ok) throw new Error(result.reason);
              pkg = withPart(pkg, result.part);
            },
            edit: (change) => {
              pkg = change(pkg);
            },
          },
          planned.plan,
          { author: 'ScioBot', date: '2026-10-22T08:00:00Z' },
          true
        );
      }
      const saved = readOoxmlPackage(writeOoxmlPackage(pkg));
      if (!saved.ok) throw new Error(saved.reason);
      for (const current of [pkg, saved.package]) {
        const value = { ...input, pkg: current },
          main = part(value),
          rows = tableRows(main, nativeTable(value).id);
        const cells = kind === 'row' ? rowCells(rows[1]!) : [rowCells(rows[0]!)[1]!];
        for (const cell of cells) {
          expect(storyParagraphs(cell).map((p) => paragraphTextOf(main, p.id))).toEqual([
            kind === 'row' ? (cell.id === cells[0]!.id ? 'C' : 'D') : 'B',
          ]);
          const props = namedChild(cell, 'tcPr')!;
          expect(formatAttr(namedChild(props, 'shd')!, 'fill')).toBe('123456');
          expect(formatAttr(namedChild(props, 'vAlign')!, 'val')).toBe('bottom');
          for (const side of ['top', 'bottom', 'left', 'right']) {
            const edge = namedChild(namedChild(props, 'tcBorders')!, side)!;
            expect(formatAttr(edge, 'val')).toBe('single');
            expect(formatAttr(edge, 'sz')).toBe('8');
            expect(formatAttr(edge, 'color')).toBe('FF0000');
          }
          expect(
            storyParagraphs(cell).map(
              (p) =>
                directParagraphProperties(main, p.id).find((value) => value.localName === 'jc')
                  ?.attributes?.val
            )
          ).toEqual(['center']);
        }
      }
    });

for (const kind of ['row', 'column'] as const)
  test(`GFM ${kind} inherits declared column alignment after allocation`, async () => {
    const input = await targetDocumentInput(table(kind === 'row' ? [['A', 'B']] : [['A'], ['C']]));
    const target =
      kind === 'row'
        ? '| A | B |\n| :--- | ---: |\n| C | D |'
        : '| A | B |\n| :--- | ---: |\n| C | D |';
    const planned = planTargetDocument(input, target);
    if (!planned.success) throw new Error(planned.error);
    const applied = previewDirectPlan(input.pkg, planned.plan);
    if (!applied.success) throw new Error(applied.error);
    const value = { ...input, pkg: applied.pkg };
    expect(
      tableRows(part(value), nativeTable(value).id).map((row) =>
        rowCells(row).map((cell) =>
          storyParagraphs(cell).map(
            (p) =>
              directParagraphProperties(part(value), p.id).find((v) => v.localName === 'jc')
                ?.attributes?.val
          )
        )
      )
    ).toEqual([
      [['left'], ['right']],
      [['left'], ['right']],
    ]);
  });

test('new property codecs reject malformed, duplicate, foreign, and invalid values', () => {
  const cells = [[{ type: 'paragraph' }]];
  expect(
    isNewCellPropertyBindings(cells, [
      { cellIndex: 0, properties: { fill: null, horizontal: null } },
    ])
  ).toBe(true);
  for (const bindings of [
    null,
    {},
    [{ cellIndex: 1, properties: {} }],
    [{ cellIndex: 0, properties: { width: 10 } }],
    [{ cellIndex: 0, properties: { vertical: 'middle' } }],
    [{ cellIndex: 0, properties: { fill: { kind: 'hex', value: 'oops' } } }],
    [
      { cellIndex: 0, properties: {} },
      { cellIndex: 0, properties: {} },
    ],
  ])
    expect(isNewCellPropertyBindings(cells, bindings)).toBe(false);
  const invalid = createEditPlan({ documentId: 'x', epoch: '1', revision: 1 }, [
    {
      kind: 'split-cell',
      partName: '/word/document.xml',
      intent: {
        tableId: 't',
        cellId: 'c',
        rows: 1,
        columns: 2,
        cells,
        payloads: [],
        cellProperties: [{ cellIndex: 9, properties: {} }],
      },
    },
  ]);
  expect(invalid.success).toBe(false);
});
