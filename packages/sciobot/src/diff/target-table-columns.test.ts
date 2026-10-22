import { propertyEditor, propertyContext } from '../../test/support/property-execution';
import { groupToken } from '../../test/support/review-groups';
import { applyDirectPlan } from '../ai/apply-direct';
import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  table,
  part,
  nativeTable,
  content,
  plan,
  landed,
} from '../../test/support/target-table-fixture';
import { targetDocumentInput } from '../../test/support/target-document';
import {
  namedChild,
  nodeChildren,
  tableRows,
  rowCells,
} from '../../test/support/native-table-oracle';
import { nativeGridColumns } from './inserted-column-native';
import { planTargetDocument } from './plan-target';
import { applyTargetPlanOracle } from '../../test/support/target-plan-oracle';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { sourceRevision } from '../../test/support/target-source-fixture';
import { protectedDocument } from '../../test/support/protected-document';
import {
  formatTable,
  formatCell,
  formatRow,
  formatAttr,
} from '../../test/support/table-format-fixture';
import { decodeEditPlan, serializeEditPlan } from './plan';

const markdown = (rows: readonly (readonly string[])[]) =>
  [
    '| ' + rows[0]!.join(' | ') + ' |',
    '| ' + rows[0]!.map(() => '---').join(' | ') + ' |',
    ...rows.slice(1).map((row) => '| ' + row.join(' | ') + ' |'),
  ].join('\n');
const beforeRows = [
  ['A', 'B', 'C'],
  ['One', 'Two', 'Three'],
  ['Tail A', 'Tail B', 'Tail C'],
];

for (const at of [0, 1, 3])
  test(`two inserted columns at slot ${at} preserve target order, original anchors, widths, and native cells`, async () => {
    const input = await targetDocumentInput(table(beforeRows)),
      old = part(input),
      oldTable = nativeTable(input),
      rows = tableRows(old, oldTable.id),
      cells = rows.flatMap(rowCells),
      grid = nativeGridColumns(old, oldTable.id),
      bytes = writeOoxmlPackage(input.pkg),
      target = beforeRows.map((values, row) => {
        const next = [...values];
        next.splice(at, 0, `X${row}`, `Y${row}`);
        return next;
      }),
      { result, input: applied } = await landed(input, markdown(target));
    expect(result.plan.operations).toHaveLength(2);
    for (const edit of result.plan.operations) {
      if (edit.kind !== 'insert-column') throw new Error('Expected column intent');
      expect(grid.some((column) => column.id === edit.intent.anchor.id)).toBe(true);
      expect(edit.intent.cells.map((cell) => cell.rowId)).toEqual(rows.map((row) => row.id));
      expect(Object.hasOwn(edit.intent.cells[0]!, 'cellId')).toBe(false);
    }
    expect(content(applied)).toEqual(target.map((row) => row.map((value) => [value])));
    for (const node of [...cells, ...grid]) expect(findNode(part(applied), node.id)).toBe(node);
    expect(namedChild(nativeTable(applied), 'tblPr')).toBe(namedChild(oldTable, 'tblPr'));
    expect(nativeGridColumns(part(applied), oldTable.id)).toHaveLength(5);
    expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
    expect(decodeEditPlan(serializeEditPlan(result.plan))).toEqual({
      success: true,
      plan: result.plan,
    });
  });

for (const at of [0, 1, 2])
  test(`deleted column ${at} removes only its original native grid and cells`, async () => {
    const input = await targetDocumentInput(table(beforeRows)),
      old = part(input),
      id = nativeTable(input).id,
      cells = tableRows(old, id).map(rowCells),
      grid = nativeGridColumns(old, id),
      target = beforeRows.map((row) => row.filter((_, index) => index !== at)),
      { result, input: applied } = await landed(input, markdown(target));
    expect(result.plan.operations).toEqual([
      {
        kind: 'table',
        partName: old.name,
        operation: { op: 'deleteTableColumn', tableId: id, gridColumnId: grid[at]!.id },
      },
    ]);
    expect(findNode(part(applied), grid[at]!.id)).toBeNull();
    for (const row of cells)
      for (let column = 0; column < row.length; column++)
        if (column === at) expect(findNode(part(applied), row[column]!.id)).toBeNull();
        else expect(findNode(part(applied), row[column]!.id)).toBe(row[column]);
  });

test('separate column gaps preserve original IDs during simultaneous insertion and deletion', async () => {
  const input = await targetDocumentInput(
      table([
        ['A', 'B', 'C', 'D'],
        ['One', 'Two', 'Three', 'Four'],
      ])
    ),
    old = part(input),
    cells = tableRows(old, nativeTable(input).id).map(rowCells),
    { result, input: applied } = await landed(
      input,
      '| B | X | C |\n| --- | --- | --- |\n| Two | New | Changed Three |'
    );
  expect(result.plan.operations.filter((op) => op.kind === 'insert-column')).toHaveLength(1);
  expect(result.plan.operations.filter((op) => op.kind === 'table')).toHaveLength(2);
  expect(findNode(part(applied), cells[0]![1]!.id)).toBe(cells[0]![1]);
  expect(findNode(part(applied), cells[0]![2]!.id)).toBe(cells[0]![2]);
  expect(paragraphTextOf(part(applied), storyParagraphs(cells[1]![2]!)[0]!.id)).toBe(
    'Changed Three'
  );
});

test('combined row and column edits use original row anchors and fill only retained rows before row insertion', async () => {
  const input = await targetDocumentInput(
      table([
        ['A', 'B'],
        ['Drop', 'Remove'],
        ['Keep', 'Row'],
        ['Tail', 'End'],
      ])
    ),
    oldRows = tableRows(part(input), nativeTable(input).id),
    target = [
      ['A', 'X', 'B'],
      ['Keep', 'new K', 'Row'],
      ['Added', 'new A', 'Data'],
      ['Tail', 'new T', 'End'],
    ],
    { result, input: applied } = await landed(input, markdown(target));
  expect(result.plan.operations.map((op) => op.kind)).toEqual([
    'insert-column',
    'insert-row',
    'table',
  ]);
  const column = result.plan.operations[0]!,
    row = result.plan.operations[1]!;
  if (column.kind !== 'insert-column' || row.kind !== 'insert-row')
    throw new Error('Expected deferred structural intents');
  expect(column.intent.cells.map((cell) => cell.rowId)).toEqual(oldRows.map((row) => row.id));
  expect(column.intent.cells[1]!.blocks).toEqual([]);
  expect(row.intent.rowId).toBe(oldRows[3]!.id);
  expect(row.intent.cells).toHaveLength(3);
  expect(content(applied)).toEqual(target.map((row) => row.map((text) => [text])));
  expect(findNode(part(applied), oldRows[1]!.id)).toBeNull();
});

test('duplicate headers use complete stable column vectors while indistinguishable values retain positional owners', async () => {
  const input = await targetDocumentInput(
      table([
        ['Same', 'Same'],
        ['Left', 'Right'],
      ])
    ),
    cells = tableRows(part(input), nativeTable(input).id).map(rowCells),
    { result, input: applied } = await landed(
      input,
      '| Same | Same | Same |\n| --- | --- | --- |\n| Left | Middle | Right |'
    );
  expect(result.plan.operations).toHaveLength(1);
  for (const cell of cells.flat()) expect(findNode(part(applied), cell.id)).toBe(cell);
  const repeated = await targetDocumentInput(
      table([
        ['Same', 'Same'],
        ['Same', 'Same'],
      ])
    ),
    original = tableRows(part(repeated), nativeTable(repeated).id).map(rowCells),
    next = await landed(
      repeated,
      '| Same | Same | Same |\n| --- | --- | --- |\n| Same | Same | Same |'
    );
  for (const cell of original.flat()) expect(findNode(part(next.input), cell.id)).toBe(cell);
});

test('new column marks and hyperlinks are explicit fragments and source resources remain unchanged', async () => {
  const input = await targetDocumentInput(
      table([
        ['A', 'B'],
        ['One', 'Two'],
      ])
    ),
    bytes = writeOoxmlPackage(input.pkg),
    { result, input: applied } = await landed(
      input,
      '| A | X | B |\n| --- | --- | --- |\n| One | **Bold** [Link](https://example.org/new) | Two |'
    );
  expect(result.plan.operations[0]).toMatchObject({
    kind: 'insert-column',
    intent: {
      cells: [
        { blocks: [{ children: [{ type: 'text', value: 'X' }] }] },
        {
          blocks: [
            {
              children: [
                { type: 'strong' },
                { type: 'text' },
                { type: 'link', url: 'https://example.org/new' },
              ],
            },
          ],
        },
      ],
    },
  });
  expect(input.pkg.externalTargets).toHaveLength(0);
  expect(
    applied.pkg.externalTargets.some((link) => link.rawTarget === 'https://example.org/new')
  ).toBe(true);
  expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
});

test('duplicate headers and concurrent row-shape changes use unique body evidence to retain native columns', async () => {
  const input = await targetDocumentInput(
      table([
        ['Same', 'Same'],
        ['Left', 'Right'],
        ['Last L', 'Last R'],
      ])
    ),
    original = tableRows(part(input), nativeTable(input).id).flatMap(rowCells),
    { result, input: applied } = await landed(
      input,
      markdown([
        ['Same', 'Same', 'Same'],
        ['Left', 'Middle', 'Right'],
        ['New L', 'New M', 'New R'],
        ['Last L', 'Last M', 'Last R'],
      ])
    );
  expect(result.plan.operations.map((edit) => edit.kind)).toEqual(['insert-column', 'insert-row']);
  for (const cell of original) expect(findNode(part(applied), cell.id)).toBe(cell);
});

test('multiple column deletions and changed retained text preserve surviving nonuniform grid widths and foreign metadata', async () => {
  const rows = [
      ['A', 'B', 'C', 'D', 'E'],
      ['One', 'Two', 'Three', 'Four', 'Five'],
    ],
    widths = [1200, 1800, 2400, 3000, 3600];
  let index = 0;
  const xml = table(rows).replaceAll(
      '<w:gridCol w:w="2400"/>',
      () =>
        `<w:gridCol xmlns:keep="urn:keep" w:w="${widths[index]}" keep:column="${index++}"><keep:child/></w:gridCol>`
    ),
    input = await targetDocumentInput(xml),
    grid = nativeGridColumns(part(input), nativeTable(input).id),
    { input: applied } = await landed(input, '| A | E |\n| --- | --- |\n| One | Changed Five |');
  expect(nativeGridColumns(part(applied), nativeTable(input).id)).toEqual([grid[0]!, grid[4]!]);
  expect(nativeGridColumns(part(applied), nativeTable(input).id)[0]).toBe(grid[0]);
  expect(nativeGridColumns(part(applied), nativeTable(input).id)[1]).toBe(grid[4]);
});

test('a table without complete native grid anchors refuses structural target edits without a partial plan or source writes', async () => {
  const input = await targetDocumentInput(
      table([
        ['A', 'B'],
        ['One', 'Two'],
      ]).replace(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/, '')
    ),
    bytes = writeOoxmlPackage(input.pkg),
    fingerprint = canonicalOoxmlFingerprint(part(input)),
    result = planTargetDocument(input, '| A | X | B |\n| --- | --- | --- |\n| One | New | Two |');
  expect(result.success).toBe(false);
  expect('plan' in result).toBe(false);
  expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
  expect(canonicalOoxmlFingerprint(part(input))).toBe(fingerprint);
});

test('a column inside an existing merged span widens the span without inventing content destinations', async () => {
  const input = await targetDocumentInput(
      formatTable(
        formatRow('A', 'B', 'C') +
          '<w:tr>' +
          formatCell('AB', 3600, '<w:gridSpan w:val="2"/>') +
          formatCell('C2', 3600) +
          '</w:tr>' +
          formatRow('A3', 'B3', 'C3')
      )
    ),
    oldRows = tableRows(part(input), nativeTable(input).id),
    merged = rowCells(oldRows[1]!)[0]!,
    paragraph = storyParagraphs(merged)[0]!,
    target =
      '| A | X | B | C |\n| --- | --- | --- | --- |\n| **AB** |  |  | **C2** |\n| **A3** | **New** | **B3** | **C3** |',
    { result, input: applied } = await landed(input, target);
  const edit = result.plan.operations.find((op) => op.kind === 'insert-column')!;
  if (edit.kind !== 'insert-column') throw new Error('Expected inserted column');
  expect(edit.intent.cells[1]!.blocks).toEqual([]);
  expect(findNode(part(applied), paragraph.id)).toBe(paragraph);
  expect(
    formatAttr(
      namedChild(namedChild(findNode(part(applied), merged.id)!, 'tcPr')!, 'gridSpan')!,
      'val'
    )
  ).toBe('3');
  expect(rowCells(tableRows(part(applied), nativeTable(input).id)[1]!)).toHaveLength(2);
  // The formerly refused covered-slot target now has an explicit native split destination.
  const editor = propertyEditor(input.pkg);
  try {
    const mountedParagraph = findNode(editor.surface!.session.part(), paragraph.id)!;
    const token = groupToken(editor),
      split = planTargetDocument(
        {
          ...input,
          pkg: editor.surface!.session.currentPackage(),
          revision: token,
          projection: { ...input.projection, revision: token },
        },
        target.replace('| **AB** |  |  |', '| **AB** | Injected |  |')
      );
    if (!split.success) throw new Error(split.error);
    expect(split.plan.operations).toContainEqual(
      expect.objectContaining({
        kind: 'table-geometry',
        intent: expect.objectContaining({ kind: 'split' }),
      })
    );
    const result = applyDirectPlan(editor, split.plan, propertyContext);
    if (!result.success) throw new Error(result.error);
    const current = editor.surface!.session.part(),
      row = rowCells(tableRows(current, nativeTable(input).id)[1]!);
    expect(row).toHaveLength(4);
    expect(
      row.map((cell) => storyParagraphs(cell).map((p) => paragraphTextOf(current, p.id)))
    ).toEqual([['AB'], ['Injected'], [''], ['C2']]);
    expect(findNode(current, paragraph.id)).toBe(mountedParagraph);
  } finally {
    editor.destroy();
  }
  const restored = await landed(
    { ...applied, projection: await projectSourcePackage(applied.pkg) },
    input.projection.text
  );
  expect(
    formatAttr(
      namedChild(namedChild(findNode(part(restored.input), merged.id)!, 'tcPr')!, 'gridSpan')!,
      'val'
    )
  ).toBe('2');
});

test('new column nested tables preserve complete target cell fragments and native grids through reopen', async () => {
  const input = await targetDocumentInput(
      table([
        ['A', 'B'],
        ['One', 'Two'],
      ])
    ),
    target =
      '<table>\n' +
      [
        ['A', 'X', 'B'],
        ['One', '| H | J |\n| --- | --- |\n| New H | New J |', 'Two'],
      ]
        .map(
          (values) =>
            '<tr>\n' +
            values.map((value) => '<td>\n\n' + value + '\n\n</td>').join('\n') +
            '\n</tr>'
        )
        .join('\n') +
      '\n</table>',
    result = plan(input, target),
    edit = result.plan.operations[0]!;
  expect(result.plan.operations).toHaveLength(1);
  if (edit.kind !== 'insert-column') throw new Error('Expected column fragments');
  expect(edit.intent.cells[1]!.blocks[0]).toMatchObject({
    type: 'table',
    children: [
      {
        children: [
          { type: 'tableCell', children: [{ type: 'text', value: 'H' }] },
          { type: 'tableCell', children: [{ type: 'text', value: 'J' }] },
        ],
      },
      { type: 'tableRow' },
    ],
  });
  const applied = applyTargetPlanOracle(input.pkg, result.plan),
    reopened = readOoxmlPackage(writeOoxmlPackage(applied));
  if (!reopened.ok) throw new Error(reopened.reason);
  for (const pkg of [applied, reopened.package]) {
    const main = pkg.parts.get(input.pkg.mainDocumentPart)!,
      outer = nodeChildren(bodyStoryRoot(main)!).find((node) => node.kind === 'table')!,
      cell = rowCells(tableRows(main, outer.id)[1]!)[1]!,
      nested = nodeChildren(cell).find((node) => node.kind === 'table')!;
    expect(nativeGridColumns(main, nested.id)).toHaveLength(2);
    expect(
      tableRows(main, nested.id).map((row) =>
        rowCells(row).map((cell) =>
          storyParagraphs(cell).map((node) => paragraphTextOf(main, node.id))
        )
      )
    ).toEqual([
      [['H'], ['J']],
      [['New H'], ['New J']],
    ]);
  }
});

test('native column changes retain imported non-body parts, background, and section settings through save/reopen', async () => {
  const entries = unzipSync(protectedDocument()),
    document = strFromU8(entries['word/document.xml']!),
    start = document.indexOf('<w:body>') + '<w:body>'.length,
    section = document.lastIndexOf('<w:sectPr>');
  entries['word/document.xml'] = strToU8(
    document.slice(0, start) +
      table([
        ['A', 'B'],
        ['One', 'Two'],
      ]) +
      document.slice(section)
  );
  const loaded = readOoxmlPackage(zipSync(entries));
  if (!loaded.ok) throw new Error(loaded.reason);
  const input = {
      pkg: loaded.package,
      revision: sourceRevision,
      projection: await projectSourcePackage(loaded.package),
    },
    oldSection = nodeChildren(bodyStoryRoot(part(input))!).at(-1)!,
    background = part(input).root.children.find(
      (node) => node.kind !== 'textValue' && node.localName === 'background'
    )!,
    before = unzipSync(writeOoxmlPackage(input.pkg)),
    { input: applied } = await landed(
      input,
      '| A | X | B |\n| --- | --- | --- |\n| One | New | Two |'
    ),
    saved = writeOoxmlPackage(applied.pkg),
    after = unzipSync(saved);
  expect(Object.keys(after).sort()).toEqual(Object.keys(before).sort());
  for (const [name, bytes] of Object.entries(before))
    if (name !== 'word/document.xml') expect(after[name]).toEqual(bytes);
  expect(findNode(part(applied), oldSection.id)).toBe(oldSection);
  expect(findNode(part(applied), background.id)).toBe(background);
  const reopened = readOoxmlPackage(saved);
  if (!reopened.ok) throw new Error(reopened.reason);
  expect(
    canonicalOoxmlFingerprint(
      nodeChildren(bodyStoryRoot(reopened.package.parts.get(input.pkg.mainDocumentPart)!)!).at(-1)!
    )
  ).toBe(canonicalOoxmlFingerprint(oldSection));
});

test('inserted cell hyperlinks and paragraph attributes cannot bypass authorability or disappear from plans', async () => {
  const input = await targetDocumentInput(
      table([
        ['A', 'B'],
        ['One', 'Two'],
      ])
    ),
    bytes = writeOoxmlPackage(input.pkg),
    targets = [
      '| A | X | B |\n| --- | --- | --- |\n| One | [Link](javascript:alert%281%29) | Changed Two |',
      '<table>\n' +
        [
          ['A', 'X', 'B'],
          ['One', '<p align="right">New</p>', 'Changed Two'],
        ]
          .map(
            (row) =>
              '<tr>\n' + row.map((value) => '<td>\n\n' + value + '\n\n</td>').join('\n') + '\n</tr>'
          )
          .join('\n') +
        '\n</table>',
    ];
  for (const target of targets) {
    const result = planTargetDocument(input, target);
    expect(result.success).toBe(false);
    expect('plan' in result).toBe(false);
  }
  expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
});
