import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  directParagraphProperties,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { targetDocumentInput, targetParagraph as p } from '../../test/support/target-document';
import { table, part, nativeTable, plan, content } from '../../test/support/target-table-fixture';
import {
  rowCells,
  tableRows,
  nodeChildren,
  namedChild,
} from '../../test/support/native-table-oracle';
import {
  formatTable,
  formatRow,
  formatCell,
  formatAttr,
} from '../../test/support/table-format-fixture';
import { applyTargetPlanOracle } from '../../test/support/target-plan-oracle';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { sourceRevision } from '../../test/support/target-source-fixture';
import { protectedDocument } from '../../test/support/protected-document';
import { nativeGridColumns } from './inserted-column-native';
import { planTargetDocument } from './plan-target';
import type { TargetSourceInput } from './target-source-types';

const plain = [
  ['A', 'B'],
  ['One', 'Two'],
];
type HtmlOptions = {
  table?: string;
  row?: (index: number) => string;
  cell?: (row: number, column: number) => string;
};
const html = (rows: readonly (readonly string[])[], options: HtmlOptions = {}) =>
  '<table' +
  (options.table ? ' ' + options.table : '') +
  '>\n' +
  rows
    .map(
      (row, index) =>
        '<tr' +
        (options.row?.(index) ? ' ' + options.row(index) : '') +
        '>\n' +
        row
          .map(
            (value, column) =>
              '<td' +
              (options.cell?.(index, column) ? ' ' + options.cell(index, column) : '') +
              '>\n\n' +
              value +
              '\n\n</td>'
          )
          .join('\n') +
        '\n</tr>'
    )
    .join('\n') +
  '\n</table>';
const at = (input: TargetSourceInput, row: number, column: number) =>
  rowCells(tableRows(part(input), nativeTable(input).id)[row]!)[column]!;
const property = (input: TargetSourceInput, row: number, column: number, name: string) =>
  namedChild(namedChild(at(input, row, column), 'tcPr')!, name)!;
const alignments = (input: TargetSourceInput) =>
  tableRows(part(input), nativeTable(input).id).map((row) =>
    rowCells(row).map((cell) =>
      storyParagraphs(cell).map(
        (paragraph) =>
          directParagraphProperties(part(input), paragraph.id).find(
            (value) => value.localName === 'jc'
          )?.attributes?.val
      )
    )
  );
async function apply(input: TargetSourceInput, target: string) {
  const result = plan(input, target),
    pkg = applyTargetPlanOracle(input.pkg, result.plan),
    reopened = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!reopened.ok) throw new Error(reopened.reason);
  return { result, applied: { ...input, pkg }, reopened: { ...input, pkg: reopened.package } };
}
async function imported() {
  const values = [
    ['A1', 'B1', 'C1'],
    ['A2', 'B2', 'C2'],
    ['A3', 'B3', 'C3'],
  ];
  return {
    input: await targetDocumentInput(
      formatTable(
        formatRow(...(values[0] as [string, string, string]), true) +
          formatRow(...(values[1] as [string, string, string])) +
          formatRow(...(values[2] as [string, string, string]))
      )
    ),
    values: values.map((row) => row.map((value) => `**${value}**`)),
  };
}

test('GFM column alignment maps to original paragraph properties without replacing content and repeats as an authored no-op', async () => {
  const input = await targetDocumentInput(table(plain)),
    bytes = writeOoxmlPackage(input.pkg),
    target = '| A | B |\n| :--- | ---: |\n| One | Two |',
    { result, applied, reopened } = await apply(input, target);
  expect(result.plan.operations).toHaveLength(4);
  expect(
    result.plan.operations.every(
      (edit) => edit.kind === 'format' && edit.target.kind === 'paragraph'
    )
  ).toBe(true);
  for (const next of [applied, reopened]) {
    expect(alignments(next)).toEqual([
      [['left'], ['right']],
      [['left'], ['right']],
    ]);
    expect(content(next)).toEqual([
      [['A'], ['B']],
      [['One'], ['Two']],
    ]);
  }
  expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
  const repeat = plan({ ...applied, projection: await projectSourcePackage(applied.pkg) }, target);
  expect(repeat.plan.operations).toEqual([]);
});

test('retained cell fill, vertical alignment, and one border preserve imported property identities and unrelated formatting', async () => {
  const { input, values } = await imported(),
    oldCell = at(input, 1, 1),
    oldShading = property(input, 1, 1, 'shd'),
    oldBorders = property(input, 1, 1, 'tcBorders'),
    target = html(values, {
      cell: (row, column) =>
        row === 1 && column === 1
          ? 'bgcolor="#abc" valign="middle" style="border-top: 1px solid #123456"'
          : '',
    }),
    { result, applied, reopened } = await apply(input, target);
  expect(result.plan.operations.map((edit) => edit.kind)).toEqual(['table', 'table', 'table']);
  for (const next of [applied, reopened]) {
    expect(formatAttr(property(next, 1, 1, 'shd'), 'fill')).toBe('AABBCC');
    expect(formatAttr(property(next, 1, 1, 'shd'), 'color')).toBe('112233');
    expect(formatAttr(property(next, 1, 1, 'shd'), 'val')).toBe('pct10');
    expect(formatAttr(property(next, 1, 1, 'vAlign'), 'val')).toBe('center');
    expect(formatAttr(namedChild(property(next, 1, 1, 'tcBorders'), 'top')!, 'sz')).toBe('6');
    expect(formatAttr(namedChild(property(next, 1, 1, 'tcBorders'), 'top')!, 'color')).toBe(
      '123456'
    );
  }
  expect(property(applied, 1, 1, 'shd').id).toBe(oldShading.id);
  expect(formatAttr(property(applied, 1, 1, 'shd'), 'value', 'urn:keep')).toBe('shading');
  for (const side of ['left', 'bottom', 'right'])
    expect(namedChild(property(applied, 1, 1, 'tcBorders'), side)).toBe(
      namedChild(oldBorders, side)
    );
  for (const paragraph of storyParagraphs(oldCell))
    expect(findNode(part(applied), paragraph.id)).toBe(paragraph);
  expect(at(applied, 0, 0)).toBe(at(input, 0, 0));
  expect(
    plan({ ...applied, projection: await projectSourcePackage(applied.pkg) }, target).plan
      .operations
  ).toEqual([]);
});

test('transparent fills and explicit border clears preserve shading patterns and neighboring sides', async () => {
  const { input, values } = await imported(),
    old = property(input, 1, 1, 'tcBorders'),
    { applied, reopened } = await apply(
      input,
      html(values, {
        cell: (row, column) =>
          row === 1 && column === 1
            ? 'style="background-color: transparent; border-left: none"'
            : '',
      })
    );
  for (const next of [applied, reopened]) {
    expect(formatAttr(property(next, 1, 1, 'shd'), 'fill')).toBeUndefined();
    expect(formatAttr(property(next, 1, 1, 'shd'), 'val')).toBe('pct10');
    expect(formatAttr(namedChild(property(next, 1, 1, 'tcBorders'), 'left')!, 'val')).toBe('none');
  }
  for (const side of ['top', 'bottom', 'right'])
    expect(namedChild(property(applied, 1, 1, 'tcBorders'), side)).toBe(namedChild(old, side));
});

test('table width preserves unequal grid proportions and exact row height retains imported row metadata', async () => {
  const { input, values } = await imported(),
    oldRows = tableRows(part(input), nativeTable(input).id),
    oldHeight = namedChild(namedChild(oldRows[1]!, 'trPr')!, 'trHeight')!,
    target = html(values, {
      table: 'style="width: 600px"',
      row: (row) => (row === 1 ? 'height="24pt"' : ''),
    }),
    { applied, reopened } = await apply(input, target);
  for (const next of [applied, reopened]) {
    expect(
      nativeGridColumns(part(next), nativeTable(next).id).map((column) => formatAttr(column, 'w'))
    ).toEqual(['1500', '3000', '4500']);
    const row = tableRows(part(next), nativeTable(next).id)[1]!,
      height = namedChild(namedChild(row, 'trPr')!, 'trHeight')!;
    expect(formatAttr(height, 'val')).toBe('480');
    expect(formatAttr(height, 'hRule')).toBe('exact');
    expect(formatAttr(height, 'value', 'urn:keep')).toBe('height');
  }
  expect(
    namedChild(
      namedChild(tableRows(part(applied), nativeTable(applied).id)[1]!, 'trPr')!,
      'trHeight'
    )!.id
  ).toBe(oldHeight.id);
  expect(
    plan({ ...applied, projection: await projectSourcePackage(applied.pkg) }, target).plan
      .operations
  ).toEqual([]);
});

test('explicit single-span column constraints share a table width without conflicting native widths', async () => {
  const { input, values } = await imported(),
    { applied, reopened } = await apply(
      input,
      html(values, {
        table: 'width="600"',
        cell: (_, column) => (column === 0 ? 'width="160px"' : ''),
      })
    );
  for (const next of [applied, reopened])
    expect(
      nativeGridColumns(part(next), nativeTable(next).id).map((column) => formatAttr(column, 'w'))
    ).toEqual(['2400', '2640', '3960']);
});

test('content and explicit properties compose over the original cell and paragraph identities', async () => {
  const { input, values } = await imported(),
    oldCell = at(input, 1, 1),
    oldParagraph = storyParagraphs(oldCell)[0]!;
  values[1]![1] = '**Changed B2**';
  const target = html(values, {
      cell: (row, column) =>
        row === 1 && column === 1
          ? 'style="background-color: rgb(12, 34, 56); text-align: center"'
          : '',
    }),
    { result, applied, reopened } = await apply(input, target);
  expect(
    result.plan.operations.some(
      (edit) => edit.kind === 'text' && edit.paragraphId === oldParagraph.id
    )
  ).toBe(true);
  expect(result.plan.operations.some((edit) => edit.kind === 'siblings')).toBe(false);
  expect(findNode(part(applied), oldCell.id)?.id).toBe(oldCell.id);
  expect(paragraphTextOf(part(applied), oldParagraph.id)).toBe('Changed B2');
  for (const next of [applied, reopened]) {
    expect(formatAttr(property(next, 1, 1, 'shd'), 'fill')).toBe('0C2238');
    expect(alignments(next)[1]![1]).toEqual(['center']);
  }
});

test('visible vertical merged cells expand to complete native style selections while preserving hidden content', async () => {
  const body = formatTable(
      '<w:tr>' +
        formatCell('AB', 3600, '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>') +
        formatCell('C1', 3600) +
        '</w:tr>' +
        '<w:tr>' +
        formatCell('', 3600, '<w:gridSpan w:val="2"/><w:vMerge/>') +
        formatCell('C2', 3600) +
        '</w:tr>' +
        formatRow('A3', 'B3', 'C3')
    ),
    input = await targetDocumentInput(body),
    oldRows = tableRows(part(input), nativeTable(input).id),
    selected = [rowCells(oldRows[0]!)[0]!.id, rowCells(oldRows[1]!)[0]!.id],
    values = [
      ['**AB**', '', '**C1**'],
      ['', '', '**C2**'],
      ['**A3**', '**B3**', '**C3**'],
    ],
    { result, applied } = await apply(
      input,
      html(values, {
        cell: (row, column) =>
          row === 0 && column === 0 ? 'style="background-color: #fed; vertical-align: bottom"' : '',
      })
    );
  expect(result.plan.operations).toMatchObject([
    { kind: 'table', operation: { op: 'setTableCellFill', cellIds: selected } },
    { kind: 'table', operation: { op: 'setTableCellVerticalAlignment', cellIds: selected } },
  ]);
  for (const cellId of selected) {
    const cell = findNode(part(applied), cellId)!;
    expect(formatAttr(namedChild(namedChild(cell, 'tcPr')!, 'shd')!, 'fill')).toBe('FFEEDD');
    expect(formatAttr(namedChild(namedChild(cell, 'tcPr')!, 'vAlign')!, 'val')).toBe('bottom');
  }
});

test('cell alignment inheritance removes only direct paragraph alignment', async () => {
  const input = await targetDocumentInput(
      table(plain).replace('<w:pPr>', '<w:pPr><w:jc w:val="center"/>')
    ),
    { applied } = await apply(
      input,
      html(plain, { cell: (row, column) => (row === 0 && column === 0 ? 'align="inherit"' : '') })
    );
  expect(alignments(applied)[0]![0]).toEqual([undefined]);
  expect(paragraphTextOf(part(applied), storyParagraphs(at(input, 0, 0))[0]!.id)).toBe('A');
});

test('explicit table, row, and cell text alignment inherit in order and preserve an explicit parent for inherit', async () => {
  const input = await targetDocumentInput(table(plain)),
    { applied, reopened } = await apply(
      input,
      html(plain, {
        table: 'style="text-align: center"',
        row: (row) => (row === 1 ? 'align="right"' : ''),
        cell: (row, column) => (row === 1 && column === 1 ? 'style="text-align: inherit"' : ''),
      })
    );
  for (const next of [applied, reopened])
    expect(alignments(next)).toEqual([
      [['center'], ['center']],
      [['right'], ['right']],
    ]);
});

test('static numeric attributes and absolute CSS units have explicit native physical sizes', async () => {
  const input = await targetDocumentInput(table(plain)),
    { applied } = await apply(
      input,
      html(plain, {
        table: 'width={400}',
        row: (row) => (row === 1 ? 'height={24}' : ''),
        cell: (row, column) => (row === 1 && column === 1 ? 'bgcolor={"#00ff00"}' : ''),
      })
    );
  expect(
    nativeGridColumns(part(applied), nativeTable(applied).id).map((column) =>
      formatAttr(column, 'w')
    )
  ).toEqual(['3000', '3000']);
  expect(
    formatAttr(
      namedChild(
        namedChild(tableRows(part(applied), nativeTable(applied).id)[1]!, 'trPr')!,
        'trHeight'
      )!,
      'val'
    )
  ).toBe('360');
  expect(formatAttr(property(applied, 1, 1, 'shd'), 'fill')).toBe('00FF00');
  for (const size of ['96px', '72pt', '1in', '2.54cm', '25.4mm']) {
    const next = await apply(
      input,
      html(plain, { row: (row) => (row === 1 ? `style="height: ${size}"` : '') })
    );
    expect(
      formatAttr(
        namedChild(
          namedChild(tableRows(part(next.applied), nativeTable(next.applied).id)[1]!, 'trPr')!,
          'trHeight'
        )!,
        'val'
      )
    ).toBe('1440');
  }
});

test('table outside borders leave internal cell edges and text intact', async () => {
  const { input, values } = await imported(),
    inner = property(input, 1, 1, 'tcBorders'),
    { applied, reopened } = await apply(
      input,
      html(values, { table: 'style="border: 2px solid navy"' })
    );
  for (const side of ['top', 'bottom', 'left', 'right'])
    expect(namedChild(property(applied, 1, 1, 'tcBorders'), side)).toBe(namedChild(inner, side));
  for (const next of [applied, reopened]) {
    const border = namedChild(property(next, 0, 0, 'tcBorders'), 'top')!;
    expect(formatAttr(border, 'color')).toBe('000080');
    expect(formatAttr(border, 'sz')).toBe('12');
  }
  expect(content(applied)).toEqual(content(input));
});

test('outer cell alignment stays outside nested table paragraphs', async () => {
  const nested = table([
      ['Inner A', 'Inner B'],
      ['Inner One', 'Inner Two'],
    ]),
    body =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="4800"/></w:tblGrid><w:tr><w:tc><w:tcPr/>' +
      p('Outer') +
      '</w:tc><w:tc><w:tcPr/>' +
      nested +
      p('Outer tail') +
      '</w:tc></w:tr></w:tbl>',
    input = await targetDocumentInput(body),
    raw = input.projection.text.slice(
      input.projection.text.indexOf('<table>'),
      input.projection.text.indexOf('</table>') + '</table>'.length
    ),
    inner = nodeChildren(at(input, 0, 1)).find((node) => node.kind === 'table')!,
    paragraphs = storyParagraphs(inner),
    { applied } = await apply(
      input,
      html([['Outer', raw + '\n\nOuter tail']], {
        cell: (_, column) => (column === 1 ? 'align="right"' : ''),
      })
    );
  for (const paragraph of paragraphs) expect(findNode(part(applied), paragraph.id)).toBe(paragraph);
  const outer = nodeChildren(at(applied, 0, 1)).find((node) => node.kind === 'paragraph')!;
  expect(
    directParagraphProperties(part(applied), outer.id).find((value) => value.localName === 'jc')
      ?.attributes?.val
  ).toBe('right');
});

test('native property application preserves imported non-body parts, background, and section settings', async () => {
  const entries = unzipSync(protectedDocument()),
    document = strFromU8(entries['word/document.xml']!),
    start = document.indexOf('<w:body>') + '<w:body>'.length,
    section = document.lastIndexOf('<w:sectPr>');
  entries['word/document.xml'] = strToU8(
    document.slice(0, start) + table(plain) + document.slice(section)
  );
  const loaded = readOoxmlPackage(zipSync(entries));
  if (!loaded.ok) throw new Error(loaded.reason);
  const input = {
      pkg: loaded.package,
      revision: sourceRevision,
      projection: await projectSourcePackage(loaded.package),
    },
    before = unzipSync(writeOoxmlPackage(input.pkg)),
    oldSection = nodeChildren(bodyStoryRoot(part(input))!).at(-1)!,
    { applied, reopened } = await apply(
      input,
      html(plain, {
        table: 'width="600px"',
        cell: (row, column) =>
          row === 1 && column === 1 ? 'style="background-color: #ffc; text-align: center"' : '',
      })
    ),
    after = unzipSync(writeOoxmlPackage(applied.pkg));
  expect(Object.keys(after).sort()).toEqual(Object.keys(before).sort());
  for (const [name, bytes] of Object.entries(before))
    if (name !== 'word/document.xml') expect(after[name]).toEqual(bytes);
  expect(findNode(part(applied), oldSection.id)).toBe(oldSection);
  expect(canonicalOoxmlFingerprint(nodeChildren(bodyStoryRoot(part(reopened))!).at(-1)!)).toBe(
    canonicalOoxmlFingerprint(oldSection)
  );
});

test('generated cell separators stay outside native edits while authored hard breaks remain editable', async () => {
  const nested = table([
      ['Inner A', 'Inner B'],
      ['Inner One', 'Inner Two'],
    ]),
    body =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="4800"/></w:tblGrid><w:tr><w:tc><w:tcPr/>' +
      p('Outer') +
      '</w:tc><w:tc><w:tcPr/>' +
      nested +
      '<w:p><w:r><w:t>Outer</w:t><w:br/><w:t xml:space="preserve"> tail</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
    input = await targetDocumentInput(body),
    bytes = writeOoxmlPackage(input.pkg),
    raw = input.projection.text.slice(
      input.projection.text.indexOf('<table>'),
      input.projection.text.indexOf('</table>') + '</table>'.length
    ),
    paragraph = nodeChildren(at(input, 0, 1)).find((node) => node.kind === 'paragraph')!,
    keep = await apply(
      input,
      html([['Outer', raw + '\n\nOuter<br/> tail']], {
        cell: (_, column) => (column === 1 ? 'align="right"' : ''),
      })
    );
  expect(keep.result.plan.operations.some((edit) => edit.kind === 'text')).toBe(false);
  expect(paragraphTextOf(part(keep.applied), paragraph.id)).toBe('Outer\n tail');
  const remove = await apply(
    input,
    html([['Outer', raw + '\n\nOuter tail']], {
      cell: (_, column) => (column === 1 ? 'align="right"' : ''),
    })
  );
  expect(remove.result.plan.operations.filter((edit) => edit.kind === 'text')).toEqual([
    {
      kind: 'text',
      partName: part(input).name,
      paragraphId: paragraph.id,
      start: 5,
      end: 6,
      before: '\n',
      after: '',
      bias: 'right',
    },
  ]);
  for (const next of [remove.applied, remove.reopened]) {
    const tail = nodeChildren(at(next, 0, 1)).find((node) => node.kind === 'paragraph')!;
    expect(paragraphTextOf(part(next), tail.id)).toBe('Outer tail');
  }
  expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
});

test('incompatible GFM alignments for one merged owner refuse the entire plan', async () => {
  const input = await targetDocumentInput(
      formatTable(
        '<w:tr>' +
          formatCell('AB', 3600, '<w:gridSpan w:val="2"/>') +
          formatCell('C', 3600) +
          '</w:tr>'
      )
    ),
    bytes = writeOoxmlPackage(input.pkg),
    target = input.projection.text.replace('| --- | --- | --- |', '| :---: | ---: | --- |'),
    result = planTargetDocument(input, target);
  expect(result.success).toBe(false);
  expect('plan' in result).toBe(false);
  if (!result.success) expect(result.error).toContain('conflicting');
  expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
});

test('complete explicit widths can author native grids without readable original width values', async () => {
  const input = await targetDocumentInput(
      table(plain).replaceAll('<w:gridCol w:w="2400"/>', '<w:gridCol/>')
    ),
    { applied, reopened } = await apply(
      input,
      html(plain, {
        table: 'width="400"',
        cell: (_, column) => (column === 0 ? 'width="100"' : 'width="300"'),
      })
    );
  for (const next of [applied, reopened])
    expect(
      nativeGridColumns(part(next), nativeTable(next).id).map((column) => formatAttr(column, 'w'))
    ).toEqual(['1500', '4500']);
});

test('row deletion that promotes a hidden merged owner refuses until its content has a native rebind', async () => {
  const input = await targetDocumentInput(
      formatTable(
        '<w:tr>' +
          formatCell('AB', 3600, '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>') +
          formatCell('C1', 3600) +
          '</w:tr><w:tr>' +
          formatCell('', 3600, '<w:gridSpan w:val="2"/><w:vMerge/>') +
          formatCell('C2', 3600) +
          '</w:tr>' +
          formatRow('A3', 'B3', 'C3')
      )
    ),
    bytes = writeOoxmlPackage(input.pkg),
    result = planTargetDocument(
      input,
      html(
        [
          ['', '', '**C2**'],
          ['**A3**', '**B3**', '**C3**'],
        ],
        { table: 'width="500px"' }
      )
    );
  expect(result.success).toBe(false);
  expect('plan' in result).toBe(false);
  if (!result.success) expect(result.error).toContain('merge');
  expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
});

test('unknown declarations, unsupported units, dynamic expressions, conflicts, and new property destinations refuse without partial writes', async () => {
  const input = await targetDocumentInput(table(plain)),
    bytes = writeOoxmlPackage(input.pkg),
    fingerprint = canonicalOoxmlFingerprint(part(input));
  const targets = [
    html(plain, { table: 'width="80%"' }),
    html(plain, { table: 'style="width: 400"' }),
    html(plain, { table: 'width={(() => { throw new Error("executed") })()}' }),
    html(plain, { cell: () => 'style="background-color: rgba(0,0,0,0.5)"' }),
    html(plain, { cell: () => 'style="padding: 4px"' }),
    html(
      [
        ['A', 'B'],
        ['Changed One', 'Two'],
      ],
      { row: (row) => (row === 1 ? 'height="0"' : '') }
    ),
    html(plain, { cell: () => 'style="border-top: 1px solid toString"' }),
    html(plain, { cell: () => 'style="background-color: red; background-color: blue"' }),
    html(plain, {
      cell: (row, column) => (column === 0 ? `width="${row === 0 ? 100 : 200}px"` : ''),
    }),
    '| A | X | B |\n| :--- | :--- | ---: |\n| One | New | Two |',
  ];
  for (const target of targets) {
    const result = planTargetDocument(input, target);
    expect(result.success).toBe(false);
    expect('plan' in result).toBe(false);
  }
  expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
  expect(canonicalOoxmlFingerprint(part(input))).toBe(fingerprint);
});
