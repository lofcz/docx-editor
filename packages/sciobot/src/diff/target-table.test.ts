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
import pairs from '../__tests__/fixtures/ai-pairs.json';
import { targetDocumentInput, targetParagraph as p } from '../../test/support/target-document';
import { applyTargetPlanOracle } from '../../test/support/target-plan-oracle';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { sourceRevision } from '../../test/support/target-source-fixture';
import { protectedDocument } from '../../test/support/protected-document';
import { formatTable, formatCell, formatRow } from '../../test/support/table-format-fixture';
import {
  namedChild,
  nodeChildren,
  tableRows,
  rowCells,
} from '../../test/support/native-table-oracle';
import { planTargetDocument } from './plan-target';
import {
  cell,
  table,
  part,
  nativeTable,
  content,
  plan,
  landed,
} from '../../test/support/target-table-fixture';

test('C01 tables and links plans cell text, one rich row, and the existing hyperlink without replacing the table', async () => {
  const input = await targetDocumentInput(
      table([
        ['A', 'B'],
        ['One', 'Two'],
      ]) + '<w:p><w:hyperlink r:id="link"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>',
      '<Relationship Id="link" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>'
    ),
    oldTable = nativeTable(input),
    oldRows = tableRows(part(input), oldTable.id),
    oldCell = rowCells(oldRows[1]!)[1]!,
    oldParagraph = storyParagraphs(oldCell)[0]!,
    bytes = writeOoxmlPackage(input.pkg),
    { result, input: applied } = await landed(input, pairs.cases[1]!.afterMarkdown);
  expect(result.plan.operations.map((op) => op.kind)).toEqual(['text', 'insert-row', 'format']);
  expect(result.plan.operations[0]).toMatchObject({
    kind: 'text',
    paragraphId: oldParagraph.id,
    before: 'Two',
    after: 'Changed',
  });
  expect(result.plan.operations[1]).toMatchObject({
    kind: 'insert-row',
    intent: {
      tableId: oldTable.id,
      rowId: oldRows[1]!.id,
      where: 'below',
      cells: [
        [{ type: 'paragraph', children: [{ type: 'text', value: 'Three' }] }],
        [{ type: 'paragraph', children: [{ type: 'text', value: 'Four' }] }],
      ],
    },
  });
  expect(nativeTable(applied).id).toBe(oldTable.id);
  expect(content(applied)).toEqual([
    [['A'], ['B']],
    [['One'], ['Changed']],
    [['Three'], ['Four']],
  ]);
  expect(findNode(part(applied), oldRows[0]!.id)).toBe(oldRows[0]);
  for (const property of ['tblPr', 'tblGrid'])
    expect(namedChild(nativeTable(applied), property)).toBe(namedChild(oldTable, property));
  expect(namedChild(findNode(part(applied), oldRows[1]!.id)!, 'trPr')).toBe(
    namedChild(oldRows[1]!, 'trPr')
  );
  expect(namedChild(findNode(part(applied), oldCell.id)!, 'tcPr')).toBe(
    namedChild(oldCell, 'tcPr')
  );
  expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
});

for (const position of ['start', 'middle', 'end'] as const)
  test(`multiple row insertions at ${position} retain original anchors and target order`, async () => {
    const input = await targetDocumentInput(
        table([
          ['Head', 'Value'],
          ['A', 'B'],
          ['Tail', 'End'],
        ])
      ),
      oldRows = tableRows(part(input), nativeTable(input).id),
      lines = ['| Head | Value |', '| --- | --- |', '| A | B |', '| Tail | End |'];
    const at = position === 'start' ? 0 : position === 'middle' ? 3 : 4;
    lines.splice(
      at,
      0,
      '| New 1 | New 2 |',
      ...(position === 'start' ? ['| --- | --- |'] : []),
      '| New 3 | New 4 |'
    );
    if (position === 'start') lines.splice(4, 1);
    const { input: applied } = await landed(input, lines.join('\n'));
    for (const old of oldRows) expect(findNode(part(applied), old.id)).toBe(old);
    expect(
      content(applied)
        .flat(2)
        .filter((text) => text?.startsWith('New'))
    ).toEqual(['New 1', 'New 2', 'New 3', 'New 4']);
  });

test('row deletion and retained cell formatting preserve unrelated native content and row identities', async () => {
  const input = await targetDocumentInput(
      table([
        ['A', 'B'],
        ['One', 'Two'],
        ['Delete', 'Me'],
        ['Tail', 'End'],
      ])
    ),
    oldRows = tableRows(part(input), nativeTable(input).id),
    { result, input: applied } = await landed(
      input,
      '| A | B |\n| --- | --- |\n| **One** | Two |\n| Tail | End |'
    );
  expect(result.plan.operations.map((op) => op.kind)).toEqual(['format', 'table']);
  expect(result.plan.operations[1]).toMatchObject({
    kind: 'table',
    operation: { op: 'deleteTableRow', rowId: oldRows[2]!.id },
  });
  expect(findNode(part(applied), oldRows[2]!.id)).toBeNull();
  expect(findNode(part(applied), oldRows[3]!.id)).toBe(oldRows[3]);
});

test('new rich rows retain explicit marks and create links only during native application', async () => {
  const input = await targetDocumentInput(
      table([
        ['A', 'B'],
        ['One', 'Two'],
      ])
    ),
    bytes = writeOoxmlPackage(input.pkg),
    { result, input: applied } = await landed(
      input,
      '| A | B |\n| --- | --- |\n| One | Two |\n| **Bold** [Link](https://example.org/new) | ~~Gone~~ |'
    );
  expect(result.plan.operations).toHaveLength(1);
  expect(result.plan.operations[0]).toMatchObject({
    kind: 'insert-row',
    intent: {
      cells: [
        [
          {
            children: [
              { type: 'strong' },
              { type: 'text' },
              { type: 'link', url: 'https://example.org/new' },
            ],
          },
        ],
        [{ children: [{ type: 'delete' }] }],
      ],
    },
  });
  expect(input.pkg.externalTargets).toHaveLength(0);
  expect(
    applied.pkg.externalTargets.some((link) => link.rawTarget === 'https://example.org/new')
  ).toBe(true);
  expect(content(applied).at(-1)).toEqual([['Bold Link'], ['Gone']]);
  expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
});

test('native table operations preserve imported non-body parts and section settings through save/reopen', async () => {
  const entries = unzipSync(protectedDocument()),
    document = strFromU8(entries['word/document.xml']!),
    bodyStart = document.indexOf('<w:body>') + '<w:body>'.length,
    sectionStart = document.lastIndexOf('<w:sectPr>');
  entries['word/document.xml'] = strToU8(
    document.slice(0, bodyStart) +
      table([
        ['A', 'B'],
        ['One', 'Two'],
      ]) +
      document.slice(sectionStart)
  );
  const loaded = readOoxmlPackage(zipSync(entries));
  if (!loaded.ok) throw new Error(loaded.reason);
  const input = {
      pkg: loaded.package,
      revision: sourceRevision,
      projection: await projectSourcePackage(loaded.package),
    },
    oldSection = nodeChildren(bodyStoryRoot(part(input))!).at(-1)!,
    before = unzipSync(writeOoxmlPackage(input.pkg)),
    { input: applied } = await landed(
      input,
      '| A | B |\n| --- | --- |\n| One | Changed |\n| Three | Four |'
    ),
    saved = writeOoxmlPackage(applied.pkg),
    after = unzipSync(saved);
  expect(findNode(part(applied), oldSection.id)).toBe(oldSection);
  expect(Object.keys(after).sort()).toEqual(Object.keys(before).sort());
  for (const [name, bytes] of Object.entries(before))
    if (name !== 'word/document.xml') expect(after[name]).toEqual(bytes);
  const reopened = readOoxmlPackage(saved);
  if (!reopened.ok) throw new Error(reopened.reason);
  expect(
    canonicalOoxmlFingerprint(
      nodeChildren(bodyStoryRoot(reopened.package.parts.get(input.pkg.mainDocumentPart)!)!).at(-1)!
    )
  ).toBe(canonicalOoxmlFingerprint(oldSection));
});

test('new row nested tables retain their complete original target fragments', async () => {
  const input = await targetDocumentInput(
      table([
        ['A', 'B'],
        ['One', 'Two'],
      ])
    ),
    target =
      '<table>\n' +
      [
        ['A', 'B'],
        ['One', 'Two'],
        ['| X | Y |\n| --- | --- |\n| New X | New Y |', 'End'],
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
  if (edit.kind !== 'insert-row') throw new Error('Expected row insertion');
  expect(edit.intent.cells[0]![0]).toMatchObject({
    type: 'table',
    children: [
      {
        type: 'tableRow',
        children: [
          { type: 'tableCell', children: [{ type: 'text', value: 'X' }] },
          { type: 'tableCell', children: [{ type: 'text', value: 'Y' }] },
        ],
      },
      { type: 'tableRow' },
    ],
  });
  const pkg = applyTargetPlanOracle(input.pkg, result.plan),
    reopened = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!reopened.ok) throw new Error(reopened.reason);
  for (const saved of [pkg, reopened.package]) {
    const main = saved.parts.get(input.pkg.mainDocumentPart)!,
      outer =
        findNode(main, nativeTable(input).id) ??
        nodeChildren(bodyStoryRoot(main)!).find((node) => node.kind === 'table')!,
      inserted = rowCells(tableRows(main, outer.id).at(-1)!),
      nested = nodeChildren(inserted[0]!).find((node) => node.kind === 'table')!;
    expect(nodeChildren(namedChild(nested, 'tblGrid')!)).toHaveLength(2);
    expect(
      tableRows(main, nested.id).map((row) =>
        rowCells(row).map((cell) =>
          storyParagraphs(cell).map((paragraph) => paragraphTextOf(main, paragraph.id))
        )
      )
    ).toEqual([
      [['X'], ['Y']],
      [['New X'], ['New Y']],
    ]);
    expect(storyParagraphs(inserted[1]!).map((node) => paragraphTextOf(main, node.id))).toEqual([
      'End',
    ]);
  }
});

test('repeated cell text stays bound to its own row and column', async () => {
  const input = await targetDocumentInput(
      table([
        ['Same', 'Same'],
        ['Same', 'Same'],
        ['Same', 'Same'],
      ])
    ),
    oldRows = tableRows(part(input), nativeTable(input).id),
    oldCells = oldRows.map(rowCells),
    changed = storyParagraphs(oldCells[1]![1]!)[0]!,
    { result, input: applied } = await landed(
      input,
      '| Same | Same |\n| --- | --- |\n| Same | Changed |\n| Same | Same |'
    );
  expect(
    result.plan.operations.every((op) => op.kind === 'text' && op.paragraphId === changed.id)
  ).toBe(true);
  for (const cells of oldCells)
    for (const cell of cells)
      if (cell !== oldCells[1]![1]) expect(findNode(part(applied), cell.id)).toBe(cell);
});

test('visible merged cell edits preserve native grid spans and hidden continuations', async () => {
  const input = await targetDocumentInput(
      formatTable(
        '<w:tr>' +
          formatCell('AB', 3600, '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>') +
          formatCell('C1', 3600) +
          '</w:tr>' +
          '<w:tr>' +
          formatCell('', 3600, '<w:gridSpan w:val="2"/><w:vMerge/>') +
          formatCell('C2', 3600) +
          '</w:tr>' +
          formatRow('A3', 'B3', 'C3')
      )
    ),
    oldRows = tableRows(part(input), nativeTable(input).id),
    oldCells = oldRows.map(rowCells),
    target = input.projection.text.replace('C2', 'Changed'),
    { input: applied } = await landed(input, target);
  expect(findNode(part(applied), oldCells[0]![0]!.id)).toBe(oldCells[0]![0]);
  expect(findNode(part(applied), oldCells[1]![0]!.id)).toBe(oldCells[1]![0]);
  expect(namedChild(findNode(part(applied), oldCells[1]![1]!.id)!, 'tcPr')).toBe(
    namedChild(oldCells[1]![1]!, 'tcPr')
  );
  expect(
    storyParagraphs(findNode(part(applied), oldCells[1]![1]!.id)!).map((node) =>
      paragraphTextOf(part(applied), node.id)
    )
  ).toEqual(['Changed']);
});

test('nested table cell changes retain the nearest native table and outer cell content', async () => {
  const nested = table([
      ['Inner A', 'Inner B'],
      ['Inner 1', 'Inner 2'],
    ]),
    outer =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="4800"/></w:tblGrid><w:tr>' +
      cell('Outer') +
      '<w:tc><w:tcPr/>' +
      nested +
      p('Outer tail') +
      '</w:tc></w:tr></w:tbl>',
    input = await targetDocumentInput(outer),
    target = input.projection.text.replace('Inner 2', 'Changed inner'),
    sourceCell = input.projection.cells.find((cell) =>
      cell.sources.some(
        (source) =>
          source.markdownStart <= input.projection.text.indexOf('Inner 2') &&
          source.markdownEnd >= input.projection.text.indexOf('Inner 2') + 'Inner 2'.length
      )
    )!,
    { result, input: applied } = await landed(input, target);
  expect(
    result.plan.operations.every(
      (op) => op.kind === 'text' && sourceCell.paragraphIds.includes(op.paragraphId)
    )
  ).toBe(true);
  const outerCell = rowCells(tableRows(part(input), nativeTable(input).id)[0]!)[1]!,
    outerTail = nodeChildren(outerCell).find((node) => node.kind === 'paragraph')!;
  expect(findNode(part(applied), outerTail.id)).toBe(outerTail);
});

test('unsupported list continuation destinations refuse atomically after earlier retained text has planned edits', async () => {
  const input = await targetDocumentInput(
      p('Earlier') +
        table([
          ['A', 'B'],
          ['One', 'Two'],
        ])
    ),
    bytes = writeOoxmlPackage(input.pkg),
    result = planTargetDocument(
      input,
      'Changed earlier\n\n<table>\n' +
        [
          ['A', 'B'],
          ['One', 'Two'],
          ['- First paragraph\n\n  Continuation paragraph', 'New value'],
        ]
          .map(
            (values) =>
              '<tr>\n' +
              values.map((value) => '<td>\n\n' + value + '\n\n</td>').join('\n') +
              '\n</tr>'
          )
          .join('\n') +
        '\n</table>'
    );
  expect(result.success).toBe(false);
  if (!result.success) expect(result.error).toContain('continuation');
  expect('plan' in result).toBe(false);
  expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
});

test('unresolved table row groups refuse the complete request without writes', async () => {
  const input = await targetDocumentInput(
      table([
        ['A', 'B'],
        ['One', 'Two'],
      ])
    ),
    bytes = writeOoxmlPackage(input.pkg),
    fingerprint = canonicalOoxmlFingerprint(nativeTable(input));
  for (const target of [
    '<table><thead><tr><td>A</td><td>B</td></tr></thead><tbody><tr><td>One</td><td>Changed</td></tr></tbody></table>',
  ]) {
    const result = planTargetDocument(input, target);
    expect(result.success).toBe(false);
    expect('plan' in result).toBe(false);
  }
  expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
  expect(canonicalOoxmlFingerprint(nativeTable(input))).toBe(fingerprint);
});
