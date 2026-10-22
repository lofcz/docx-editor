// Resolved revision views omit pending cell insertions and deletions.

import { expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlNode } from '@docx-editor.dev/core/store';
import { readTableStructure } from '../semantic-table.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const marker = (name: string, author = 'Ada') =>
  `<w:${name} w:id="1" w:author="${author}" w:date="2026-10-20T00:00:00Z"/>`;
const cell = (text: string, width: number, extra = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
function table(rows: string, widths: readonly number[], history = '') {
  const xml = `<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr><w:tblW w:w="${widths.reduce((a, b) => a + b, 0)}" w:type="dxa"/></w:tblPr><w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}${history}</w:tblGrid>${rows}</w:tbl></w:body></w:document>`;
  const loaded = readOoxmlPart(xml, { name: '/word/document.xml', contentType: 'application/xml' });
  if (!loaded.ok) throw new Error(loaded.reason);
  const body = (loaded.part.root as { children: readonly OoxmlNode[] }).children[0] as {
    children: readonly OoxmlNode[];
  };
  return body.children[0]!;
}
const shape = (node: OoxmlNode, mode: 'original' | 'proposed' | 'all-markup') =>
  readTableStructure(node, 500, 0, undefined, mode)!.rows.map((row) =>
    row.cells.map((cell) => cell.gridSpan)
  );

test('the Proposed view omits a deleted cell and gives its span to the previous cell', () => {
  const node = table(
    `<w:tr>${cell('a', 1200)}${cell('gone', 2400, marker('cellDel'))}${cell('c', 3600)}</w:tr>` +
      `<w:tr>${cell('gone', 1200, marker('cellDel'))}${cell('e', 2400)}${cell('f', 3600)}</w:tr>`,
    [1200, 2400, 3600]
  );
  expect(shape(node, 'all-markup')).toEqual([
    [1, 1, 1],
    [1, 1, 1],
  ]);
  expect(shape(node, 'original')).toEqual([
    [1, 1, 1],
    [1, 1, 1],
  ]);
  // A deleted first cell gives its span to the next survivor.
  expect(shape(node, 'proposed')).toEqual([
    [2, 1],
    [2, 1],
  ]);
});

test('the Original view omits an inserted cell and keeps the recorded grid', () => {
  const history =
    '<w:tblGridChange w:id="2"><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="3600"/></w:tblGrid></w:tblGridChange>';
  const node = table(
    `<w:tr>${cell('a', 1200)}${cell('new', 2400, marker('cellIns'))}${cell('c', 3600)}</w:tr>`,
    [1200, 2400, 3600],
    history
  );
  expect(shape(node, 'proposed')).toEqual([[1, 1, 1]]);
  expect(shape(node, 'original')).toEqual([[1, 1]]);
});

test('a pending cell merge reads its proposed and original vertical merge', () => {
  const merge = (value: string) => `<w:cellMerge w:id="3" w:author="Ada" w:vMerge="${value}"/>`;
  const node = table(
    `<w:tr>${cell('a', 1200, merge('rest'))}${cell('b', 2400)}</w:tr>` +
      `<w:tr>${cell('d', 1200, merge('cont'))}${cell('e', 2400)}</w:tr>`,
    [1200, 2400]
  );
  const continued = (mode: 'original' | 'proposed' | 'all-markup') =>
    readTableStructure(node, 500, 0, undefined, mode)!.rows[1]!.cells[0]!.vMergeContinue === true;
  expect(continued('proposed')).toBe(true);
  expect(continued('original')).toBe(false);
  expect(continued('all-markup')).toBe(false);
});

for (const deleted of [0, 1, 2]) {
  test(`resolved export compacts deleted column ${deleted} while physical coordinates stay native`, () => {
    const node = table(
      Array.from(
        { length: 2 },
        () =>
          '<w:tr>' +
          [0, 1, 2]
            .map((index) => cell(String(index), 1200, index === deleted ? marker('cellDel') : ''))
            .join('') +
          '</w:tr>'
      ).join(''),
      [1200, 1200, 1200]
    );
    const before = JSON.stringify(node);
    const proposed = readTableStructure(node, 500, 0, undefined, 'proposed')!;
    for (const row of proposed.rows) {
      expect(
        row.cells.map((cell) => [
          cell.resolvedGridColumn,
          cell.resolvedGridSpan,
          cell.resolvedGridColumnCount,
        ])
      ).toEqual([
        [0, 1, 2],
        [1, 1, 2],
      ]);
      expect(row.cells.reduce((total, cell) => total + cell.gridSpan, 0)).toBe(3);
    }
    expect(proposed.columnWidthsPt).toHaveLength(3);
    for (const mode of ['original', 'all-markup'] as const)
      expect(
        readTableStructure(node, 500, 0, undefined, mode)!.rows[0]!.cells.every(
          (cell) => cell.resolvedGridColumn === undefined
        )
      ).toBe(true);
    expect(JSON.stringify(node)).toBe(before);
  });
}

test('an authored merged cell keeps its original grid without revision omissions', () => {
  const node = table(
    '<w:tr>' + cell('Merged', 2400, '<w:gridSpan w:val="2"/>') + '</w:tr>',
    [1200, 1200]
  );
  for (const mode of ['original', 'proposed', 'all-markup'] as const) {
    const cell = readTableStructure(node, 500, 0, undefined, mode)!.rows[0]!.cells[0]!;
    expect(cell.gridSpan).toBe(2);
    expect(cell.resolvedGridColumn).toBeUndefined();
    expect(cell.resolvedGridColumnCount).toBeUndefined();
  }
});

test('an inserted cell later covered by a merge is absent from both resolved views', () => {
  const node = table(
      '<w:tr>' +
        cell('a', 1200) +
        cell('new', 2400, marker('cellIns') + marker('cellDel')) +
        cell('c', 3600) +
        '</w:tr>',
      [1200, 2400, 3600]
    ),
    before = JSON.stringify(node);
  expect(shape(node, 'original')).toEqual([[1, 1]]);
  expect(shape(node, 'proposed')).toEqual([[2, 1]]);
  expect(shape(node, 'all-markup')).toEqual([[1, 1, 1]]);
  expect(JSON.stringify(node)).toBe(before);
});
test('a rejected inserted owner donates no width when a separate deletion projects Proposed', () => {
  const node = table(
    '<w:tr>' +
      cell('a', 1200) +
      cell('new', 2400, marker('cellIns', 'Ada') + marker('cellDel', 'Grace')) +
      cell('c', 3600) +
      '</w:tr>',
    [1200, 2400, 3600]
  );
  const result = readTableStructure(node, 500, 0, undefined, 'proposed', {
    hiddenAuthors: new Set(['Ada']),
    includesNode: (_id, author) => author !== 'Ada',
    excludedNodeMode: () => 'original',
    cacheKey: 'reject-insertion-accept-deletion',
  })!;
  expect(result.rows[0]!.cells.map((cell) => cell.gridSpan)).toEqual([1, 1]);
});
test('a row inserted and deleted in one native frame is absent from both resolved views', () => {
  const node = table(
      '<w:tr><w:trPr>' +
        marker('ins') +
        marker('del') +
        '</w:trPr>' +
        cell('gone', 1200) +
        '</w:tr><w:tr>' +
        cell('kept', 1200) +
        '</w:tr>',
      [1200]
    ),
    before = JSON.stringify(node);
  expect(shape(node, 'original')).toEqual([[1]]);
  expect(shape(node, 'proposed')).toEqual([[1]]);
  expect(shape(node, 'all-markup')).toEqual([[1], [1]]);
  expect(JSON.stringify(node)).toBe(before);
});
