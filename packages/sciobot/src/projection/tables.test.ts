import { expect, test } from 'bun:test';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import { docxFixture } from '../../test/support/document';
import { markdownRangeAt } from './markdown';
import { projectTables, tableCellAt } from './tables';

const revision = { documentId: 'doc', epoch: 'load-1', revision: 0 };
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (content: string) => `<w:tc>${content}</w:tc>`;
const row = (...cells: string[]) => `<w:tr>${cells.join('')}</w:tr>`;
const table = (rows: string) =>
  '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  rows +
  '</w:tbl>';
async function project(body: string) {
  const opened = openDocumentForExport(docxFixture(body));
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const layout = await opened.session.layout();
    const before = JSON.stringify(layout);
    const result = projectTables(layout, revision);
    if (!result.success) throw new Error(result.error);
    expect(JSON.stringify(layout)).toBe(before);
    return { layout, projection: result.projection };
  } finally {
    opened.session.dispose();
  }
}

test('repeated values resolve to the selected canonical table cell', async () => {
  const { layout, projection } = await project(
    table(row(cell(p('Same')), cell(p('Same'))) + row(cell(p('Same')), cell(p('Other'))))
  );
  expect(projection.text).toBe('| Same | Same |\n| --- | --- |\n| Same | Other |');
  const sourceTable = layout.pages[0]!.fragments[0]!;
  if (sourceTable.kind !== 'table') throw new Error('No native table');
  expect(projection.tables[0]?.id).toBe(sourceTable.tableId);
  expect(tableCellAt(projection, 9)?.id).toBe(sourceTable.rows[0]!.cells[1]!.id);
  expect(tableCellAt(projection, projection.text.lastIndexOf('Same'))?.id).toBe(
    sourceTable.rows[1]!.cells[0]!.id
  );
  expect(tableCellAt(projection, 7)).toBeNull();
  expect(projection.boundaries.map((boundary) => boundary.role)).toEqual(['cell', 'row', 'cell']);
  const separator = projection.boundaries[1]!.extent!;
  expect(projection.text.slice(separator.start, separator.end)).toBe(' |\n| --- | --- |\n| ');
  expect(JSON.parse(JSON.stringify(projection))).toEqual(projection);
});

test('pipes and hard breaks retain native cell offsets beside empty cells', async () => {
  const { projection } = await project(
    table(row(cell('<w:p><w:r><w:t>A|B</w:t><w:br/><w:t>C</w:t></w:r></w:p>'), cell('<w:p/>')))
  );
  expect(projection.text).toBe('| A\\|B<br>C |  |\n| --- | --- |');
  expect(projection.cells).toHaveLength(2);
  const sourceCell = tableCellAt(projection, 10)!;
  expect(sourceCell.id).toBe(projection.cells[0]!.id);
  expect(markdownRangeAt(projection, 2, 11)).toMatchObject({
    start: { offset: 0 },
    end: { offset: 5 },
  });
  expect(tableCellAt(projection, 4)).toBeNull();
  expect(tableCellAt(projection, 8)).toBeNull();
  expect(projection.cells[1]?.paragraphIds).toHaveLength(1);
  expect(projection.cells[1]?.sources).toEqual([]);
  expect(projection.boundaries[0]?.extent).toBeNull();
});

test('nested table text belongs to its closest cell', async () => {
  const nested = table(row(cell(p('Nested')), cell(p('Inner'))));
  const { projection } = await project(
    table(row(cell(p('Before') + nested + p('After')), cell(p('Outer'))))
  );
  expect(projection.tables).toHaveLength(2);
  expect(projection.cells).toHaveLength(4);
  const parent = tableCellAt(projection, projection.text.indexOf('Before'))!;
  const inner = tableCellAt(projection, projection.text.indexOf('Nested'))!;
  expect(inner.id === parent.id).toBe(false);
  expect(projection.tables.find((value) => value.id === inner.tableId)?.parentCellId).toBe(
    parent.id
  );
  expect(tableCellAt(projection, projection.text.indexOf('After'))?.id).toBe(parent.id);
});

test('resolved grid spans and vertical continuations retain their native identities', async () => {
  const { projection } = await project(
    table(
      '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr>' +
        p('Merged') +
        '</w:tc></w:tr>' +
        '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr>'
    )
  );
  expect(projection.cells).toHaveLength(2);
  expect(projection.cells[0]).toMatchObject({
    gridColumn: 0,
    gridSpan: 2,
    verticalContinuation: false,
  });
  expect(projection.cells[1]).toMatchObject({
    gridColumn: 0,
    gridSpan: 2,
    verticalContinuation: true,
    sources: [],
  });
  expect(tableCellAt(projection, projection.text.indexOf('Merged'))?.id).toBe(
    projection.cells[0]!.id
  );
});

test('page continuations and repeated header rows retain one logical cell identity', async () => {
  const header =
    '<w:tr><w:trPr><w:tblHeader/></w:trPr>' + cell(p('Header')) + cell(p('Value')) + '</w:tr>';
  const content =
    '<w:p><w:r><w:t>First</w:t>' + '<w:br/><w:t>Continued</w:t>'.repeat(18) + '</w:r></w:p>';
  const { layout, projection } = await project(
    table(header + row(cell(content), cell(p('Tail')))) +
      '<w:sectPr><w:pgSz w:w="6000" w:h="2160"/><w:pgMar w:top="360" w:right="360" w:bottom="360" w:left="360"/></w:sectPr>'
  );
  expect(layout.pages.length > 1).toBe(true);
  expect(projection.tables).toHaveLength(1);
  expect(projection.cells).toHaveLength(4);
  expect(projection.tables[0]?.rowIds).toHaveLength(2);
  expect(projection.sourceMap.tables).toHaveLength(1);
  const construct = projection.sourceMap.tables[0]!;
  expect(construct.tableId).toBe(projection.tables[0]!.id);
  expect(projection.text.slice(construct.markdownStart, construct.markdownEnd)).toBe(
    projection.text
  );
  expect(projection.text.match(/Header/g)).toHaveLength(1);
  expect(projection.cells.every((value) => value.paragraphIds.length === 1)).toBe(true);
  expect(tableCellAt(projection, projection.text.lastIndexOf('Continued'))?.id).toBe(
    tableCellAt(projection, projection.text.indexOf('First'))?.id
  );
});
