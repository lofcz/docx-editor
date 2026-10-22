import { concatMarkdown, literalMarkdown, type MappedMarkdown } from './markdown-source-map.ts';

export function tableWidth(projected: number | undefined, edgeCount: number | undefined): number {
  return Math.max(projected ?? (edgeCount ?? 1) - 1, 1);
}

/**
 * Serialize the one table shape GFM cannot express without interpreting DOCX records.
 *
 * Plain `<table>/<tr>/<td>/<th>` markup, all on one line: element names survive strict
 * sanitizers (GitHub's pipeline and stock `rehype-sanitize` both keep table elements while
 * stripping classes and roles), so the structure is self-contained and needs no consumer CSS.
 * The default stays on one line. Multiline mode gives MDX consumers block-content boundaries.
 */
export function nestedTableHtml<TRow extends { readonly isHeaderRow: boolean }>(
  rows: readonly TRow[],
  width: number,
  cellMarkdown: (row: TRow, columnIndex: number) => MappedMarkdown,
  multiline = false,
  shape?: (row: TRow, columnIndex: number) => { columnSpan: number; rowSpan: number; skip: boolean }
): MappedMarkdown {
  const renderedRows = rows.map((row) => {
    const tag = row.isHeaderRow ? 'th' : 'td';
    const renderedCells: MappedMarkdown[] = [];
    for (let columnIndex = 0; columnIndex < width; columnIndex++) {
      const cell = shape?.(row, columnIndex);
      if (cell?.skip) continue;
      const attributes =
        (cell && cell.columnSpan > 1 ? ` colspan="${cell.columnSpan}"` : '') +
        (cell && cell.rowSpan > 1 ? ` rowspan="${cell.rowSpan}"` : '');
      renderedCells.push(
        concatMarkdown(
          [
            literalMarkdown(`<${tag}${attributes}>`),
            cellMarkdown(row, columnIndex),
            literalMarkdown(`</${tag}>`),
          ],
          multiline ? '\n\n' : ''
        )
      );
    }
    return concatMarkdown([literalMarkdown('<tr>'), ...renderedCells, literalMarkdown('</tr>')]);
  });
  return concatMarkdown([literalMarkdown('<table>'), ...renderedRows, literalMarkdown('</table>')]);
}
