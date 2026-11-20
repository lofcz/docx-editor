import type { TableLook } from '../store/table-look.ts';
export { readTableLook, tableLookOf, type TableLook } from '../store/table-look.ts';

/**
 * ECMA-376 Part 1 §17.7.6 layers a table style's conditional formats weakest first: the whole table,
 * then the row bands, then the column bands, then first/last column, then first/last row,
 * then the four corners. A column band paints over a row band where the two cross.
 */
const CONDITION_PRECEDENCE = [
  'band1Horz',
  'band2Horz',
  'band1Vert',
  'band2Vert',
  'firstCol',
  'lastCol',
  'firstRow',
  'lastRow',
  'nwCell',
  'neCell',
  'swCell',
  'seCell',
] as const;

/**
 * How many leading rows form the table's header block.
 *
 * The header row condition does not stop at row 0. With the look's header row on, it covers
 * row 0 and then every following row that repeats as a header (`w:tblHeader`, 17.4.49),
 * whether or not row 0 itself repeats. With the header row off, the block is the run of
 * repeating rows that starts at row 0, or nothing. Rows in the block take no row band, and
 * banding counts from the first row after it. The last row format still reaches a block row
 * when the header row format is off.
 *
 * `repeats(rowIndex)` answers for one row; it is asked in order and only while the run lasts.
 */
export function headerBlockRowCount(
  look: TableLook,
  rowCount: number,
  repeats: (rowIndex: number) => boolean
): number {
  if (rowCount === 0 || (!look.firstRow && !repeats(0))) return 0;
  let count = 1;
  while (count < rowCount && repeats(count)) count += 1;
  return count;
}

/**
 * Which of the style's conditional formats apply to one cell, weakest first.
 *
 * Everything is derived from the look and the cell's place in the grid. A `w:cnfStyle`
 * (17.4.7, 17.4.8) is a record the producer wrote of the conditions it applied; it can be
 * stale after an edit and is not read. Structural conditions key on the GRID COLUMN the
 * cell occupies, so a `gridSpan` or a `w:gridBefore` earlier in the row cannot shift them.
 */
export function conditionalTypesFor(input: {
  readonly look: TableLook;
  readonly rowIndex: number;
  readonly rowCount: number;
  /** Leading rows of the header block ({@link headerBlockRowCount}). */
  readonly headerRows: number;
  readonly gridColumn: number;
  readonly gridSpan: number;
  readonly columnCount: number;
}): readonly string[] {
  const active = new Set<string>();
  const { look, rowIndex, rowCount, headerRows, gridColumn, gridSpan, columnCount } = input;
  const inHeaderBlock = rowIndex < headerRows;
  const isFirstRow = look.firstRow && inHeaderBlock;
  // The header row format, where it covers the last row, takes it from the last row format.
  const isLastRow = look.lastRow && rowIndex === rowCount - 1 && !isFirstRow;
  const isFirstColumn = look.firstColumn && gridColumn === 0;
  const isLastColumn = look.lastColumn && gridColumn + gridSpan >= columnCount;

  if (look.columnBanding && !isFirstColumn && !isLastColumn) {
    const band = gridColumn - (look.firstColumn ? 1 : 0);
    active.add(band % 2 === 0 ? 'band1Vert' : 'band2Vert');
  }
  if (look.rowBanding && !inHeaderBlock && !isLastRow) {
    const band = rowIndex - headerRows;
    active.add(band % 2 === 0 ? 'band1Horz' : 'band2Horz');
  }
  if (isFirstColumn) active.add('firstCol');
  if (isLastColumn) active.add('lastCol');
  if (isFirstRow) active.add('firstRow');
  if (isLastRow) active.add('lastRow');
  if (isFirstRow && isFirstColumn) active.add('nwCell');
  if (isFirstRow && isLastColumn) active.add('neCell');
  if (isLastRow && isFirstColumn) active.add('swCell');
  if (isLastRow && isLastColumn) active.add('seCell');

  const ordered: string[] = [];
  for (const condition of CONDITION_PRECEDENCE) if (active.has(condition)) ordered.push(condition);
  return ordered;
}

/**
 * The conditions a whole row is under: the cell conditions of one cell spanning the grid.
 *
 * Whether a row repeats (`w:tblHeader` through a conditional format) decides how far the
 * header block reaches, and the block decides the row's conditions, so the caller reads the
 * repeat flags first with only row 0 as the header row.
 */
export function rowConditionTypes(
  look: TableLook,
  rowIndex: number,
  rowCount: number,
  headerRows: number,
  columnCount: number
): readonly string[] {
  const columns = Math.max(1, columnCount);
  return conditionalTypesFor({
    look,
    rowIndex,
    rowCount,
    headerRows,
    gridColumn: 0,
    gridSpan: columns,
    columnCount: columns,
  });
}
