import {
  findNode,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';

interface PlacedNativeCell {
  readonly id: string;
  readonly row: number;
  readonly column: number;
  readonly span: number;
  readonly continuation: boolean;
}

const children = (node: OoxmlNode | null | undefined): readonly OoxmlNode[] =>
  !node || node.kind === 'textValue' ? [] : node.children;
const word = (node: OoxmlNode | undefined, name: string) =>
  children(node).find(
    (child) =>
      child.kind !== 'textValue' &&
      child.namespaceUri === WML_NAMESPACE_URI &&
      child.localName === name
  );
const value = (node: OoxmlNode | undefined): string | undefined =>
  node && node.kind !== 'textValue'
    ? node.attributes.find(
        (attribute) => attribute.localName === 'val' && attribute.namespaceUri === WML_NAMESPACE_URI
      )?.value
    : undefined;

/** Grid placement of every physical cell, including `w:gridBefore` offsets and spans. */
function placedCells(part: OoxmlPart, tableId: string): readonly PlacedNativeCell[] {
  const cells: PlacedNativeCell[] = [];
  children(findNode(part, tableId))
    .filter((node) => node.kind === 'tableRow')
    .forEach((row, rowIndex) => {
      let column = Number(value(word(word(row, 'trPr'), 'gridBefore')) ?? 0) || 0;
      for (const cell of children(row)) {
        if (cell.kind !== 'tableCell') continue;
        const properties = word(cell, 'tcPr'),
          span = Math.max(1, Number(value(word(properties, 'gridSpan')) ?? 1) || 1),
          merge = word(properties, 'vMerge');
        cells.push({
          id: cell.id,
          row: rowIndex,
          column,
          span,
          continuation: !!merge && value(merge) !== 'restart',
        });
        column += span;
      }
    });
  return cells;
}

/** Rows and grid columns covered by one visual cell, from its head cell. */
export function nativeVisualSize(
  part: OoxmlPart,
  tableId: string,
  cellId: string
): { readonly rows: number; readonly columns: number } | null {
  const cells = placedCells(part, tableId),
    head = cells.find((cell) => cell.id === cellId);
  if (!head || head.continuation) return null;
  let rows = 1;
  while (
    cells.some(
      (cell) =>
        cell.row === head.row + rows &&
        cell.column === head.column &&
        cell.span === head.span &&
        cell.continuation
    )
  )
    rows++;
  return { rows, columns: head.span };
}

/** Visual cell heads inside a split region after the split, row-major, excluding the kept head. */
export function splitResultCells(
  part: OoxmlPart,
  tableId: string,
  cellId: string,
  size: { readonly rows: number; readonly columns: number }
): readonly string[] | null {
  const cells = placedCells(part, tableId),
    head = cells.find((cell) => cell.id === cellId);
  if (!head) return null;
  return cells
    .filter(
      (cell) =>
        cell.id !== cellId &&
        !cell.continuation &&
        cell.row >= head.row &&
        cell.row < head.row + size.rows &&
        cell.column >= head.column &&
        cell.column < head.column + size.columns
    )
    .sort((a, b) => a.row - b.row || a.column - b.column)
    .map((cell) => cell.id);
}
