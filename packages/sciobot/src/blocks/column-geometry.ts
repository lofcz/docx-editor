import {
  contentControlContentNodeOf,
  contentControlPropertiesOf,
  escapeXml,
  findNode,
  paragraphTextOf,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { activeStorage } from '../registry/active';
import { buildBlockIndex, type BlockIndex, type IndexedBlock } from './block-index';
import { blockIdFromTag } from './payload-opc-codec';

const MAX_COLUMNS = 64;
const MIN_WIDTH = 240;
const MAX_WIDTH = 31680;
// Percentage metadata is serialized to ten decimals. Quantize sub-twip noise before
// rounding a boundary, so equivalent ratios retain the same half-twip tie direction.
const roundedTwip = (value: number) => Math.round(Number(value.toFixed(7)));
const caption = () => `${activeStorage().adapterPrefix}:columns:1`;
const children = (node: OoxmlNode | undefined | null): readonly OoxmlNode[] =>
  !node || node.kind === 'textValue' ? [] : node.children;
const word = (node: OoxmlNode, name: string) =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && node.localName === name;
const attr = (node: OoxmlNode | undefined, name: string) =>
  node && node.kind !== 'textValue'
    ? node.attributes.find(
        (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === name
      )?.value
    : undefined;

/** Read the stored relative weights without changing their scale or missing-value defaults. */
export function columnPercentages(values: readonly unknown[]): readonly number[] {
  if (!values.length || values.length > MAX_COLUMNS) throw new Error('Invalid column count.');
  return values.map((value) => {
    if (value === undefined || value === null) return 100 / values.length;
    if (typeof value !== 'string' || value.length > 32 || !/^(?:\d+(?:\.\d+)?|\.\d+)%$/.test(value))
      throw new Error('Column widths must use percentages.');
    const number = Number(value.slice(0, -1));
    // These percentages are relative weights. Concurrent count changes can retain a
    // total above 100; resizing a survivor can then give it an individual weight above 100.
    // The count, text length, and resulting native grid remain independently bounded.
    if (!Number.isFinite(number) || number <= 0)
      throw new Error('Column percentages must be finite and greater than zero.');
    return number;
  });
}

/** Convert relative application widths to one exact native grid. Never allocate from a file number. */
export function columnWidthsTwips(
  values: readonly unknown[],
  available: number
): readonly number[] {
  if (
    !Number.isInteger(available) ||
    available > MAX_WIDTH ||
    available < values.length * MIN_WIDTH
  )
    throw new Error('Column geometry exceeds the supported native grid.');
  const weights = columnPercentages(values);
  const sum = weights.reduce((total, weight) => total + weight, 0);
  let previous = 0,
    cumulative = 0;
  return weights.map((weight) => {
    cumulative += weight;
    const edge = roundedTwip((available * cumulative) / sum),
      width = edge - previous;
    previous = edge;
    if (width < MIN_WIDTH) throw new Error('A column is narrower than the native minimum.');
    return width;
  });
}

/** Only renderer-produced column XML enters this builder. Attribute values remain escaped. */
export function columnTableXml(
  widths: readonly number[],
  render: (index: number, width: number) => string
): string {
  if (
    !widths.length ||
    widths.length > MAX_COLUMNS ||
    widths.some((width) => !Number.isInteger(width) || width < MIN_WIDTH) ||
    widths.reduce((sum, width) => sum + width, 0) > MAX_WIDTH
  )
    throw new Error('Invalid native column widths.');
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((edge) => `<w:${edge} w:val="nil"/>`)
    .join('');
  const grid = widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('');
  const cells = widths
    .map(
      (width, index) =>
        `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${render(index, width)}</w:tc>`
    )
    .join('');
  // Width authoring uses an absolute grid. Keep creation in that same coordinate system;
  // Word's percentage table width also includes outside cell margins and would jump on resize.
  const total = widths.reduce((sum, width) => sum + width, 0);
  return `<w:tbl><w:tblPr><w:tblW w:w="${total}" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCaption w:val="${escapeXml(caption())}"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid><w:tr>${cells}</w:tr></w:tbl>`;
}

export interface ColumnGeometry {
  readonly group: IndexedBlock;
  readonly tableId: string;
  readonly rowId: string;
  readonly columns: readonly {
    readonly block: IndexedBlock;
    readonly cellId: string;
    readonly width: number;
  }[];
  readonly separatorIds: ReadonlySet<string>;
}

/** Recognize only the adapter's complete owned shape. Ordinary nested tables retain table ownership. */
export function readColumnGeometry(
  pkg: OoxmlPackage,
  group: IndexedBlock,
  index: BlockIndex = buildBlockIndex(pkg)
): ColumnGeometry | null {
  if (group.payload.type !== 'column_group') return null;
  const part = pkg.parts.get(group.partName);
  const control = part && findNode(part, group.controlId);
  if (!part || !control) return null;
  const content = contentControlContentNodeOf(control),
    blocks = children(content);
  const table = blocks[0];
  if (table?.kind !== 'table') return null;
  const properties = children(table).find((node) => word(node, 'tblPr'));
  if (
    attr(
      children(properties).find((node) => word(node, 'tblCaption')),
      'val'
    ) !== caption()
  )
    return null;
  const rows = children(table).filter((node) => node.kind === 'tableRow');
  if (children(table).some((node) => !['tblPr', 'tblGrid', 'tr'].some((name) => word(node, name))))
    return null;
  const grid = children(children(table).find((node) => word(node, 'tblGrid')));
  if (
    rows.length !== 1 ||
    !grid.length ||
    grid.length > MAX_COLUMNS ||
    grid.some((node) => !word(node, 'gridCol'))
  )
    return null;
  const cells = children(rows[0]).filter((node) => node.kind === 'tableCell');
  if (
    cells.length !== grid.length ||
    children(rows[0]).some((node) => !word(node, 'tc') && !word(node, 'trPr'))
  )
    return null;
  const columns: ColumnGeometry['columns'][number][] = [],
    separatorIds = new Set<string>();
  const separator = (node: OoxmlNode) => {
    if (
      node.kind !== 'paragraph' ||
      paragraphTextOf(part, node.id) !== '' ||
      !children(node).every((child) => word(child, 'pPr'))
    )
      return false;
    separatorIds.add(node.id);
    return true;
  };
  if (!blocks.slice(1).every(separator)) return null;
  for (const [position, cell] of cells.entries()) {
    const contents = children(cell).filter((node) => !word(node, 'tcPr'));
    const control = contents[0];
    const id =
      control?.kind === 'contentControl'
        ? blockIdFromTag(contentControlPropertiesOf(control).tag ?? '')
        : null;
    const column = id === null ? undefined : index.byId.get(id);
    const width = Number(attr(grid[position], 'w'));
    if (
      !column ||
      column.controlId !== control?.id ||
      column.partName !== group.partName ||
      column.parentId !== group.id ||
      column.payload.type !== 'column' ||
      !Number.isInteger(width) ||
      width < MIN_WIDTH ||
      width > MAX_WIDTH ||
      !contents.slice(1).every(separator)
    )
      return null;
    const cellProperties = children(cell).find((node) => word(node, 'tcPr'));
    if (children(cellProperties).some((node) => word(node, 'gridSpan') || word(node, 'vMerge')))
      return null;
    columns.push({ block: column, cellId: cell.id, width });
  }
  const total = columns.reduce((sum, column) => sum + column.width, 0);
  if (total > MAX_WIDTH) return null;
  const preferred = children(properties).find((node) => word(node, 'tblW'));
  const fixed = children(properties).find((node) => word(node, 'tblLayout'));
  const stated = Number(attr(preferred, 'w'));
  // In this complete fixed grid, Word treats grid widths as relative weights when their
  // sum exceeds a positive absolute table width. Concurrent insertions can produce that
  // valid shape. Keep the column controls in the same coordinate system as native layout;
  // a subsequent width edit then writes an exact grid without changing the group width.
  if (
    attr(fixed, 'type') === 'fixed' &&
    attr(preferred, 'type') === 'dxa' &&
    Number.isInteger(stated) &&
    stated > 0 &&
    stated < total
  ) {
    let cumulative = 0,
      previous = 0;
    for (let i = 0; i < columns.length; i++) {
      cumulative += columns[i]!.width;
      const edge = roundedTwip((stated * cumulative) / total);
      const width = edge - previous;
      if (width < MIN_WIDTH) return null;
      columns[i] = { ...columns[i]!, width };
      previous = edge;
    }
  }
  return { group, tableId: table.id, rowId: rows[0]!.id, columns, separatorIds };
}
