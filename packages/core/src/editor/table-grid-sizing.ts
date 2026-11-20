import type { EditorCommand } from '../contracts/editor.ts';
import type { TableFragmentRecord } from '../layout/semantic-records.ts';
import {
  buildStyleCascadeTable,
  cascadeTableFormatting,
  type StyleCascadeTable,
} from '../layout/style-cascade.ts';
import { readTablePreferredWidth } from '../layout/table-widths.ts';
import { tableInteractionIndex } from '../layout/semantic-table-interaction.ts';
import { findNode } from '../store/package/ooxml-edit.ts';
import type { OoxmlElement } from '../store/package/ooxml-tree.ts';
import { MAX_TABLE_WIDTH_TWIPS } from '../store/store/table-constraints.ts';
import { readTwipsMeasure } from '../store/units.ts';
import { wmlAttributeValue, wmlChildNamed } from '../store/store/tree-op-table-shared.ts';
import { readEditableTableTopology } from '../store/store/tree-op-table-topology.ts';
import type { TreeDocOp } from '../store/store/tree-op-types.ts';
import type { TableCommandPlannerInput } from './table-command-plan.ts';

const tableStyles = new WeakMap<OoxmlElement, StyleCascadeTable>();

function tableWidthProperty(table: OoxmlElement, stylesRoot?: OoxmlElement | null) {
  const properties = wmlChildNamed(table, 'tblPr');
  const width = properties && wmlChildNamed(properties, 'tblW');
  if (width || !stylesRoot) return width;
  let styles = tableStyles.get(stylesRoot);
  if (!styles) {
    styles = buildStyleCascadeTable(stylesRoot);
    tableStyles.set(stylesRoot, styles);
  }
  const style = properties && wmlChildNamed(properties, 'tblStyle');
  const layers = cascadeTableFormatting(
    styles,
    style && wmlAttributeValue(style, 'val')
  ).tablePropertyNodes;
  for (let i = layers.length - 1; i >= 0; i--) {
    const inherited = wmlChildNamed(layers[i]!, 'tblW');
    if (inherited) return inherited;
  }
  return undefined;
}

/** A copied column shares the table's existing width. Each row retains its own cell preferences. */
export function columnInsertionOps(
  input: TableCommandPlannerInput,
  op: Extract<TreeDocOp, { op: 'insertTableColumn' }>
): readonly TreeDocOp[] {
  const table = findNode(input.part, op.tableId);
  if (table?.kind !== 'table') return [op];
  if (readTablePreferredWidth(tableWidthProperty(table, input.stylesRoot)).value > 0) return [op];
  const occurrence = tableInteractionIndex(input.layout).occurrences.find(
    (entry) => entry.table.tableId === op.tableId
  );
  if (!occurrence || input.layout.revision !== input.storeRevision) return [op];
  const width = Math.round(occurrence.table.box.width * 20);
  if (!(width > 0 && width <= MAX_TABLE_WIDTH_TWIPS)) return [op];
  return [
    op,
    {
      op: 'setTableProperties',
      tableId: op.tableId,
      preferredWidth: { type: 'dxa', value: width },
    },
  ];
}

/** Round boundaries before subtracting, so fractional columns retain their exact total. */
export function displayedColumnWidths(table: TableFragmentRecord): number[] {
  const edges = table.columnEdges.map((pt) => Math.round(pt * 20));
  const widths = edges.slice(1).map((edge, i) => edge - edges[i]!);
  return table.rows.some((row) => row.cells.some((cell) => cell.logicalGridColumn !== undefined))
    ? widths.reverse()
    : widths;
}

export function tablePropertyOp(
  input: TableCommandPlannerInput,
  tableId: string,
  command: Extract<EditorCommand, { type: 'setTableProperties' }>
): Extract<TreeDocOp, { op: 'setTableProperties' }> | null {
  const allowed = [
    'type',
    'width',
    'widthType',
    'justification',
    'styleId',
    'styleOptions',
    'autoFit',
  ];
  if (Object.keys(command).some((key) => !allowed.includes(key))) return null;
  const topo = readEditableTableTopology(input.part.root, tableId);
  if (!topo.ok) return null;
  const { autoFit, styleId, styleOptions, justification, width, widthType } = command;
  let preferredWidth: Extract<TreeDocOp, { op: 'setTableProperties' }>['preferredWidth'];
  if (width !== undefined || widthType !== undefined) {
    if (autoFit !== undefined) return null;
    if (width === null || widthType === null) preferredWidth = null;
    else {
      const props = wmlChildNamed(topo.topology.table, 'tblPr');
      const current = props && wmlChildNamed(props, 'tblW');
      const unit = widthType ?? (current && wmlAttributeValue(current, 'type')) ?? 'dxa';
      if (unit !== 'dxa' && unit !== 'pct' && unit !== 'auto') return null;
      const raw = current && wmlAttributeValue(current, 'w');
      const value =
        width ??
        (unit === 'auto' ? 0 : unit === 'pct' ? Number(raw ?? 5000) : (readTwipsMeasure(raw) ?? 0));
      preferredWidth = { type: unit, value };
    }
  }
  let columnWidthsTwips: number[] | undefined;
  if (autoFit === 'fixed') {
    const occurrence = tableInteractionIndex(input.layout).occurrences.find(
      (occ) => occ.table.tableId === tableId
    );
    if (!occurrence || input.layout.revision !== input.storeRevision) return null;
    columnWidthsTwips = displayedColumnWidths(occurrence.table);
  }
  const op = {
    op: 'setTableProperties' as const,
    tableId,
    ...(columnWidthsTwips ? { columnWidthsTwips } : {}),
    ...(autoFit !== undefined ? { autoFit } : {}),
    ...(styleId !== undefined ? { styleId } : {}),
    ...(styleOptions !== undefined ? { styleOptions } : {}),
    ...(justification !== undefined ? { justification } : {}),
    ...(preferredWidth !== undefined ? { preferredWidth } : {}),
  };
  return Object.keys(op).length > 2 ? op : null;
}
