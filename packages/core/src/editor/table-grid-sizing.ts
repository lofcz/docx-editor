import type { EditorCommand } from '../contracts/editor.ts';
import type { TableFragmentRecord } from '../layout/semantic-records.ts';
import { tableInteractionIndex } from '../layout/semantic-table-interaction.ts';
import { readTwipsMeasure } from '../store/units.ts';
import { wmlAttributeValue, wmlChildNamed } from '../store/store/tree-op-table-shared.ts';
import { readEditableTableTopology } from '../store/store/tree-op-table-topology.ts';
import type { TreeDocOp } from '../store/store/tree-op-types.ts';
import type { TableCommandPlannerInput } from './table-command-plan.ts';

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
