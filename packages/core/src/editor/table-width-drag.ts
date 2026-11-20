import { validateTreeOp } from '../store/store/tree-ops.ts';
import { MAX_TABLE_WIDTH_TWIPS } from '../store/store/table-constraints.ts';
import type { TableInteractionHit } from '../layout/semantic-table-interaction.ts';
import type {
  TableRowOccurrenceTarget,
  TableColumnDividerResizeTarget,
  TableRightEdgeResizeTarget,
} from '../contracts/editor.ts';
import {
  planTableCommand,
  type TableCommandPlan,
  type TableCommandPlannerInput,
} from './table-command-plan.ts';

export interface WidthDrag {
  readonly leftTwips: number;
  readonly rightTwips: number;
  readonly tableWidthTwips: number;
  readonly displayed: readonly number[];
  readonly leftIndex: number;
  readonly rightIndex: number;
  readonly normalizeGrid: boolean;
  readonly leftMinimum: number;
  readonly rightMinimum: number;
  readonly target: Extract<
    TableInteractionHit,
    { kind: 'columnDivider' | 'rightEdge' | 'rowDivider' }
  >;
  readonly resizeTarget:
    TableColumnDividerResizeTarget | TableRightEdgeResizeTarget | TableRowOccurrenceTarget;
  readonly rightEdge: boolean;
}

export function clampedWidthDelta(drag: WidthDrag, delta: number): number {
  const lower = drag.leftMinimum - drag.leftTwips;
  const upper = drag.rightEdge
    ? MAX_TABLE_WIDTH_TWIPS - drag.tableWidthTwips
    : drag.rightTwips - drag.rightMinimum;
  return Math.min(upper, Math.max(lower, delta));
}

export function planWidthDrag(
  drag: WidthDrag,
  moveTwips: number,
  input: Omit<TableCommandPlannerInput, 'command'>
): TableCommandPlan {
  const delta = clampedWidthDelta(drag, moveTwips);
  if (drag.normalizeGrid) {
    // AutoFit and percentage tables can place each boundary independently of tblGrid.
    // Freeze the displayed grid before fixed layout can redistribute the other columns.
    if (
      drag.target.sourceRevision !== input.storeRevision ||
      !input.editable ||
      input.viewing ||
      drag.target.isHeaderRepeat
    )
      return { ok: false, code: 'unsupported', reason: 'the resize target is no longer editable' };
    if (!delta)
      return { ok: false, code: 'invalidArgs', reason: 'the column is at its width limit' };
    const widths = [...drag.displayed];
    widths[drag.leftIndex] = drag.leftTwips + delta;
    if (!drag.rightEdge) widths[drag.rightIndex] = drag.rightTwips - delta;
    const op = {
      op: 'setTableProperties',
      tableId: drag.target.tableId,
      autoFit: 'fixed',
      columnWidthsTwips: widths,
    } as const;
    const invalid = validateTreeOp(input.part, op);
    return invalid
      ? { ok: false, code: 'unsupported', reason: invalid }
      : { ok: true, ops: [op], selection: { kind: 'preserveSelection' } };
  }
  return planTableCommand({
    ...input,
    command: drag.rightEdge
      ? {
          type: 'commitTableRightEdgeResize',
          target: drag.resizeTarget as TableRightEdgeResizeTarget,
          columnWidthTwips: drag.leftTwips + delta,
          tableWidthTwips: drag.tableWidthTwips + delta,
        }
      : {
          type: 'commitTableColumnDividerResize',
          target: drag.resizeTarget as TableColumnDividerResizeTarget,
          leftWidthTwips: drag.leftTwips + delta,
          rightWidthTwips: drag.rightTwips - delta,
        },
  });
}
