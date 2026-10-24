import { expect, test } from 'bun:test';
import { attributeTrackedOps, isTrackedEdit } from '../surface-tracked-ops';
import type { TreeDocOp } from '../../store/store/tree-op-types';
const revision = { author: 'Reviewer', date: '2026-10-24T12:00:00Z' };
const ops: TreeDocOp[] = [
  { op: 'setTableProperties', tableId: 't', indentTwips: 360 },
  {
    op: 'setTableColumnWidths',
    tableId: 't',
    leftGridColumnId: 'a',
    rightGridColumnId: 'b',
    leftWidthTwips: 1000,
    rightWidthTwips: 2000,
  },
  {
    op: 'setTableRightEdgeWidth',
    tableId: 't',
    gridColumnId: 'a',
    columnWidthTwips: 1000,
    tableWidthTwips: 1000,
  },
  { op: 'setTableRowHeight', tableId: 't', rowId: 'r', heightTwips: 720 },
  { op: 'setTableCellBorders', tableId: 't', cellIds: ['c'], scope: 'none', target: 'all' },
  { op: 'setTableCellFill', tableId: 't', cellIds: ['c'], color: null },
  { op: 'setTableCellVerticalAlignment', tableId: 't', cellIds: ['c'], alignment: 'center' },
];
for (const op of ops)
  test(`${op.op} receives native revision attribution and obeys doNotTrackFormatting`, () => {
    const attributed = attributeTrackedOps([op], revision, true)[0]!;
    expect(attributed).toEqual({ ...op, revision });
    expect(isTrackedEdit(attributed)).toBe(true);
    expect(attributeTrackedOps([op], revision, false)[0]).toBe(op);
    const other = { ...op, revision: { author: 'Other', date: revision.date } } as TreeDocOp;
    expect(attributeTrackedOps([other], revision, true)[0]).toBe(other);
  });
