import { expect, test } from 'bun:test';
import { isInsertedCellLists, cellListBindings } from './inserted-cell-lists';
import { createEditPlan } from './plan';
import { isInsertedRowIntent } from './inserted-row-types';
import { isInsertedColumnIntent } from './inserted-column-types';
import type { TargetAstNode } from '../projection/parse-target';
const cells: TargetAstNode[][] = [
  [{ type: 'paragraph', children: [{ type: 'text', value: 'A' }] }],
  [{ type: 'paragraph', children: [{ type: 'text', value: 'B' }] }],
];
const list = {
  path: [0, 0],
  target: {
    kind: 'ordered' as const,
    level: 0,
    instance: { kind: 'new' as const, key: 'shared', start: 3 },
  },
};
const row = {
  tableId: 'table',
  rowId: 'row',
  where: 'below',
  cells,
  properties: {},
  payloads: [],
  lists: [list],
};
const column = {
  tableId: 'table',
  anchor: { kind: 'grid-column', id: 'grid' },
  where: 'right',
  cells: cells.map((blocks, index) => ({ rowId: `row-${index}`, blocks })),
  properties: {},
  payloads: [],
  lists: [list],
};

test('new cell list codec rejects malformed, duplicate, nested, foreign, and conflicting destinations', () => {
  expect(isInsertedRowIntent(row)).toBe(true);
  expect(isInsertedColumnIntent(column)).toBe(true);
  for (const lists of [
    null,
    {},
    [null],
    [{ ...list, path: [] }],
    [{ ...list, path: [0] }],
    [{ ...list, path: [0, 0, 0] }],
    [{ ...list, path: [-1, 0] }],
    [{ ...list, path: [0, 2] }],
    [{ ...list, path: [9, 0] }],
    [list, list],
    [{ ...list, target: { kind: 'bullet', level: 9, instance: { kind: 'new', key: 'x' } } }],
    [
      list,
      {
        ...list,
        path: [1, 0],
        target: { kind: 'bullet', level: 0, instance: { kind: 'new', key: 'shared' } },
      },
    ],
  ]) {
    expect(isInsertedCellLists(cells, lists)).toBe(false);
    expect(isInsertedRowIntent({ ...row, lists })).toBe(false);
    expect(isInsertedColumnIntent({ ...column, lists })).toBe(false);
  }
  expect(cellListBindings([list], 0)).toEqual([{ ...list, path: [0] }]);
  expect(cellListBindings([list], 1)).toEqual([]);
});

test('shared new-list keys stay consistent across ordinary, row, and column intentions', () => {
  const conflicting = {
    ...column,
    lists: [
      { ...list, target: { kind: 'bullet', level: 0, instance: { kind: 'new', key: 'shared' } } },
    ],
  };
  const result = createEditPlan({ documentId: 'doc', epoch: '1', revision: 1 }, [
    { kind: 'insert-row', partName: '/word/document.xml', intent: row },
    { kind: 'insert-column', partName: '/word/document.xml', intent: conflicting },
  ] as unknown as Parameters<typeof createEditPlan>[1]);
  expect(result.success).toBe(false);
  if (!result.success) expect(result.error).toContain('conflicting-new-list-instance');
});
