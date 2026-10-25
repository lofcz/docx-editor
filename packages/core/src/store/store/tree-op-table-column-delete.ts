import {
  createNodeIdAllocator,
  insertChildren,
  removeNode,
  replaceNode,
  applyEdits,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlPart } from '../package/ooxml-tree.ts';
import { paragraphIdsWithin } from './tree-op-blocks.ts';
import { fromEdit, TEXT_DEPS } from './tree-op-nodes.ts';
import type { EditableTableTopology } from './tree-op-table-topology.ts';
import type { TreeOpEffect, TreeOpResult } from './tree-op-types.ts';

export interface DeleteColumnTarget {
  readonly kind: 'gridColumn' | 'referenceCell';
  readonly id: string;
}

export interface DeleteColumnHelpers {
  readonly firstParagraphId: (
    cell: EditableTableTopology['rows'][number]['cells'][number]
  ) => string | null;
  readonly patchGridColumn: (
    grid: OoxmlElement,
    id: string
  ) => { readonly ok: true; readonly grid: OoxmlElement } | { readonly ok: false };
  readonly synthesizeGrid: (
    topology: EditableTableTopology,
    nextId: () => string,
    index: number
  ) => OoxmlElement;
  readonly gridChildIndex: (table: OoxmlElement) => number;
}

/** Apply a validated reference-cell deletion, including grid repair for imported tables. */
export function applyResolvedDeleteColumn(
  part: OoxmlPart,
  topology: EditableTableTopology,
  target: DeleteColumnTarget,
  columnIndex: number,
  options: EditOptions | undefined,
  helpers: DeleteColumnHelpers
): TreeOpResult {
  const deletedCells = topology.rows.map((entry) => entry.cells[columnIndex]!);
  const deletedParagraphs = deletedCells.flatMap((cell) => paragraphIdsWithin(cell));
  const caretSourceIndex =
    columnIndex < topology.rows[0]!.cells.length - 1 ? columnIndex + 1 : columnIndex - 1;
  const caretParagraphId = helpers.firstParagraphId(topology.rows[0]!.cells[caretSourceIndex]!);
  if (!caretParagraphId) return { ok: false, reason: 'tree-invariant' };
  const completeGrid =
    topology.grid !== undefined && topology.gridColumns.length === topology.rows[0]!.cells.length;
  const nextId = createNodeIdAllocator(part);
  const repairedGrid = !completeGrid
    ? helpers.synthesizeGrid(topology, nextId, columnIndex)
    : undefined;
  const patch =
    completeGrid && topology.grid && target.kind === 'gridColumn'
      ? helpers.patchGridColumn(topology.grid, target.id)
      : undefined;
  if (completeGrid && (!patch || !patch.ok)) return { ok: false, reason: 'unknown-grid-column' };
  const patchedGrid = patch?.ok ? patch.grid : undefined;
  const edits: ((current: OoxmlPart) => ReturnType<typeof removeNode>)[] = deletedCells.map(
    (cell) => (current) => removeNode(current, cell.id, options)
  );
  edits.push((current) =>
    !completeGrid
      ? topology.grid
        ? replaceNode(current, topology.grid.id, repairedGrid!, options)
        : insertChildren(
            current,
            topology.table.id,
            helpers.gridChildIndex(topology.table),
            [repairedGrid!],
            options
          )
      : replaceNode(current, topology.grid!.id, patchedGrid!, options)
  );
  const created = repairedGrid
    ? [
        repairedGrid.id,
        ...repairedGrid.children
          .filter((child) => child.kind !== 'textValue')
          .map((child) => child.id),
      ]
    : [];
  const deleted = [
    ...(target.kind === 'gridColumn' ? [target.id] : []),
    ...deletedCells.map((cell) => cell.id),
    ...deletedParagraphs,
    ...(repairedGrid && topology.grid ? [topology.grid.id] : []),
  ];
  const effect: TreeOpEffect = {
    dirty: [
      topology.table.id,
      ...(topology.grid ? [topology.grid.id] : []),
      ...topology.rows.map((entry) => entry.row.id),
    ],
    created,
    deleted,
    dependencyKeys: TEXT_DEPS,
    impact: 'flow-structural',
    caret: { paragraphId: caretParagraphId },
  };
  return fromEdit(applyEdits(part, edits, options), effect);
}
