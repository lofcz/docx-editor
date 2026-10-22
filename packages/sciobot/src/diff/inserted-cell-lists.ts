import type { TargetAstNode } from '../projection/parse-target';
import { isInsertedListBindings, type InsertedListBinding } from './inserted-list-types';

/** Paths start with the physical cell stream, then its paragraph index. Native IDs arrive later. */
export type InsertedCellListBinding = InsertedListBinding;

/** Validate the complete cell namespace, including cross-cell new-instance definitions. */
export function isInsertedCellLists(
  cells: readonly (readonly TargetAstNode[])[],
  value: unknown
): value is readonly InsertedCellListBinding[] {
  if (
    !Array.isArray(value) ||
    value.some((binding) => !Array.isArray(binding?.path) || binding.path.length !== 2)
  )
    return false;
  return isInsertedListBindings(
    cells.map((children) => ({ type: 'cell', children })),
    value
  );
}

/** Drop only the cell prefix after its native destination is known. */
export function cellListBindings(
  bindings: readonly InsertedCellListBinding[] | undefined,
  cell: number
): readonly InsertedListBinding[] {
  return (bindings ?? [])
    .filter((binding) => binding.path[0] === cell)
    .map((binding) => ({ ...binding, path: binding.path.slice(1) }));
}
