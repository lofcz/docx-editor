import { isInsertedCellLists, type InsertedCellListBinding } from './inserted-cell-lists';
import type { TargetAstNode } from '../projection/parse-target';
import {
  isInsertedCellFragments,
  type InsertedCellPayloadBinding,
} from './inserted-cell-fragments';

export interface InsertedRowProperties {
  /** Explicit native exact height. Omission retains native row inheritance. */
  readonly heightTwips?: number;
}

export type InsertedRowPayloadBinding = InsertedCellPayloadBinding;

export interface InsertedTableRowIntent {
  readonly tableId: string;
  /** Existing canonical anchor; allocate the inserted row/cell/paragraph IDs during execution. */
  readonly rowId: string;
  readonly where: 'above' | 'below';
  /** One block stream per physical native cell, including hidden merge continuations. */
  readonly cells: readonly (readonly TargetAstNode[])[];
  readonly properties: InsertedRowProperties;
  readonly payloads: readonly InsertedRowPayloadBinding[];
  readonly lists?: readonly InsertedCellListBinding[];
}

const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const only = (value: object, keys: readonly string[]) =>
  Object.keys(value).every((key) => keys.includes(key));
const id = (value: unknown): value is string => typeof value === 'string' && value.length > 0;

/** Structural/data checks run only after the bounded plain-JSON gate. */
export function isInsertedRowIntent(value: unknown): value is InsertedTableRowIntent {
  return (
    record(value) &&
    only(value, ['tableId', 'rowId', 'where', 'cells', 'properties', 'payloads', 'lists']) &&
    id(value.tableId) &&
    id(value.rowId) &&
    (value.where === 'above' || value.where === 'below') &&
    Array.isArray(value.cells) &&
    value.cells.length > 0 &&
    record(value.properties) &&
    only(value.properties, ['heightTwips']) &&
    (!Object.hasOwn(value.properties, 'heightTwips') ||
      (Number.isSafeInteger(value.properties.heightTwips) &&
        Number(value.properties.heightTwips) >= 20 &&
        Number(value.properties.heightTwips) <= 31680)) &&
    isInsertedCellFragments(value.cells, value.payloads) &&
    (!Object.hasOwn(value, 'lists') ||
      isInsertedCellLists(value.cells as TargetAstNode[][], value.lists))
  );
}
