import { isInsertedCellAtoms } from './inserted-cell-atoms';
import type { InsertedInlineAtom } from './inserted-atoms';
import { isNewCellPropertyBindings, type NewCellPropertyBinding } from './new-cell-properties';
import type { TargetAstNode } from '../projection/parse-target';
import {
  isInsertedCellFragments,
  type InsertedCellPayloadBinding,
} from './inserted-cell-fragments';

/** Native split limit per dimension (`MAX_CELL_SPLIT_COUNT`). */
const MAX_SPLIT = 63;

/**
 * Split one existing visual cell along existing grid boundaries and fill the new cells. The first
 * resulting cell keeps the native cell, its ID, and its content; executors allocate the others.
 */
export interface SplitTableCellIntent {
  readonly tableId: string;
  /** Existing canonical head cell of the merged visual cell. */
  readonly cellId: string;
  readonly rows: number;
  readonly columns: number;
  /** Row-major block streams for every resulting cell except the first. */
  readonly cells: readonly (readonly TargetAstNode[])[];
  readonly payloads: readonly InsertedCellPayloadBinding[];
  readonly atoms?: readonly InsertedInlineAtom[];
  readonly cellProperties?: readonly NewCellPropertyBinding[];
}

const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const count = (value: unknown): value is number =>
  Number.isSafeInteger(value) && Number(value) >= 1 && Number(value) <= MAX_SPLIT;

/** Structural/data checks run only after the bounded plain-JSON gate. */
export function isSplitCellIntent(value: unknown): value is SplitTableCellIntent {
  return (
    record(value) &&
    Object.keys(value).every((key) =>
      [
        'tableId',
        'cellId',
        'rows',
        'columns',
        'cells',
        'payloads',
        'atoms',
        'cellProperties',
      ].includes(key)
    ) &&
    typeof value.tableId === 'string' &&
    value.tableId.length > 0 &&
    typeof value.cellId === 'string' &&
    value.cellId.length > 0 &&
    count(value.rows) &&
    count(value.columns) &&
    Number(value.rows) * Number(value.columns) > 1 &&
    Array.isArray(value.cells) &&
    value.cells.length === Number(value.rows) * Number(value.columns) - 1 &&
    isInsertedCellFragments(value.cells, value.payloads, true) &&
    isInsertedCellAtoms(
      value.cells as TargetAstNode[][],
      value.payloads as InsertedCellPayloadBinding[],
      value.atoms
    ) &&
    (!Object.hasOwn(value, 'cellProperties') ||
      isNewCellPropertyBindings(value.cells as TargetAstNode[][], value.cellProperties))
  );
}
