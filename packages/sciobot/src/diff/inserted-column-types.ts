import type { TargetAstNode } from '../projection/parse-target';
import {
  isInsertedCellFragments,
  type InsertedCellPayloadBinding,
} from './inserted-cell-fragments';

export type InsertedColumnAnchor =
  | { readonly kind: 'grid-column'; readonly id: string }
  | { readonly kind: 'reference-cell'; readonly id: string };
export interface InsertedColumnCell {
  readonly rowId: string;
  /** Empty when insertion widens an existing span or creates a hidden continuation. */
  readonly blocks: readonly TargetAstNode[];
}
export interface InsertedColumnProperties {
  /** Native explicit grid sizing, in twips. Omission retains native inheritance. */
  readonly widthTwips?: number;
}
export type InsertedColumnPayloadBinding = InsertedCellPayloadBinding;
export interface InsertedTableColumnIntent {
  readonly tableId: string;
  readonly anchor: InsertedColumnAnchor;
  readonly where: 'left' | 'right';
  /** One stream per original canonical row, addressed by row ID instead of preview cell IDs. */
  readonly cells: readonly InsertedColumnCell[];
  readonly properties: InsertedColumnProperties;
  readonly payloads: readonly InsertedColumnPayloadBinding[];
}
const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const only = (value: object, keys: readonly string[]) =>
  Object.keys(value).every((key) => keys.includes(key));
const id = (value: unknown): value is string => typeof value === 'string' && value.length > 0;

/** Structural validation follows the bounded plain-JSON gate. Native geometry stays authoritative. */
export function isInsertedColumnIntent(value: unknown): value is InsertedTableColumnIntent {
  if (
    !record(value) ||
    !only(value, ['tableId', 'anchor', 'where', 'cells', 'properties', 'payloads']) ||
    !id(value.tableId) ||
    !record(value.anchor) ||
    !only(value.anchor, ['kind', 'id']) ||
    (value.anchor.kind !== 'grid-column' && value.anchor.kind !== 'reference-cell') ||
    !id(value.anchor.id) ||
    (value.where !== 'left' && value.where !== 'right') ||
    !Array.isArray(value.cells) ||
    !value.cells.length ||
    !record(value.properties) ||
    !only(value.properties, ['widthTwips']) ||
    (Object.hasOwn(value.properties, 'widthTwips') &&
      (!Number.isSafeInteger(value.properties.widthTwips) ||
        Number(value.properties.widthTwips) < 300 ||
        Number(value.properties.widthTwips) > 31680))
  )
    return false;
  const rowIds = new Set<string>(),
    blocks: unknown[] = [];
  for (const cell of value.cells) {
    if (
      !record(cell) ||
      !only(cell, ['rowId', 'blocks']) ||
      !id(cell.rowId) ||
      rowIds.has(cell.rowId)
    )
      return false;
    rowIds.add(cell.rowId);
    blocks.push(cell.blocks);
  }
  return isInsertedCellFragments(blocks, value.payloads);
}
