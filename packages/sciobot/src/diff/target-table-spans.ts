import { EDIT_PLAN_LIMITS } from './plan';
import type { NativeEditFailure } from '../contracts/document';
import type { JsonObject } from '../data/json';
import type { BoundSourceBlock } from './bound-source-types';
import type { LogicalBlock, LogicalGroup } from './logical-block-types';
import { staticPropertyValue, TargetPropertyError } from './target-property-values';

/** Word tables cannot exceed 63 grid columns; rows share the same bound for one request. */
const MAX_SPAN = 63;

const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});

export interface TargetSlotRegion {
  readonly row: number;
  readonly column: number;
  readonly rows: number;
  readonly columns: number;
}
export interface ExpandedTargetTable {
  readonly table: LogicalGroup;
  /** Anchor region for every slot of the expanded regular grid. */
  readonly regions: readonly (readonly TargetSlotRegion[])[];
  /** Slots that only an explicit `colspan`/`rowspan` covers. They never carry content. */
  readonly covered: ReadonlySet<string>;
}

const slot = (row: number, column: number) => `${row}/${column}`;

function spanValue(raw: unknown): number {
  const text = staticPropertyValue(raw);
  if (!/^\s*\d+\s*$/.test(text)) throw new TargetPropertyError('Table spans need whole numbers.');
  const value = Number(text);
  if (value < 1 || value > MAX_SPAN) throw new TargetPropertyError('Table spans exceed 63.');
  return value;
}

/**
 * Expand explicit HTML `colspan`/`rowspan` into a regular slot grid. Spans become geometry data;
 * covered slots receive empty cells so retained metadata and content comparisons stay positional.
 * Source and target projections can both use explicit spans. Absence carries no meaning.
 */
export function expandTargetSpans(
  table: LogicalGroup
): { readonly success: true; readonly expanded: ExpandedTargetTable } | NativeEditFailure {
  const sourceRows = table.children as readonly LogicalGroup[];
  if (sourceRows.some((row) => row.kind !== 'row' || !('children' in row)))
    return {
      success: true,
      expanded: { table, regions: [], covered: new Set() },
    };
  try {
    const grid: (TargetSlotRegion | undefined)[][] = sourceRows.map(() => []),
      cells: (LogicalGroup | undefined)[][] = sourceRows.map(() => []),
      covered = new Set<string>();
    let explicit = false,
      slots = 0;
    sourceRows.forEach((row, rowIndex) => {
      let column = 0;
      for (const block of row.children) {
        if (block.kind !== 'cell' || !('children' in block))
          throw new TargetPropertyError('Table rows contain only cells.');
        while (grid[rowIndex]![column]) column++;
        const header: Record<string, JsonObject[string]> = { ...block.header };
        let columns = 1,
          rows = 1;
        if (Array.isArray(header.attributes)) {
          const kept: JsonObject[string][] = [];
          for (const raw of header.attributes) {
            const attribute = raw as JsonObject;
            const name =
              attribute?.type === 'mdxJsxAttribute' && typeof attribute.name === 'string'
                ? attribute.name.toLowerCase()
                : '';
            if (name === 'colspan' || name === 'rowspan') {
              const value = spanValue(attribute.value);
              if (name === 'colspan') columns = value;
              else rows = value;
              explicit = true;
              continue;
            }
            kept.push(raw);
          }
          header.attributes = kept;
        }
        if (rowIndex + rows > sourceRows.length)
          throw new TargetPropertyError('A table row span crosses the end of its table.');
        if (column + columns > MAX_SPAN)
          throw new TargetPropertyError('A table column span exceeds the native grid.');
        const region = { row: rowIndex, column, rows, columns };
        for (let r = rowIndex; r < rowIndex + rows; r++)
          for (let c = column; c < column + columns; c++) {
            if (grid[r]![c]) throw new TargetPropertyError('Table spans overlap.');
            if (++slots > EDIT_PLAN_LIMITS.maxNodes)
              throw new TargetPropertyError('Expanded table spans exceed their node limit.');
            grid[r]![c] = region;
            if (r !== rowIndex || c !== column) covered.add(slot(r, c));
          }
        cells[rowIndex]![column] = { ...block, header };
        column += columns;
      }
    });
    if (!explicit)
      return {
        success: true,
        expanded: {
          table,
          regions: sourceRows.map((row, r) =>
            row.children.map((_, c) => ({ row: r, column: c, rows: 1, columns: 1 }))
          ),
          covered,
        },
      };
    let width = 0;
    for (const row of grid) if (row.length > width) width = row.length;
    if (grid.some((row) => row.length !== width || row.some((entry) => !entry)))
      return fail('Irregular table slot counts need explicit native grid intentions.');
    const rows = sourceRows.map((row, r) => ({
      ...row,
      children: Array.from({ length: width }, (_, c): LogicalBlock => {
        const own = cells[r]![c];
        if (own) return own;
        const anchor = cells[grid[r]![c]!.row]![grid[r]![c]!.column]!;
        return {
          kind: 'cell',
          path: anchor.path,
          list: null,
          key: 'covered',
          header: { type: 'tableCell' },
          domainType: null,
          children: [],
          rowGroups: [],
        };
      }),
    }));
    return {
      success: true,
      expanded: {
        table: { ...table, children: rows },
        regions: grid as TargetSlotRegion[][],
        covered,
      },
    };
  } catch (error) {
    if (error instanceof TargetPropertyError) return fail(error.message);
    throw error;
  }
}

export interface SourceSlotGrid {
  /** Visual cell head slot for every source slot. */
  readonly heads: readonly (readonly { readonly row: number; readonly column: number }[])[];
  /** Head cell ID by `row/column` of the head slot. */
  readonly cellIds: ReadonlyMap<string, string>;
  /** Complete rectangle of each visual cell, keyed by its head slot. */
  readonly regions: ReadonlyMap<string, TargetSlotRegion>;
}

/** Read native merge geometry from D22 bindings. Placeholders resolve to their visual head. */
export function sourceSlotGrid(
  bindings: ReadonlyMap<string, BoundSourceBlock>,
  path: readonly number[],
  rows: number,
  columns: number
): SourceSlotGrid | null {
  const heads: { row: number; column: number }[][] = [],
    cellIds = new Map<string, string>(),
    regions = new Map<string, TargetSlotRegion>();
  for (let r = 0; r < rows; r++) {
    heads.push([]);
    for (let c = 0; c < columns; c++) {
      const owner = bindings.get([...path, r, c].join('/'))?.table;
      if (!owner?.cellId) return null;
      let head: { row: number; column: number } | undefined;
      if (owner.role === 'cell') head = { row: r, column: c };
      else if (owner.role === 'horizontal-placeholder' && c > 0) head = heads[r]![c - 1];
      else if (owner.role === 'vertical-placeholder' && r > 0) head = heads[r - 1]![c];
      if (!head) return null;
      heads[r]!.push(head);
      const key = slot(head.row, head.column);
      if (owner.role === 'cell') cellIds.set(key, owner.cellId);
      const region = regions.get(key);
      regions.set(
        key,
        region
          ? {
              row: region.row,
              column: region.column,
              rows: Math.max(region.rows, r - region.row + 1),
              columns: Math.max(region.columns, c - region.column + 1),
            }
          : { row: r, column: c, rows: 1, columns: 1 }
      );
    }
  }
  // A visual cell must cover exactly its bounding rectangle.
  for (const [key, region] of regions)
    for (let r = region.row; r < region.row + region.rows; r++)
      for (let c = region.column; c < region.column + region.columns; c++)
        if (slot(heads[r]![c]!.row, heads[r]![c]!.column) !== key) return null;
  return { heads, cellIds, regions };
}

export interface TargetMergeRequest {
  readonly kind: 'merge';
  readonly region: TargetSlotRegion;
  /** Source visual heads inside the region in native reading order; the first is the new head. */
  readonly heads: readonly TargetSlotRegion[];
}
export interface TargetSplitRequest {
  readonly kind: 'split';
  /** The merged source visual cell. */
  readonly region: TargetSlotRegion;
  readonly rows: number;
  readonly columns: number;
  /** Target sub-cell anchors in row-major order; the first keeps the native cell. */
  readonly parts: readonly TargetSlotRegion[];
}
export type TargetGeometryRequest = TargetMergeRequest | TargetSplitRequest;

const empty = (block: LogicalBlock | undefined) =>
  !!block && 'children' in block && block.children.length === 0;

/**
 * Compare target and native partitions slot by slot. An empty target slot inside a native merge
 * keeps that merge; content or an explicit span inside it requests a split. A target region that
 * covers several complete native cells requests a merge. Crossing regions refuse.
 */
export function targetGeometryRequests(
  source: SourceSlotGrid,
  target: ExpandedTargetTable
):
  | { readonly success: true; readonly requests: readonly TargetGeometryRequest[] }
  | NativeEditFailure {
  const rows = target.table.children as readonly LogicalGroup[],
    requests: TargetGeometryRequest[] = [],
    explicit = (region: TargetSlotRegion) => region.rows > 1 || region.columns > 1;
  // Native merged cells: retained, or split by target content/spans.
  for (const [key, region] of source.regions) {
    if (region.rows === 1 && region.columns === 1) continue;
    const parts = new Map<string, TargetSlotRegion>();
    let requested = false,
      larger = 0;
    for (let r = region.row; r < region.row + region.rows; r++)
      for (let c = region.column; c < region.column + region.columns; c++) {
        const own = target.regions[r]![c]!;
        if (
          own.row < region.row ||
          own.column < region.column ||
          own.row + own.rows > region.row + region.rows ||
          own.column + own.columns > region.column + region.columns
        ) {
          // A target region larger than this native cell merges it with neighbors (below).
          larger++;
          continue;
        }
        parts.set(slot(own.row, own.column), own);
        const head = r === region.row && c === region.column;
        if (!head && !target.covered.has(slot(r, c)) && !empty(rows[r]!.children[c]))
          requested = true;
        if (explicit(own) && (own.rows !== region.rows || own.columns !== region.columns))
          requested = true;
      }
    if (larger === region.rows * region.columns) continue;
    if (larger) return fail('A target region crosses a native merged cell.');
    if (!requested) continue;
    const list = [...parts.values()].sort((a, b) => a.row - b.row || a.column - b.column);
    if (list.some((part) => part.rows * part.columns !== list[0]!.rows * list[0]!.columns))
      return fail('A native merged cell can split only into equal target cells.');
    const height = list[0]!.rows,
      width = list[0]!.columns;
    if (
      list.length !== (region.rows / height) * (region.columns / width) ||
      region.rows % height ||
      region.columns % width ||
      list.some(
        (part) =>
          part.rows !== height ||
          part.columns !== width ||
          (part.row - region.row) % height ||
          (part.column - region.column) % width
      )
    )
      return fail('A native merged cell can split only into a regular grid of target cells.');
    if (list.length === 1) continue;
    requests.push({
      kind: 'split',
      region: source.regions.get(key)!,
      rows: region.rows / height,
      columns: region.columns / width,
      parts: list,
    });
  }
  // Target regions covering several native visual cells.
  const seen = new Set<string>();
  rows.forEach((_, r) =>
    target.regions[r]!.forEach((own) => {
      const key = slot(own.row, own.column);
      if (seen.has(key) || !explicit(own)) return;
      seen.add(key);
      const heads = new Map<string, TargetSlotRegion>();
      for (let rr = own.row; rr < own.row + own.rows; rr++)
        for (let cc = own.column; cc < own.column + own.columns; cc++) {
          const head = source.heads[rr]![cc]!;
          heads.set(slot(head.row, head.column), source.regions.get(slot(head.row, head.column))!);
        }
      if (heads.size < 2) return;
      requests.push({
        kind: 'merge',
        region: own,
        heads: [...heads.values()].sort((a, b) => a.row - b.row || a.column - b.column),
      });
    })
  );
  for (const request of requests) {
    if (request.kind !== 'merge') continue;
    for (const head of request.heads)
      if (
        head.row < request.region.row ||
        head.column < request.region.column ||
        head.row + head.rows > request.region.row + request.region.rows ||
        head.column + head.columns > request.region.column + request.region.columns
      )
        return fail('A target merge must cover complete native cells.');
  }
  // A split region that a merge also covers would need two conflicting native operations.
  for (const split of requests)
    if (
      split.kind === 'split' &&
      requests.some(
        (merge) =>
          merge.kind === 'merge' &&
          merge.heads.some(
            (head) => head.row === split.region.row && head.column === split.region.column
          )
      )
    )
      return fail('A target cannot merge and split the same native cell.');
  return { success: true, requests };
}

export const targetSlotKey = slot;
