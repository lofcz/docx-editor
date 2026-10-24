/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type { DocumentRegistry } from './registry.ts';
import type { LogicalId } from './identity.ts';
import { nodeRecordUndone, isElementRecord, isTextRecord, type ElementRecord } from './schema.ts';
import { BIRTH_GEOMETRY_NAMES } from './birth-undo-dependencies.ts';
import { birthTableGeometries, hasBirthGeometryWrite } from './birth-table-geometry.ts';
import { BirthCoverageFrontier } from './birth-coverage-frontier.ts';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
/** An undo-owned marker on a retained main-document birth. Redo removes the marker. */
export function undoneBirth(registry: DocumentRegistry, id: LogicalId): boolean {
  if (!nodeRecordUndone(registry.schema.nodes.get(id))) return false;
  const root = registry
    .partEntries()
    .find((part) => part.name === registry.mainDocumentPart())?.rootLogicalId;
  const seen = new Set<LogicalId>();
  for (
    let at: LogicalId | null = id;
    at && seen.size < registry.limits.maxTreeDepth;
    at = registry.parentOf(at)
  ) {
    if (at === root) return true;
    if (seen.has(at)) return false;
    seen.add(at);
  }
  return false;
}
interface TableContext {
  readonly columns: ReadonlySet<number>;
  readonly cells: ReadonlyMap<LogicalId, readonly [number, number]>;
  readonly grid: readonly LogicalId[];
  readonly spans: ReadonlyMap<LogicalId, number>;
  readonly bornCells: ReadonlyMap<LogicalId, boolean>;
  readonly widths: ReadonlyMap<LogicalId, number>;
}
/** One immutable materialization pass. Retained shells carry foreign content, never original text. */
export class BirthProjection {
  private readonly content = new Map<LogicalId, boolean>();
  private readonly contexts = new Map<LogicalId, TableContext>();
  private readonly visibility = new Map<LogicalId, boolean>();
  private readonly columnVisibility = new Map<LogicalId, string>();
  private readonly spanProjection = new Map<LogicalId, ReadonlyMap<LogicalId, number>>();
  constructor(private readonly registry: DocumentRegistry) {}
  /** Visibility can depend on text in another row. Invalidate every affected native parent. */
  prepare(changed: ReadonlySet<LogicalId>): readonly LogicalId[] {
    this.content.clear();
    this.contexts.clear();
    const touched = new Set<LogicalId>();
    for (const id of changed) {
      const seen = new Set<LogicalId>();
      for (
        let at: LogicalId | null = id;
        at && seen.size < this.registry.limits.maxTreeDepth;
        at = this.registry.parentOf(at)
      ) {
        if (seen.has(at)) break;
        seen.add(at);
        touched.add(at);
      }
    }
    const dirty = new Set<LogicalId>();
    for (const id of touched) {
      const parent = this.registry.parentOf(id);
      if (!parent) continue;
      const shown = this.shown(parent, id);
      if (shown !== (this.visibility.get(id) ?? true)) dirty.add(parent);
      if (!shown) this.visibility.set(id, false);
      else this.visibility.delete(id);
    }
    for (const id of touched) {
      const record = this.element(id);
      if (record?.kind !== 'table') continue;
      const context = this.table(id);
      const geometry = new Map([...context.spans, ...context.widths]);
      // A surviving split can change spans without changing the visible column mask.
      // Rebuild the exact native property paths whose projected geometry depends on peer content.
      const previousSpans = this.spanProjection.get(id) ?? new Map<LogicalId, number>();
      for (const span of new Set([...previousSpans.keys(), ...geometry.keys()])) {
        if (previousSpans.get(span) !== geometry.get(span)) {
          dirty.add(id);
          const seen = new Set<LogicalId>();
          for (
            let at: LogicalId | null = span;
            at && at !== id && seen.size < this.registry.limits.maxTreeDepth;
            at = this.registry.parentOf(at)
          ) {
            if (seen.has(at)) break;
            seen.add(at);
            dirty.add(at);
          }
        }
      }
      if (geometry.size) this.spanProjection.set(id, geometry);
      else this.spanProjection.delete(id);
      const grid = record.childIds
        .map((child) => this.element(child))
        .find((child) => child?.namespaceUri === W && child.localName === 'tblGrid');
      if (
        !grid ||
        (!this.columnVisibility.has(id) &&
          !grid.childIds.some((child) => undoneBirth(this.registry, child)))
      )
        continue;
      const next = [...context.columns].sort((a, b) => a - b).join('/');
      if (this.columnVisibility.get(id) !== next) {
        dirty.add(id);
        dirty.add(grid.logicalId);
        for (const row of record.childIds)
          if (this.element(row)?.kind === 'tableRow') dirty.add(row);
      }
      this.columnVisibility.set(id, next);
    }
    return [...dirty];
  }
  private element(id: LogicalId): ElementRecord | null {
    const record = this.registry.record(id);
    return record && isElementRecord(record) ? record : null;
  }
  private hasContent(id: LogicalId, path = new Set<LogicalId>()): boolean {
    const cached = this.content.get(id);
    if (cached !== undefined) return cached;
    if (
      path.has(id) ||
      path.size >= this.registry.limits.maxTreeDepth ||
      this.registry.isTombstoned(id)
    )
      return false;
    path.add(id);
    const record = this.registry.record(id);
    const liveShell = !!record && !undoneBirth(this.registry, id);
    const text = this.registry.inline.textOf(id);
    const visible =
      liveShell ||
      !!text?.length ||
      (!!record &&
        (isTextRecord(record)
          ? record.value.length > 0
          : record.attributes.some(
              (attr) =>
                !(attr.namespaceUri === W14 && ['paraId', 'textId'].includes(attr.localName)) &&
                !(record.namespaceUri === W && BIRTH_GEOMETRY_NAMES.has(record.localName))
            ) || record.childIds.some((child) => this.hasContent(child, path))));
    path.delete(id);
    this.content.set(id, visible);
    return visible;
  }
  private property(id: LogicalId, container: string, property: string): string | undefined {
    const record = this.element(id);
    const props = record?.childIds
      .map((child) => this.element(child))
      .find((child) => child?.namespaceUri === W && child.localName === container);
    const value = props?.childIds
      .map((child) => this.element(child))
      .find((child) => child?.namespaceUri === W && child.localName === property);
    return value?.attributes.find((attr) => attr.namespaceUri === W && attr.localName === 'val')
      ?.value;
  }
  private tableOf(id: LogicalId): LogicalId | null {
    const seen = new Set<LogicalId>();
    for (
      let at: LogicalId | null = id;
      at && seen.size < this.registry.limits.maxTreeDepth;
      at = this.registry.parentOf(at)
    ) {
      if (seen.has(at)) return null;
      seen.add(at);
      if (this.registry.kindOf(at) === 'table') return at;
    }
    return null;
  }
  private table(id: LogicalId): TableContext {
    const cached = this.contexts.get(id);
    if (cached) return cached;
    const table = this.element(id)!;
    const gridNode = table.childIds
      .map((child) => this.element(child))
      .find((child) => child?.namespaceUri === W && child.localName === 'tblGrid');
    const grid =
      gridNode?.childIds.filter((child) => this.element(child)?.localName === 'gridCol') ?? [];
    const columns = new Set(
      grid.flatMap((child, index) => (undoneBirth(this.registry, child) ? [] : [index]))
    );
    const cells = new Map<LogicalId, readonly [number, number]>();
    const rawCells: LogicalId[] = [];
    const visibleCells = new Set<LogicalId>();
    for (const rowId of table.childIds) {
      const row = this.element(rowId);
      if (row?.kind !== 'tableRow') continue;
      let column = Number(this.property(rowId, 'trPr', 'gridBefore') ?? 0);
      const visible = !undoneBirth(this.registry, rowId) || this.hasContent(rowId);
      for (const cellId of row.childIds) {
        const cell = this.element(cellId);
        if (cell?.kind !== 'tableCell') continue;
        rawCells.push(cellId);
        if (visible) visibleCells.add(cellId);
        const span = Math.max(1, Number(this.property(cellId, 'tcPr', 'gridSpan') ?? 1));
        cells.set(cellId, [column, column + span]);
        column += span;
      }
    }
    const spans = new Map<LogicalId, number>();
    const widths = new Map<LogicalId, number>();
    const bornCells = new Map<LogicalId, boolean>();
    const frames = birthTableGeometries(this.registry, id);
    const rawCellSet = new Set(rawCells);
    const gridIndexes = new Map(grid.map((column, at) => [column, at]));
    const birthCoverage = new Map<LogicalId, readonly LogicalId[]>();
    for (const [, frame] of frames) {
      const after = new Map(frame.after.cells.map((cell) => [cell.id, cell]));
      for (const born of frame.born) {
        const cell = after.get(born);
        if (cell && !birthCoverage.has(born)) birthCoverage.set(born, cell.columns);
      }
    }
    // Content activates the column identities owned by its birth, not a scalar span
    // restored by a different history step or a later grid refinement.
    for (const cell of rawCells) {
      if (!visibleCells.has(cell) || !undoneBirth(this.registry, cell) || !this.hasContent(cell))
        continue;
      const range = cells.get(cell)!;
      for (const column of birthCoverage.get(cell) ?? grid.slice(range[0], range[1])) {
        const index = gridIndexes.get(column);
        if (index !== undefined) columns.add(index);
      }
    }
    const active = new Set([...columns].map((column) => grid[column]!));
    const ownedByRow = new Map<LogicalId, Map<LogicalId, readonly LogicalId[]>>();
    for (const [, frame] of frames) {
      const after = new Map(frame.after.cells.map((cell) => [cell.id, cell]));
      // Recorded column identities, rather than restored scalar spans, locate retained births.
      for (const cellId of rawCells) {
        const coverage = after.get(cellId);
        if (!coverage?.columns.length) continue;
        const indexes = coverage.columns
          .map((column) => grid.indexOf(column))
          .filter((column) => column >= 0);
        if (indexes.length) cells.set(cellId, [Math.min(...indexes), Math.max(...indexes) + 1]);
      }
      for (const cellId of frame.born) {
        const coverage = after.get(cellId);
        if (!coverage || !rawCellSet.has(cellId)) continue;
        const needed =
          this.hasContent(cellId) ||
          coverage.columns.some(
            (column) =>
              active.has(column) &&
              !frame.before.cells.some(
                (old) => old.row === coverage.row && old.columns.includes(column)
              )
          );
        bornCells.set(cellId, needed);
        if (needed) {
          const owned = ownedByRow.get(coverage.row) ?? new Map<LogicalId, readonly LogicalId[]>();
          owned.set(cellId, coverage.columns);
          ownedByRow.set(coverage.row, owned);
        }
      }
    }
    const frontier = new BirthCoverageFrontier(frames);
    const spanTargets = new Map<LogicalId, LogicalId>();
    for (const [carrier, frame] of frames) {
      const after = new Map(frame.after.cells.map((cell) => [cell.id, cell]));
      const originalGrid = new Set(frame.before.grid);
      const restoring = frame.born.some((cell) => undoneBirth(this.registry, cell));
      for (const old of frame.before.cells) {
        const current = after.get(old.id);
        if (!current || !rawCellSet.has(old.id) || this.registry.isTombstoned(old.id)) continue;
        const covered = new Set(
          (restoring ? old.columns : current.columns).filter((column) => active.has(column))
        );
        for (const column of current.columns)
          if (!originalGrid.has(column) && active.has(column)) covered.add(column);
        for (const [cell, owned] of ownedByRow.get(old.row) ?? []) {
          if (cell !== old.id) for (const column of owned) covered.delete(column);
        }
        const span = old.span ?? current.span;
        if (span && covered.size && !hasBirthGeometryWrite(this.registry, span, carrier)) {
          frontier.add(old.id, carrier, covered);
          spanTargets.set(old.id, span);
        }
      }
    }
    for (const [cell, span] of spanTargets) {
      const coverage = frontier.coverage(cell);
      const size = coverage.size;
      const raw = this.element(span)?.attributes.find(
        (attribute) => attribute.namespaceUri === W && attribute.localName === 'val'
      )?.value;
      if (size && size !== Number(raw ?? 1)) spans.set(span, size);
      if (frontier.concurrent(cell)) {
        const props = this.element(cell)
          ?.childIds.map((id) => this.element(id))
          .find((record) => record?.namespaceUri === W && record.localName === 'tcPr');
        const total = [...coverage].reduce(
          (sum, id) =>
            sum +
            Number(
              this.element(id)?.attributes.find((a) => a.namespaceUri === W && a.localName === 'w')
                ?.value ?? NaN
            ),
          0
        );
        if (props && Number.isFinite(total))
          for (const child of props.childIds) {
            const record = this.element(child);
            if (
              record?.namespaceUri === W &&
              record.localName === 'tcW' &&
              record.attributes.some((a) => a.localName === 'type' && a.value === 'dxa')
            )
              widths.set(child, total);
          }
      }
    }
    const context = { columns, cells, grid, spans, bornCells, widths };
    this.contexts.set(id, context);
    return context;
  }
  /** Derive native span attributes without writing shared state or replacing content nodes. */
  project(record: ElementRecord): ElementRecord {
    if (record.namespaceUri !== W || !['gridSpan', 'tcW'].includes(record.localName)) return record;
    const table = this.tableOf(record.logicalId);
    const span = table
      ? (record.localName === 'tcW' ? this.table(table).widths : this.table(table).spans).get(
          record.logicalId
        )
      : undefined;
    if (span === undefined) return record;
    const attributes = record.attributes.map((attribute) =>
      attribute.namespaceUri === W &&
      attribute.localName === (record.localName === 'tcW' ? 'w' : 'val')
        ? { ...attribute, value: String(span) }
        : attribute
    );
    const name = record.localName === 'tcW' ? 'w' : 'val';
    if (
      !attributes.some((attribute) => attribute.namespaceUri === W && attribute.localName === name)
    )
      attributes.push({
        namespaceUri: W,
        localName: name,
        prefix: record.prefix ?? 'w',
        value: String(span),
      });
    return { ...record, attributes };
  }
  private geometry(id: LogicalId, seen = new Set<LogicalId>()): boolean {
    if (seen.has(id) || seen.size >= this.registry.limits.maxTreeDepth) return false;
    seen.add(id);
    const record = this.element(id);
    return (
      !!record &&
      ((record.namespaceUri === W && BIRTH_GEOMETRY_NAMES.has(record.localName)) ||
        record.childIds.some((child) => this.geometry(child, seen)))
    );
  }
  shown(parent: LogicalId, child: LogicalId): boolean {
    // Projected journals can omit shared record maps. Ordinary live nodes need only
    // shape reads; retained undone paragraphs also use the foreign-content projection.
    if (
      this.registry.kindOf(child) !== 'table' &&
      this.tableOf(parent) === null &&
      !nodeRecordUndone(this.registry.schema?.nodes.get(child))
    )
      return true;
    const record = this.element(child);
    const ownerProps =
      record?.namespaceUri === W && record.localName === 'tcW' ? this.element(parent) : null;
    if (ownerProps?.namespaceUri === W && ownerProps.localName === 'tcPr') {
      const copies = ownerProps.childIds.filter((id) => {
        const sibling = this.element(id);
        return (
          sibling?.namespaceUri === W &&
          sibling.localName === 'tcW' &&
          !this.registry.isTombstoned(id)
        );
      });
      if (copies.some((id) => id.startsWith('lid:')) && copies[0] !== child) return false;
    }
    if (!undoneBirth(this.registry, child)) return true;
    if (!record) return this.hasContent(child);
    if (record.kind === 'tableRow') return this.hasContent(child);
    if (
      record.kind === 'tableCell' ||
      (record.namespaceUri === W && record.localName === 'gridCol')
    ) {
      const tableId = this.tableOf(parent);
      if (tableId) {
        const context = this.table(tableId);
        if (record.kind === 'tableCell') {
          const needed = context.bornCells.get(child);
          if (needed !== undefined) return needed;
          const range = context.cells.get(child);
          return (
            !!range &&
            [...context.columns].some((column) => column >= range[0] && column < range[1])
          );
        }
        return context.columns.has(context.grid.indexOf(child));
      }
    }
    const owner = this.element(parent);
    if (
      record.kind === 'paragraph' &&
      owner?.kind === 'tableCell' &&
      owner.childIds.find((id) => this.element(id)?.kind === 'paragraph') === child
    )
      return true;
    return (
      this.hasContent(child) ||
      (record.namespaceUri === W &&
        [
          'tcPr',
          'tblGrid',
          'tblPr',
          'trPr',
          'gridSpan',
          'vMerge',
          'tcW',
          'gridCol',
          'tblW',
          'gridBefore',
          'gridAfter',
        ].includes(record.localName) &&
        this.geometry(child))
    );
  }
}
