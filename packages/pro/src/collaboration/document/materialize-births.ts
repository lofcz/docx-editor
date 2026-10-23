/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type { DocumentRegistry } from './registry.ts';
import type { LogicalId } from './identity.ts';
import { nodeRecordUndone, isElementRecord, isTextRecord, type ElementRecord } from './schema.ts';
import { BIRTH_GEOMETRY_NAMES } from './birth-undo-dependencies.ts';
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
}
/** One immutable materialization pass. Retained shells carry foreign content, never original text. */
export class BirthProjection {
  private readonly content = new Map<LogicalId, boolean>();
  private readonly contexts = new Map<LogicalId, TableContext>();
  private readonly visibility = new Map<LogicalId, boolean>();
  private readonly columnVisibility = new Map<LogicalId, string>();
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
      const grid = record.childIds
        .map((child) => this.element(child))
        .find((child) => child?.namespaceUri === W && child.localName === 'tblGrid');
      if (
        !grid ||
        (!this.columnVisibility.has(id) &&
          !grid.childIds.some((child) => undoneBirth(this.registry, child)))
      )
        continue;
      const context = this.table(id);
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
      if (this.element(at)?.kind === 'table') return at;
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
    for (const rowId of table.childIds) {
      const row = this.element(rowId);
      if (row?.kind !== 'tableRow') continue;
      let column = Number(this.property(rowId, 'trPr', 'gridBefore') ?? 0);
      const visible = !undoneBirth(this.registry, rowId) || this.hasContent(rowId);
      for (const cellId of row.childIds) {
        const cell = this.element(cellId);
        if (cell?.kind !== 'tableCell') continue;
        const span = Math.max(1, Number(this.property(cellId, 'tcPr', 'gridSpan') ?? 1));
        cells.set(cellId, [column, column + span]);
        if (visible && undoneBirth(this.registry, cellId) && this.hasContent(cellId))
          for (let c = column; c < column + span; c++) columns.add(c);
        column += span;
      }
    }
    const context = { columns, cells, grid };
    this.contexts.set(id, context);
    return context;
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
    if (!undoneBirth(this.registry, child)) return true;
    const record = this.element(child);
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
