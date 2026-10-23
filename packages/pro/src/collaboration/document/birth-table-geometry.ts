/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import type { DocumentRegistry } from './registry.ts';
import type { LogicalId } from './identity.ts';
import { isElementRecord, type ElementRecord } from './schema.ts';
import type { SharedEffect } from './shared-effect.ts';

/** Immutable native topology evidence, never paragraph text or a canonical document snapshot. */
export const NODE_TABLE_BIRTH_GEOMETRY = 'tableBirthGeometry';
export const NODE_TABLE_BIRTH_GEOMETRY_WRITES = 'tableBirthGeometryWrites';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export interface BirthCellCoverage {
  readonly id: LogicalId;
  readonly row: LogicalId;
  readonly columns: readonly LogicalId[];
  readonly span: LogicalId | null;
}
export interface BirthTableCoverage {
  readonly grid: readonly LogicalId[];
  readonly cells: readonly BirthCellCoverage[];
}
export interface BirthTableGeometry {
  readonly table: LogicalId;
  readonly born: readonly LogicalId[];
  readonly before: BirthTableCoverage;
  readonly after: BirthTableCoverage;
}
function element(registry: DocumentRegistry, id: LogicalId): ElementRecord | null {
  const record = registry.record(id);
  return record && isElementRecord(record) ? record : null;
}
function integer(record: ElementRecord | null, fallback: number): number {
  const value = Number(
    record?.attributes.find((a) => a.namespaceUri === W && a.localName === 'val')?.value ?? fallback
  );
  return Number.isSafeInteger(value) && value >= 0 ? value : fallback;
}
interface GeometryView {
  project(record: ElementRecord): ElementRecord;
  shown(parent: LogicalId, child: LogicalId): boolean;
}
export function tableCoverage(
  registry: DocumentRegistry,
  table: LogicalId,
  view?: GeometryView
): BirthTableCoverage {
  const read = (id: LogicalId): ElementRecord | null => {
    const record = element(registry, id);
    return record && view
      ? {
          ...view.project(record),
          childIds: record.childIds.filter((child) => view.shown(id, child)),
        }
      : record;
  };
  const child = (id: LogicalId, name: string) =>
    (read(id)?.childIds ?? [])
      .map(read)
      .find((record) => record?.namespaceUri === W && record.localName === name) ?? null;
  const grid = (child(table, 'tblGrid')?.childIds ?? []).filter(
    (id) => read(id)?.localName === 'gridCol'
  );
  const cells: BirthCellCoverage[] = [];
  for (const rowId of read(table)?.childIds ?? []) {
    if (read(rowId)?.kind !== 'tableRow') continue;
    const trPr = child(rowId, 'trPr');
    let at = integer(trPr ? child(trPr.logicalId, 'gridBefore') : null, 0);
    for (const id of read(rowId)?.childIds ?? []) {
      if (read(id)?.kind !== 'tableCell') continue;
      const tcPr = child(id, 'tcPr');
      const span = tcPr ? child(tcPr.logicalId, 'gridSpan') : null;
      const width = Math.max(1, integer(span, 1));
      cells.push({
        id,
        row: rowId,
        columns: grid.slice(at, at + width),
        span: span?.logicalId ?? null,
      });
      at += width;
    }
  }
  return { grid, cells };
}
function tableOf(registry: DocumentRegistry, id: LogicalId): LogicalId | null {
  const seen = new Set<LogicalId>();
  for (
    let at: LogicalId | null = id;
    at && seen.size < registry.limits.maxTreeDepth;
    at = registry.parentOf(at)
  ) {
    if (seen.has(at)) return null;
    seen.add(at);
    if (registry.kindOf(at) === 'table') return at;
  }
  return null;
}
/** Capture before writing; publish the exact after coverage in that same admitted transaction. */
export function prepareBirthTableGeometry(
  registry: DocumentRegistry,
  effects: readonly SharedEffect[],
  view?: GeometryView
): () => void {
  const born = effects.flatMap((effect) =>
    effect.kind === 'putNode' &&
    effect.descriptor.kind === 'tableCell' &&
    !registry.hasNode(effect.descriptor.logicalId)
      ? [effect.descriptor.logicalId]
      : []
  );
  if (!born.length) return () => {};
  const before = new Map<LogicalId, BirthTableCoverage>();
  for (const effect of effects) {
    const id =
      'parentLogicalId' in effect
        ? effect.parentLogicalId
        : 'logicalId' in effect
          ? effect.logicalId
          : null;
    const table = id ? tableOf(registry, id) : null;
    if (table && !before.has(table)) before.set(table, tableCoverage(registry, table, view));
  }
  return () => {
    for (const [table, coverage] of before) {
      const after = tableCoverage(registry, table);
      // Derived parent indexes update after the transaction. Read the written child lists here.
      const listed = new Set(after.cells.map((cell) => cell.id));
      const members = born.filter((id) => listed.has(id));
      if (!members.length) continue;
      const prior = new Map(coverage.cells.map((cell) => [cell.id, cell]));
      // Row births need ancestry only. Record geometry when an existing cell changes coverage.
      if (
        !after.cells.some((cell) => {
          const old = prior.get(cell.id);
          return (
            old &&
            (old.columns.length !== cell.columns.length ||
              old.columns.some((id, at) => id !== cell.columns[at]))
          );
        })
      )
        continue;
      const frame: BirthTableGeometry = { table, born: members, before: coverage, after };
      const carrier = registry.schema.nodes.get(members[0]!);
      if (carrier && !carrier.has(NODE_TABLE_BIRTH_GEOMETRY))
        carrier.set(NODE_TABLE_BIRTH_GEOMETRY, frame);
    }
  };
}
/** A later native span write owns its value, even if the birth creator subsequently undoes. */
export function prepareBirthGeometryWrites(
  registry: DocumentRegistry,
  effects: readonly SharedEffect[]
): () => void {
  const writes: { span: LogicalId; carrier: LogicalId }[] = [];
  for (const effect of effects) {
    if (
      effect.kind !== 'setAttribute' ||
      effect.qname.namespaceUri !== W ||
      effect.qname.localName !== 'val'
    )
      continue;
    const record = element(registry, effect.logicalId);
    if (record?.namespaceUri !== W || record.localName !== 'gridSpan') continue;
    const table = tableOf(registry, effect.logicalId);
    if (!table) continue;
    for (const cell of tableCoverage(registry, table).cells) {
      const frame = birthTableGeometry(registry, cell.id);
      if (
        frame?.table === table &&
        [...frame.before.cells, ...frame.after.cells].some(
          (coverage) => coverage.span === effect.logicalId
        )
      )
        writes.push({ span: effect.logicalId, carrier: cell.id });
    }
  }
  return () => {
    for (const { span, carrier } of writes) {
      const record = registry.schema.nodes.get(span);
      if (!record) continue;
      let owners = record.get(NODE_TABLE_BIRTH_GEOMETRY_WRITES);
      if (!(owners instanceof Y.Map)) {
        owners = new Y.Map<boolean>();
        record.set(NODE_TABLE_BIRTH_GEOMETRY_WRITES, owners);
      }
      (owners as Y.Map<boolean>).set(carrier, true);
    }
  };
}
export function hasBirthGeometryWrite(
  registry: DocumentRegistry,
  span: LogicalId,
  carrier: LogicalId
): boolean {
  const owners = registry.schema.nodes.get(span)?.get(NODE_TABLE_BIRTH_GEOMETRY_WRITES);
  return owners instanceof Y.Map && owners.get(carrier) === true;
}
/** A shared field is untrusted. Bound every collection before using IDs to read native records. */
export function birthTableGeometry(
  registry: DocumentRegistry,
  carrier: LogicalId
): BirthTableGeometry | null {
  const source = registry.schema.nodes.get(carrier)?.get(NODE_TABLE_BIRTH_GEOMETRY);
  if (!source || typeof source !== 'object' || Array.isArray(source)) return null;
  try {
    const frame = source as BirthTableGeometry;
    const id = (value: unknown): value is LogicalId =>
      typeof value === 'string' &&
      value.length > 0 &&
      value.length <= registry.limits.maxStringLength &&
      !['__proto__', 'constructor', 'prototype'].includes(value);
    const ids = (value: unknown): value is LogicalId[] =>
      Array.isArray(value) && value.length <= registry.limits.maxNodes && value.every(id);
    const coverage = (value: BirthTableCoverage): boolean =>
      !!value &&
      ids(value.grid) &&
      Array.isArray(value.cells) &&
      value.cells.length <= registry.limits.maxNodes &&
      value.cells.every(
        (cell) =>
          !!cell &&
          id(cell.id) &&
          id(cell.row) &&
          ids(cell.columns) &&
          (cell.span === null || id(cell.span))
      );
    if (
      !id(frame.table) ||
      !ids(frame.born) ||
      !frame.born.includes(carrier) ||
      !coverage(frame.before) ||
      !coverage(frame.after)
    )
      return null;
    return frame;
  } catch {
    return null;
  }
}
