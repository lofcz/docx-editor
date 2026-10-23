/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type { LogicalId } from './identity.ts';
import type { BirthTableGeometry } from './birth-table-geometry.ts';
export type BirthGeometryFrame = readonly [LogicalId, BirthTableGeometry];

/** Native cell and column identities encode which earlier births a writer observed. */
function dependencies(
  frames: readonly BirthGeometryFrame[]
): ReadonlyMap<LogicalId, ReadonlySet<LogicalId>> {
  const owners = new Map<LogicalId, LogicalId>();
  for (const [carrier, frame] of frames) {
    for (const born of frame.born) if (!owners.has(born)) owners.set(born, carrier);
    const before = new Set(frame.before.grid);
    for (const column of frame.after.grid)
      if (!before.has(column) && !owners.has(column)) owners.set(column, carrier);
  }
  const result = new Map<LogicalId, ReadonlySet<LogicalId>>();
  for (const [carrier, frame] of frames) {
    const prior = new Set<LogicalId>();
    for (const id of [...frame.before.cells.map((cell) => cell.id), ...frame.before.grid]) {
      const owner = owners.get(id);
      if (owner && owner !== carrier) prior.add(owner);
    }
    result.set(carrier, prior);
  }
  return result;
}
/** Stable causal order, independent of the order shared map entries reached a replica. */
export function orderBirthGeometryFrames(
  input: readonly BirthGeometryFrame[]
): readonly BirthGeometryFrame[] {
  const frames = [...input].sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
  const prior = dependencies(frames);
  const counts = new Map<LogicalId, number>();
  const followers = new Map<LogicalId, LogicalId[]>();
  const byId = new Map(frames);
  const ready: LogicalId[] = [];
  for (const [carrier] of frames) {
    const parents = prior.get(carrier)!;
    counts.set(carrier, parents.size);
    if (!parents.size) ready.push(carrier);
    for (const id of parents) {
      const next = followers.get(id) ?? [];
      next.push(carrier);
      followers.set(id, next);
    }
  }
  const ordered: BirthGeometryFrame[] = [];
  for (let at = 0; at < ready.length; at++) {
    const carrier = ready[at]!;
    ordered.push([carrier, byId.get(carrier)!]);
    for (const next of followers.get(carrier) ?? []) {
      const remaining = counts.get(next)! - 1;
      counts.set(next, remaining);
      if (!remaining) ready.push(next);
    }
  }
  // Cyclic peer-written evidence does not describe a native history.
  return ordered;
}

/** Later observed edits replace their predecessors. Independent edits join column identities. */
export class BirthCoverageFrontier {
  private readonly past = new Map<LogicalId, ReadonlySet<LogicalId>>();
  private readonly cells = new Map<LogicalId, Map<LogicalId, ReadonlySet<LogicalId>>>();
  constructor(frames: readonly BirthGeometryFrame[]) {
    const prior = dependencies(frames);
    for (const [carrier] of frames) {
      const past = new Set<LogicalId>();
      for (const parent of prior.get(carrier) ?? []) {
        past.add(parent);
        for (const ancestor of this.past.get(parent) ?? []) past.add(ancestor);
      }
      this.past.set(carrier, past);
    }
  }
  add(cell: LogicalId, carrier: LogicalId, columns: ReadonlySet<LogicalId>): void {
    const opinions = this.cells.get(cell) ?? new Map<LogicalId, ReadonlySet<LogicalId>>();
    for (const predecessor of this.past.get(carrier) ?? []) opinions.delete(predecessor);
    opinions.set(carrier, columns);
    this.cells.set(cell, opinions);
  }
  coverage(cell: LogicalId): ReadonlySet<LogicalId> {
    const result = new Set<LogicalId>();
    for (const columns of this.cells.get(cell)?.values() ?? [])
      for (const column of columns) result.add(column);
    return result;
  }
  concurrent(cell: LogicalId): boolean {
    return (this.cells.get(cell)?.size ?? 0) > 1;
  }
}
