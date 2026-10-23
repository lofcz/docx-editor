/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { asLogicalId, type LogicalId } from './identity.ts';
import type { DocumentRegistry } from './registry.ts';
import type { SharedEffect } from './shared-effect.ts';
/** A canonical splice addresses the visible prefix, not retained hidden birth scaffolds. */
export interface BirthChildGap {
  readonly left: LogicalId | null;
  readonly right: LogicalId | null;
}
export interface BirthChildSplice extends BirthChildGap {
  readonly removed: readonly LogicalId[];
}
export interface ChildMovePlan {
  readonly source: LogicalId | null;
  readonly gap: BirthChildGap | null;
}
function gapIndex(children: readonly string[], gap: BirthChildGap): number {
  const right = gap.right === null ? -1 : children.indexOf(gap.right);
  const left = gap.left === null ? -1 : children.indexOf(gap.left);
  return right >= 0 ? right : left >= 0 ? left + 1 : gap.right === null ? children.length : 0;
}
export function applyBirthChildSplice(
  registry: DocumentRegistry,
  effect: Extract<SharedEffect, { kind: 'spliceChildren' }>,
  splice: BirthChildSplice
): void {
  const array = registry.childArray(effect.parentLogicalId);
  const remove = new Set(splice.removed);
  // Delete only the observed visible entries. Hidden births can still receive late peer edits.
  const children = array.toArray();
  for (let at = children.length - 1; at >= 0; at--)
    if (remove.has(asLogicalId(children[at]!)))
      registry.spliceChildren(effect.parentLogicalId, at, 1, []);
  if (!effect.childLogicalIds.length) return;
  const at = gapIndex(array.toArray(), splice);
  registry.spliceChildren(effect.parentLogicalId, at, 0, effect.childLogicalIds);
}
/** Detach first: same-parent destinations and earlier moves use post-detachment coordinates. */
export function applyChildMove(
  registry: DocumentRegistry,
  effect: Extract<SharedEffect, { kind: 'moveNode' }>,
  plan: ChildMovePlan
): void {
  registry.unlinkFromAllParents(effect.logicalId);
  // The registry parent index updates after the transaction. A preceding move may have
  // changed this node's source already; the admitted prefix names that source exactly.
  if (plan.source) {
    const children = registry.childArray(plan.source).toArray();
    for (let at = children.length - 1; at >= 0; at--)
      if (children[at] === effect.logicalId) registry.spliceChildren(plan.source, at, 1, []);
  }
  const destination = registry.childArray(effect.destinationParentLogicalId);
  const at = plan.gap
    ? gapIndex(destination.toArray(), plan.gap)
    : Math.min(effect.destinationIndex, destination.length);
  registry.spliceChildren(effect.destinationParentLogicalId, at, 0, [effect.logicalId]);
}
