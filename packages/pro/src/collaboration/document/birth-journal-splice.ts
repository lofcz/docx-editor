/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { asLogicalId, type LogicalId } from './identity.ts';
import type { DocumentRegistry } from './registry.ts';
import type { SharedEffect } from './shared-effect.ts';
/** A canonical splice addresses the visible prefix, not retained hidden birth scaffolds. */
export interface BirthChildSplice {
  readonly removed: readonly LogicalId[];
  readonly left: LogicalId | null;
  readonly right: LogicalId | null;
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
  const remaining = array.toArray();
  const right = splice.right === null ? -1 : remaining.indexOf(splice.right);
  const left = splice.left === null ? -1 : remaining.indexOf(splice.left);
  const at =
    right >= 0 ? right : left >= 0 ? left + 1 : splice.right === null ? remaining.length : 0;
  registry.spliceChildren(effect.parentLogicalId, at, 0, effect.childLogicalIds);
}
