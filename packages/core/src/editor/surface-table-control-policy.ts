import type { PaginatedSurface } from './paginated-surface-contract.ts';
import type { PaginatedSurfaceOptions } from './paginated-surface-options.ts';
import type { TableCommandPlan } from './table-command-plan.ts';

const policies = new WeakMap<
  PaginatedSurface,
  NonNullable<PaginatedSurfaceOptions['tableControlRefusal']>
>();

/** Keep host control ownership identical in command admission, pointer furniture, and execution. */
export function registerTableControlPolicy(
  surface: PaginatedSurface,
  policy: PaginatedSurfaceOptions['tableControlRefusal']
): void {
  if (policy) policies.set(surface, policy);
}

export function tableControlRefusal(surface: PaginatedSurface, tableId: string): string | null {
  return policies.get(surface)?.(tableId, surface.session.currentPackage()) ?? null;
}

/** Selection and content edits stay native. Only generic table command plans meet this policy. */
export function guardTableControlPlan(
  surface: PaginatedSurface,
  plan: TableCommandPlan
): TableCommandPlan {
  if (!plan.ok || !policies.has(surface)) return plan;
  for (const op of plan.ops) {
    const id = 'tableId' in op ? op.tableId : op.op === 'deleteBlock' ? op.blockId : null;
    const reason = id === null ? null : tableControlRefusal(surface, id);
    if (reason !== null) return { ok: false, code: 'unsupported', reason };
  }
  return plan;
}
