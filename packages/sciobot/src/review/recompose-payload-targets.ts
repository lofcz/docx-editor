import { type OoxmlPackage } from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import type { BlockPayload } from '../blocks/payload-schema';
import type { NativeEditFailure } from '../contracts/document';
import { copyBoundedJson } from '../data/json';
import { EDIT_PLAN_LIMITS, type CanonicalEditPlan } from '../diff/plan';
import { payloadFallback } from './payload-fallback';
import type { TrackedPayloadOptions } from './tracked-stage-payloads';

/** Plain host data from the ownership-relative baseline. It grants no write authority. */
export interface RecomposePayloadTarget {
  readonly executor: 'properties' | 'atomic';
  readonly blockId: string;
  readonly before: BlockPayload;
  readonly after: BlockPayload;
  readonly beforeText: string;
}
export type RecomposePayloadOptions = Pick<TrackedPayloadOptions, 'properties' | 'atomic'>;

/** Resolve only missing host labels. Never project the whole document into a replacement. */
export function recomposePayloadTargets(
  pkg: OoxmlPackage,
  plan: CanonicalEditPlan,
  supplied: TrackedPayloadOptions
):
  | { readonly success: true; readonly targets: readonly RecomposePayloadTarget[] }
  | NativeEditFailure {
  const fail = (error: string): NativeEditFailure => ({
    success: false,
    code: 'unsupported-operation',
    error,
  });
  const index = buildBlockIndex(pkg);
  const targets = new Map<string, RecomposePayloadTarget>();
  for (const edit of plan.operations) {
    if (edit.kind !== 'payload') continue;
    const executor = edit.path.length === 0 ? 'atomic' : 'properties';
    if (supplied[executor]) continue;
    const owner = index.byId.get(edit.blockId);
    const part = owner && pkg.parts.get(owner.partName);
    if (!owner || !part || owner.controlId !== edit.controlId || owner.partName !== edit.partName)
      return fail('A payload target has no exact native owner.');
    let target = targets.get(owner.id);
    if (target && target.executor !== executor)
      return fail('A block has conflicting payload executors.');
    if (!target) {
      const fallback = payloadFallback(part, owner.controlId, executor === 'atomic');
      const beforeText = fallback?.text;
      if (typeof beforeText !== 'string')
        return fail('The changed block needs an exact native fallback owner.');
      target = {
        executor,
        blockId: owner.id,
        before: owner.payload,
        after: owner.payload,
        beforeText,
      };
    }
    let after: BlockPayload;
    if (executor === 'atomic') {
      if (!edit.after.present) return fail('An atomic payload target cannot be absent.');
      after = edit.after.value as unknown as BlockPayload;
    } else {
      if (
        edit.path.length !== 2 ||
        edit.path[0] !== 'attributes' ||
        typeof edit.path[1] !== 'string'
      )
        return fail('The payload target is not an attribute edit.');
      const attributes = { ...target.after.attributes };
      if (edit.after.present) attributes[edit.path[1]] = edit.after.value;
      else delete attributes[edit.path[1]];
      after = { ...target.after, attributes };
    }
    targets.set(owner.id, { ...target, after });
  }
  // The callback receives detached JSON values, never references into the canonical tree or plan.
  const copied = copyBoundedJson([...targets.values()], EDIT_PLAN_LIMITS);
  return copied.ok
    ? { success: true, targets: copied.value as unknown as readonly RecomposePayloadTarget[] }
    : fail(copied.reason);
}
