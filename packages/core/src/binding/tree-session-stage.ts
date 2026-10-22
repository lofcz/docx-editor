import { isHeaderFooterLifecycleOp } from '../store/package/hf-lifecycle.ts';
import { isNoteLifecycleOp } from '../store/package/note-lifecycle.ts';
import type { TransactionContext } from '../store/store/tree-store.ts';
import type { TreeDocOp } from '../store/store/tree-ops.ts';

/** A dependent writer uses the same transaction, with a synchronous, callback-scoped context. */
export function runSessionStage(
  stage: ((ctx: TransactionContext) => void) | undefined,
  ctx: TransactionContext,
  countOperation: () => void
): boolean {
  if (!stage) return true;
  let active = true,
    failed = false;
  let violation: Error | null = null;
  const check = () => {
    if (!active) throw new Error('The native transaction stage is no longer active.');
  };
  const apply = (op: TreeDocOp, run: () => boolean) => {
    check();
    if (failed) return false;
    if (isHeaderFooterLifecycleOp(op) || isNoteLifecycleOp(op)) {
      violation = new Error('Lifecycle operations cannot enter a native transaction stage.');
      throw violation;
    }
    countOperation();
    const ok = run();
    failed ||= !ok;
    return ok;
  };
  try {
    const result: unknown = stage({
      currentPackage: () => {
        check();
        return ctx.currentPackage();
      },
      apply: (op) => apply(op, () => ctx.apply(op)),
      applyTo: (name, op) => apply(op, () => ctx.applyTo(name, op)),
      applyPackage: (edit) => {
        check();
        if (failed) return false;
        const ok = ctx.applyPackage(edit);
        failed ||= !ok;
        return ok;
      },
      selectionBefore: (selection) => {
        check();
        ctx.selectionBefore(selection);
      },
      selectionAfter: (selection) => {
        check();
        ctx.selectionAfter(selection);
      },
    });
    if (violation) throw violation;
    if (
      result !== null &&
      (typeof result === 'object' || typeof result === 'function') &&
      typeof (result as { then?: unknown }).then === 'function'
    ) {
      Promise.resolve(result).catch(() => {});
      throw new Error('Native transaction stages must finish synchronously.');
    }
    return !failed;
  } finally {
    active = false;
  }
}
