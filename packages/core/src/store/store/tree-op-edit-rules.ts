// The edit rules around one canonical tree operation, before its validated write.
import type { EditOptions } from '../package/ooxml-edit.ts';
import type { OoxmlPart } from '../package/ooxml-tree.ts';
import { withBookmarkEditRules } from './tree-op-bookmarks.ts';
import { healedTableProperties, tableIdOf } from './tree-op-table-heal.ts';
import { withInsertionHints } from './tree-op-text-splice-hints.ts';
import type { TreeDocOp, TreeOpResult } from './tree-op-types.ts';

/** Heal table properties, capture insertion placement, then check bookmark edit rules. */
export function withTreeOpEditRules(
  part: OoxmlPart,
  op: TreeDocOp,
  options: EditOptions | undefined,
  apply: (part: OoxmlPart, op: TreeDocOp, options?: EditOptions) => TreeOpResult
): TreeOpResult {
  // A table edit first merges the property containers a concurrent merge repeated.
  const tableId = tableIdOf(op);
  if (tableId !== null) {
    const healed = healedTableProperties(part, tableId, options);
    if (healed === null) return { ok: false, reason: 'tree-invariant' };
    part = healed;
  }
  const run = () => withBookmarkEditRules(part, op, options, () => apply(part, op, options));
  return op.op === 'insertText' ? withInsertionHints(part, op.paragraphId, op.offset, run) : run();
}
