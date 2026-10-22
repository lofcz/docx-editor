import type { LogicalBlock } from './logical-block-types';
import { expandTargetSpans } from './target-table-spans';
import { EDIT_PLAN_LIMITS } from './plan';

/** Expand source HTML spans into read-only logical slots. Native bindings still validate every slot. */
export function expandLogicalSourceSpans(input: readonly LogicalBlock[]): {
  blocks: readonly LogicalBlock[];
  virtual: ReadonlySet<LogicalBlock>;
} {
  const virtual = new Set<LogicalBlock>();
  let count = 0;
  const charge = () => {
    if (++count > EDIT_PLAN_LIMITS.maxNodes)
      throw new Error('Expanded logical source exceeds its node limit.');
  };
  const visit = (block: LogicalBlock): LogicalBlock => {
    charge();
    if (!('children' in block)) return block;
    const current = { ...block, children: block.children.map(visit) };
    if (current.kind !== 'table') return current;
    const expanded = expandTargetSpans(current);
    if (!expanded.success) throw new Error(expanded.error);
    for (const key of expanded.expanded.covered) {
      const [row, column] = key.split('/').map(Number);
      const rows = expanded.expanded.table.children;
      const selected = rows[row!];
      if (!selected || !('children' in selected) || !selected.children[column!])
        throw new Error('A logical source span has no covered slot.');
      const slot = selected.children[column!]!;
      charge();
      virtual.add(slot);
    }
    return expanded.expanded.table;
  };
  return { blocks: input.map(visit), virtual };
}
