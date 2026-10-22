import type { DirectListAssignment } from '../ai/direct-lists';
import type { InsertedListBinding } from './inserted-list-types';

/** Bind destinations only after native allocation; malformed addresses never fall back by ordinal. */
export function newCellListAssignments(
  paragraphs: readonly string[],
  bindings: readonly InsertedListBinding[] = []
): DirectListAssignment[] {
  return bindings.map((binding) => {
    const paragraphId = binding.path.length === 1 ? paragraphs[binding.path[0]!] : undefined;
    if (!paragraphId) throw new Error('An inserted cell list binding has no allocated paragraph.');
    return { paragraphId, before: null, target: binding.target };
  });
}
