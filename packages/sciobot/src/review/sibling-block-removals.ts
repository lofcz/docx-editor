import { contentControlLevelOf, findNode, type OoxmlPackage } from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import type { CanonicalEdit } from '../diff/plan';
import { readPresenceOwners } from './block-presence-owner';
import { refuseWrite } from '../diff/fragment-seeds';

/** Resolve complete application block roots. Native field wrappers never acquire application ownership. */
export function siblingBlockRemovals(
  pkg: OoxmlPackage,
  operations: readonly CanonicalEdit[]
): readonly string[] {
  const index = buildBlockIndex(pkg);
  const byControl = new Map([...index.byId.values()].map((block) => [block.controlId, block]));
  const ids: string[] = [];
  for (const edit of operations) {
    if (edit.kind !== 'siblings') continue;
    const part = pkg.parts.get(edit.partName);
    if (!part) continue;
    for (const id of edit.removedIds) {
      const node = findNode(part, id),
        block = byControl.get(id);
      if (node?.kind !== 'contentControl' || !block) continue;
      if (block.partName !== part.name || contentControlLevelOf(node) !== 'block')
        refuseWrite(
          'Removed application roots need complete native block ownership.',
          'unsupported-operation'
        );
      ids.push(block.id);
    }
  }
  if (ids.length) {
    const owners = readPresenceOwners(pkg, ids);
    if (!owners.success) refuseWrite(owners.error, owners.code);
  }
  return ids;
}
