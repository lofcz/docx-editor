import type { OoxmlPart } from '../package/ooxml-tree.ts';
import { findNode, parentNodeOf } from '../package/ooxml-edit.ts';
import { holds, nodeTouchesContentRestriction } from './tree-op-validate-controls.ts';
import {
  isParagraph,
  paragraphLength,
  paragraphOffsetIndex,
  insertionLandingNodeId,
} from './tree-op-segments.ts';
import type { TreeOpRejection } from './tree-op-types.ts';

/** A named text owner must cover this exact paragraph offset. */
export function namedOwnerRefusal(
  part: OoxmlPart,
  paragraphId: string,
  offset: number,
  inside: string,
  allowHyperlink = false
): TreeOpRejection | null {
  if (typeof inside !== 'string' || inside.length === 0) return 'invalidArgs';
  const owner = findNode(part, inside);
  if (!owner) return 'unknown-content-control';
  if (owner.kind !== 'contentControl' && !(allowHyperlink && owner.kind === 'hyperlink'))
    return 'not-a-content-control';
  const paragraph = findNode(part, paragraphId);
  if (!paragraph || !isParagraph(paragraph)) return null;
  const span = holds(owner, paragraphId)
    ? { start: 0, end: paragraphLength(paragraph) }
    : holds(paragraph, inside)
      ? paragraphOffsetIndex(paragraph).spanOf(owner)
      : null;
  if (!span) return 'unknown-content-control';
  if (!Number.isInteger(offset) || offset < span.start || offset > span.end) {
    return 'offset-out-of-range';
  }
  if (owner.kind === 'hyperlink') {
    const landing = insertionLandingNodeId(paragraph, offset, owner);
    const restricted = nodeTouchesContentRestriction(part, landing);
    if (restricted) return restricted;
    for (let node = findNode(part, landing); node; node = parentNodeOf(part, node.id))
      if (node.kind === 'revisionDelete' || node.kind === 'revisionMoveFrom') return 'unsupported';
  }
  return null;
}
