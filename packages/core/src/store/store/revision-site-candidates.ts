import { isContentRevisionKind, WML_NAMESPACE_URI } from '../package/ooxml-shared.ts';
import type { OoxmlElement, OoxmlNode } from '../package/ooxml-tree.ts';
import { keepsSubtreeMemo } from '../package/subtree-memo-policy.ts';
import { REVISION_PROPERTY_CONTAINERS } from './revision-property-records.ts';
import { CELL_REVISION_NAMES } from './revision-table-plan.ts';

export const MARK_MOVE_NAMES: ReadonlySet<string> = new Set(['moveFrom', 'moveTo']);

/** The detailed classifier also checks the parent of a legacy numbering record. */
export function isNamedRevision(node: OoxmlElement, numberingParent: boolean): boolean {
  return (
    isContentRevisionKind(node.kind) ||
    MARK_MOVE_NAMES.has(node.localName) ||
    CELL_REVISION_NAMES.has(node.localName) ||
    REVISION_PROPERTY_CONTAINERS.has(node.localName) ||
    node.localName === 'ins' ||
    node.localName === 'del' ||
    (numberingParent && node.localName === 'numberingChange')
  );
}

/** Cheap canonical preflight for an element that may produce a review revision item. @internal */
export function isPotentialRevisionElement(node: OoxmlNode): boolean {
  return (
    node.kind === 'moveFromRangeStart' ||
    node.kind === 'moveToRangeStart' ||
    (node.kind !== 'textValue' &&
      node.namespaceUri === WML_NAMESPACE_URI &&
      isNamedRevision(node, true))
  );
}

const candidates = new WeakMap<OoxmlElement, boolean>();

/** A neighboring edit retains the exact negative answer for each unchanged subtree. */
export function mayHaveRevisionSite(node: OoxmlNode, depth = 0): boolean {
  if (node.kind === 'textValue') return false;
  const cached = candidates.get(node);
  if (cached !== undefined) return cached;
  // Let the ordinary classifier decide deeply nested content. It owns all positive answers.
  if (depth >= 64) return true;
  const result =
    isPotentialRevisionElement(node) ||
    node.children.some((child) => mayHaveRevisionSite(child, depth + 1));
  if (keepsSubtreeMemo(node)) candidates.set(node, result);
  return result;
}
