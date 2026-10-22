import { findNode } from '../package/ooxml-edit.ts';
import { WML_NAMESPACE_URI, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { parentOf } from './tree-op-nodes.ts';
import { paragraphOffsetIndex, splitsSurrogate } from './tree-op-segments.ts';
import { isValidXmlText } from '../package/sinks.ts';
import type { TreeOpRejection } from './tree-op-types.ts';

/** Longest `r:id`, `w:anchor` or `w:tooltip` an op may write. */
export const MAX_HYPERLINK_ATTRIBUTE_LENGTH = 512;

/** Partial link writes support ordinary text runs; opaque inline content remains untouched. */
export function isTextHyperlink(node: OoxmlNode): boolean {
  return (
    node.kind === 'hyperlink' &&
    node.children.every(
      (run) =>
        run.kind === 'run' &&
        run.children.every(
          (child) =>
            child.namespaceUri === WML_NAMESPACE_URI &&
            (child.localName === 'rPr' || child.localName === 't')
        )
    )
  );
}

export function validateHyperlinkRange(
  part: OoxmlPart,
  linkId: string,
  range: { readonly start: number; readonly end: number }
): TreeOpRejection | null {
  const link = findNode(part, linkId);
  const paragraph = link && parentOf(part, linkId);
  if (!link || !isTextHyperlink(link) || paragraph?.kind !== 'paragraph')
    return 'invalid-property-value';
  const span = paragraphOffsetIndex(paragraph).spanOf(link);
  if (
    !span ||
    !Number.isInteger(range.start) ||
    !Number.isInteger(range.end) ||
    range.start < span.start ||
    range.end > span.end ||
    range.start >= range.end
  )
    return 'invalid-range';
  if (splitsSurrogate(paragraph, range.start) || splitsSurrogate(paragraph, range.end))
    return 'splits-surrogate-pair';
  return null;
}

/**
 * The target half of an insert or a retarget: EXACTLY ONE of relationship or anchor, and
 * every value legal in XML.
 *
 * Both at once is refused rather than resolved by precedence. In a FILE that pair means "a
 * bookmark inside another document", which the read side honours; as an authored op it means
 * the caller does not know which it wants, and admitting it would write a link whose
 * behaviour depends on a resolution rule the caller never saw.
 */
export function validateHyperlinkTarget(op: {
  readonly relationshipId?: string;
  readonly anchor?: string;
  readonly tooltip?: string;
}): TreeOpRejection | null {
  const hasRelationship = op.relationshipId !== undefined;
  const hasAnchor = op.anchor !== undefined;
  if (hasRelationship === hasAnchor) return 'invalid-property-value';
  for (const value of [op.relationshipId, op.anchor, op.tooltip]) {
    if (value === undefined) continue;
    if (typeof value !== 'string' || value.length === 0) return 'invalid-property-value';
    if (value.length > MAX_HYPERLINK_ATTRIBUTE_LENGTH) return 'invalid-property-value';
    if (!isValidXmlText(value)) return 'invalid-property-value';
  }
  return null;
}
