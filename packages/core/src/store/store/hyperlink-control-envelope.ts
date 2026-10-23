import { contentControlsIn } from '../package/content-control-nodes.ts';
import type { OoxmlParagraphNode } from '../package/ooxml-tree.ts';
import { paragraphOffsetIndex } from './tree-op-segments.ts';

/** A link envelope moves complete controls unchanged; neither boundary may cut their content. */
export function preservesCoveredControl(
  span: { readonly start: number; readonly end: number } | null,
  range: { readonly start: number; readonly end: number } | undefined
): boolean {
  return !!span && !!range && span.start >= range.start && span.end <= range.end;
}

export function hyperlinkPreservedControls(
  paragraph: OoxmlParagraphNode,
  range: { readonly start: number; readonly end: number },
  enabled: boolean
): ReadonlySet<string> {
  if (!enabled) return new Set();
  const offsets = paragraphOffsetIndex(paragraph);
  return new Set(
    contentControlsIn(paragraph).flatMap(({ node }) =>
      preservesCoveredControl(offsets.spanOf(node), range) ? [node.id] : []
    )
  );
}
