import {
  contentControlTextOf,
  contentControlsIn,
  findNode,
  paragraphOffsetIndex,
  paragraphTextOf,
  parentNodeOf,
  type OoxmlNode,
  type OoxmlPart,
  type OoxmlParagraphNode,
} from '@docx-editor.dev/core/store';
import { atomicFallbackParagraph, propertyFallbackParagraph } from './property-fallback';

/** Resolve native ownership, never a label search or a cached numeric span. */
export interface PayloadFallback {
  readonly paragraph: OoxmlParagraphNode;
  readonly node: OoxmlNode;
  readonly text: string;
  readonly offset: number;
  readonly inline: boolean;
}

export function payloadFallback(
  part: OoxmlPart,
  controlId: string,
  atomic: boolean
): PayloadFallback | null {
  const paragraph = (atomic ? atomicFallbackParagraph : propertyFallbackParagraph)(part, controlId);
  if (paragraph?.kind === 'paragraph') {
    const text = paragraphTextOf(part, paragraph.id);
    return typeof text === 'string'
      ? { paragraph, node: paragraph, text, offset: 0, inline: false }
      : null;
  }
  if (!atomic) return null;
  const control = findNode(part, controlId);
  if (control?.kind !== 'contentControl' || contentControlsIn(control).length !== 0) return null;
  let parent = parentNodeOf(part, controlId);
  // Inline controls may be wrapped by native hyperlinks or run-level generic owners.
  for (let depth = 0; parent && depth < 64; depth++) {
    if (parent.kind === 'paragraph') {
      const span = paragraphOffsetIndex(parent).spanOf(control);
      return span
        ? {
            paragraph: parent,
            node: control,
            text: contentControlTextOf(control, '\n'),
            offset: span.start,
            inline: true,
          }
        : null;
    }
    if (
      parent.kind !== 'hyperlink' &&
      !(parent.kind === 'generic' && parent.localName === 'sdtContent')
    )
      return null;
    parent = parentNodeOf(part, parent.id);
  }
  return null;
}

/** Native sites belong to their control even when another citation shares the paragraph. */
export function payloadFallbackContains(
  part: OoxmlPart,
  fallback: PayloadFallback,
  node: OoxmlNode
): boolean {
  let current: OoxmlNode | null = node;
  for (let depth = 0; current && depth < 256; depth++) {
    if (current.id === fallback.node.id) return true;
    current = parentNodeOf(part, current.id);
  }
  return false;
}
