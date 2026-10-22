import {
  contentControlPropertiesOf,
  contentControlContentNodeOf,
  contentControlsIn,
  findNode,
  parentNodeOf,
  paragraphTextOf,
  storyParagraphs,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
  type RevisionAddress,
} from '@docx-editor.dev/core/store';
import { blockIdFromTag } from '../blocks/payload-opc-codec';
import { GENERATED_LABEL_TAG } from '../projection/block-tags';

/** Resolve one explicitly generated label without descending into another application block. */
export function propertyFallbackParagraph(part: OoxmlPart, controlId: string) {
  const control = findNode(part, controlId);
  if (control?.kind !== 'contentControl') return null;
  const labels = contentControlsIn(control).filter(
    (entry) =>
      contentControlPropertiesOf(entry.node).tag === GENERATED_LABEL_TAG &&
      !entry.ancestors.some(
        (ancestor) =>
          ancestor.id !== controlId &&
          blockIdFromTag(contentControlPropertiesOf(ancestor).tag ?? '') !== null
      )
  );
  if (labels.length !== 1) return null;
  const content = contentControlContentNodeOf(labels[0]!.node);
  const paragraphs = content ? storyParagraphs(content) : [];
  return paragraphs.length === 1 ? paragraphs[0]! : null;
}

/** Atomic text fallbacks can use the generated label or one direct native paragraph. */
export function atomicFallbackParagraph(part: OoxmlPart, controlId: string) {
  const control = findNode(part, controlId);
  if (control?.kind !== 'contentControl') return null;
  const content = contentControlContentNodeOf(control);
  if (!content) return null;
  // Never edit a nested application block as the outer atomic fallback.
  if (
    contentControlsIn(content).some(
      (entry) => blockIdFromTag(contentControlPropertiesOf(entry.node).tag ?? '') !== null
    )
  )
    return null;
  const paragraphs = storyParagraphs(content);
  if (paragraphs.length === 1) return paragraphs[0]!;
  const generated = propertyFallbackParagraph(part, controlId);
  // Inserted widgets can retain an empty native caret paragraph after their generated label.
  return generated &&
    paragraphs.every(
      (paragraph) => paragraph.id === generated.id || paragraphTextOf(part, paragraph.id) === ''
    )
    ? generated
    : null;
}

export function paragraphOwner(part: OoxmlPart, node: OoxmlNode): string | null {
  let current: OoxmlNode | null = node;
  for (let depth = 0; current && depth < 64; depth++) {
    if (current.kind === 'paragraph') return current.id;
    current = parentNodeOf(part, current.id);
  }
  return null;
}
export function nativePropertyAddress(node: OoxmlNode): RevisionAddress | null {
  if (node.kind === 'textValue') return null;
  const read = (name: string) =>
      node.attributes.find(
        (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === name
      )?.value,
    id = read('id'),
    author = read('author'),
    date = read('date');
  return id !== undefined && author !== undefined
    ? { id, author, ...(date !== undefined ? { date } : {}) }
    : null;
}
export const propertyAddressKey = (partName: string, address: RevisionAddress) =>
  JSON.stringify([partName, address.id, address.author, address.date ?? null]);

/** Respect control ownership for metadata-only changes as well as generated text changes. */
export function propertyControlRefusal(part: OoxmlPart, id: string): string | null {
  let current = findNode(part, id);
  for (let depth = 0; current && depth < 64; depth++) {
    if (current.kind === 'contentControl') {
      const properties = contentControlPropertiesOf(current);
      if (properties.lock !== 'unlocked' || properties.dataBinding || properties.temporary)
        return 'Locked, bound, or temporary property controls require explicit native reconciliation.';
    }
    current = parentNodeOf(part, current.id);
  }
  return current ? 'Property control nesting exceeds its limit.' : null;
}
