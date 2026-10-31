import { reviewGroupNamespace } from '../registry/adapter-names';
import type { OoxmlElement } from '@docx-editor.dev/core/store';

/** Read-only compatibility for the earlier native attribute prototype. New groups use custom XML. */

export const GROUP_ID_LIMIT = 16_384;

export function reviewGroupMarker(node: OoxmlElement): string | undefined {
  const values = node.attributes.filter(
    (attribute) =>
      attribute.namespaceUri === reviewGroupNamespace() && attribute.localName === 'group'
  );
  if (values.length > 1) throw new Error('Duplicate native review group markers.');
  const id = values[0]?.value;
  if (
    id !== undefined &&
    (!id.startsWith('g1:') || id.length > GROUP_ID_LIMIT || /[\u0000-\u001f]/.test(id))
  )
    throw new Error('Unsupported or malformed native review group marker.');
  return id;
}
