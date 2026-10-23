import {
  allocateDrawingPropertyId,
  withBinaryPart,
  withEmbeddedImage,
  withRelationship,
  IMAGE_RELATIONSHIP_TYPE,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { withInitializedPayloadStore } from '../blocks/payload-opc-codec';
import type { ResolvedInlineImage } from './inline-image-resources';
const extensions = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/webp': 'webp',
} as const;
/** Every relationship target is correct when first journaled, including nonstandard Word owners. */
export function embedInlineImage(
  pkg: OoxmlPackage,
  owner: string,
  id: string,
  resource: ResolvedInlineImage
): {
  readonly pkg: OoxmlPackage;
  readonly partName: string;
  readonly relationshipId: string;
  readonly docPrId: number;
} {
  if (owner.slice(0, owner.lastIndexOf('/')) === '/word') {
    const embedded = withEmbeddedImage(pkg, owner, resource);
    if (!embedded.ok) throw Error(`Image resource embedding refused: ${embedded.reason}.`);
    return embedded;
  }
  // The public Word-folder mint uses owner-relative "media/...". An absolute OPC target
  // keeps the same native primitives correct for imported main documents in other folders.
  const initialized = withInitializedPayloadStore(pkg);
  if (!initialized.ok)
    throw Error(`Image payload storage initialization refused: ${initialized.reason}.`);
  const partName = `/word/media/scio-${id}.${extensions[resource.mime]}`;
  const key = partName.toLowerCase();
  if (
    [...initialized.pkg.partBytes.keys(), ...initialized.pkg.parts.keys()].some(
      (name) => '/' + name.replace(/^\//, '').toLowerCase() === key
    )
  )
    throw Error('The new inline image media part already exists.');
  const property = allocateDrawingPropertyId(initialized.pkg);
  if (!property.ok) throw Error('The native image property identity is unavailable.');
  const next = withBinaryPart(initialized.pkg, partName, resource.bytes, resource.mime);
  const related = withRelationship(next, owner, IMAGE_RELATIONSHIP_TYPE, partName);
  if (!related.ok) throw Error('The native image relationship could not be created.');
  return {
    pkg: related.pkg,
    partName,
    relationshipId: related.relationshipId,
    docPrId: property.id,
  };
}
