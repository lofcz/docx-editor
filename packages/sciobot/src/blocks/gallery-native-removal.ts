import {
  contentControlPropertiesOf,
  contentControlsIn,
  findNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { IndexedBlock } from './block-index';
import { readGallery } from './gallery';
import { readNativeGalleryImage } from './gallery-native-record';
import { readManagedImage } from './inline-image-update';

/** Identify only native picture resources and validated host locks inside a native gallery. */
export function nativeGalleryRemovalSources(
  pkg: OoxmlPackage,
  part: OoxmlPart,
  block: IndexedBlock
) {
  if (
    block.payload.version !== 3 ||
    !['image_gallery', 'image_gallery_item'].includes(block.payload.type)
  )
    return null;
  const galleryId =
    block.payload.type === 'image_gallery' ? block.id : String(block.payload.galleryId);
  const gallery = readGallery(pkg, galleryId);
  if (!gallery) throw Error('The native gallery is missing or ambiguous.');
  const images =
    block.id === galleryId
      ? gallery.images
      : gallery.images.filter((image) => image.id === block.id);
  if (!images.length) throw Error('The image no longer belongs to this gallery.');
  const unlock: string[] = [];
  const resources = images.map((image) => {
    const control = findNode(part, image.controlId)!;
    const native = readNativeGalleryImage(pkg, part, control);
    if (!native) throw Error('The native gallery image is invalid.');
    const nested = contentControlsIn(native.paragraph);
    if (nested.length) {
      const managed = readManagedImage(pkg, part.name, native.fields.drawing.id);
      if (nested.length !== 1 || !managed || managed.control.id !== nested[0]!.node.id)
        throw Error('The gallery image contains an unsupported native control.');
      const properties = contentControlPropertiesOf(managed.control);
      if (
        properties.dataBinding ||
        (properties.lock && !['unlocked', 'sdtContentLocked'].includes(properties.lock))
      )
        throw Error('The native image control is protected.');
      if (properties.lock === 'sdtContentLocked') unlock.push(managed.control.id);
    }
    return { media: native.resource.partName, relationships: [native.resource.relationshipId] };
  });
  return { unlock, resources };
}
