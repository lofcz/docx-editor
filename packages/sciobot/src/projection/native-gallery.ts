import type { BlockPayload } from '../blocks/payload-schema';
import { readGallery } from '../blocks/gallery';
import { readGalleryImageAsset } from '../blocks/gallery-native-asset';
import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import { validGalleryAssetPayload, type GalleryAssetImage } from './gallery-assets';

/** Export only exact native owners. The saved document never receives attachment paths or view URLs. */
export function galleryAssetPayload(pkg: OoxmlPackage, blockId: string): BlockPayload {
  const gallery = readGallery(pkg, blockId);
  if (gallery?.payload.version !== 3) throw Error('The embedded gallery owner is unavailable.');
  const images = gallery.images.map((image): GalleryAssetImage => {
    const asset = readGalleryImageAsset(pkg, blockId, image.id);
    if (!asset) throw Error('The embedded gallery image cannot be exported safely.');
    const a = asset.attributes;
    return {
      asset: asset.path,
      width: a.width as number,
      height: a.height as number,
      caption: (a.caption as string | undefined) ?? '',
      alt: (a.alt as string | undefined) ?? '',
      title: (a.title as string | null | undefined) ?? null,
    };
  });
  const payload: BlockPayload = {
    id: blockId,
    type: 'image_gallery',
    version: 1,
    attributes: {
      resourceVersion: '1',
      layout: gallery.payload.attributes.layout!,
      ...(gallery.payload.attributes.align === undefined
        ? {}
        : { align: gallery.payload.attributes.align }),
      images,
    },
  };
  if (!validGalleryAssetPayload(payload))
    throw Error('The gallery exceeds the attachment grammar bounds.');
  return payload;
}
