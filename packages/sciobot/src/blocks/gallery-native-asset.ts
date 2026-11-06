import {
  findNode,
  resolveImageResourceLimits,
  sha256Bytes,
  sniffImageMime,
  validateRasterHeader,
  type OoxmlPackage,
  type SupportedImageMime,
} from '@docx-editor.dev/core/store';
import { readGallery } from './gallery';
import { readNativeGalleryImage } from './gallery-native-record';

const extensions: Readonly<Record<string, string>> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/webp': 'webp',
};

/** Read a validated embedded image from its exact document owner. This function performs no loading. */
export function readGalleryImageAsset(pkg: OoxmlPackage, blockId: string, imageId: string) {
  const gallery = readGallery(pkg, blockId);
  if (gallery?.payload.version !== 3) return null;
  const item = gallery.images.find((image) => image.id === imageId);
  const part = pkg.parts.get(gallery.partName);
  const control = item && part && findNode(part, item.controlId);
  const native = control && part && readNativeGalleryImage(pkg, part, control);
  if (!native) return null;
  const source =
    pkg.partBytes.get(native.resource.partName) ??
    pkg.partBytes.get(native.resource.partName.replace(/^\//, ''));
  const mime = source && sniffImageMime(source);
  if (!source || !mime || !extensions[mime]) return null;
  const limits = resolveImageResourceLimits();
  if (!source.length || source.length > limits.maxEncodedBytes) return null;
  const header = validateRasterHeader(source, mime as SupportedImageMime);
  if (
    !header ||
    header.pixelWidth > limits.maxDimension ||
    header.pixelHeight > limits.maxDimension ||
    header.pixelWidth * header.pixelHeight > limits.maxPixels
  )
    return null;
  const contentHash = sha256Bytes(source).replace(/^sha256:/, '');
  return {
    bytes: source,
    mime: mime as SupportedImageMime,
    contentHash,
    path: `media/${contentHash}.${extensions[mime]}`,
    partName: native.resource.partName,
    drawingNodeId: native.fields.drawing.id,
    attributes: native.attributes,
  };
}
