import {
  resolveImageResourceLimits,
  sha256Bytes,
  sniffImageMime,
  validateRasterHeader,
  type SupportedImageMime,
} from '@docx-editor.dev/core/store';
import type { MarkdownImageAsset } from '@docx-editor.dev/docx-to-markdown';
import { GALLERY_ASSET_PATH } from '../projection/gallery-assets';
import type { InlineImageResolver } from './inline-image-resources';

/** Resolve an explicitly supplied Markdown attachment bundle. No network or live document is used. */
export function imageAssetBundleResolver(
  assets: readonly Pick<MarkdownImageAsset, 'path' | 'id' | 'mimeType' | 'bytes'>[]
): InlineImageResolver {
  if (!Array.isArray(assets) || assets.length > 10_000)
    throw Error('The image attachment count exceeds its limit.');
  const byPath = new Map<string, { bytes: Uint8Array; mime: SupportedImageMime }>();
  const limits = resolveImageResourceLimits();
  const extensions: Readonly<Record<string, string>> = {
    'image/png': 'png',
    'image/jpeg': 'jpg',
    'image/gif': 'gif',
    'image/bmp': 'bmp',
    'image/webp': 'webp',
  };
  let total = 0;
  for (const asset of assets) {
    if (
      !asset ||
      typeof asset.path !== 'string' ||
      !GALLERY_ASSET_PATH.test(asset.path) ||
      byPath.has(asset.path)
    )
      throw Error('Image attachment paths must be unique content-addressed paths.');
    if (
      !(asset.bytes instanceof Uint8Array) ||
      !asset.bytes.length ||
      asset.bytes.length > limits.maxEncodedBytes ||
      asset.bytes.length > 64 * 1024 * 1024 - total
    )
      throw Error('The image attachments exceed their byte limit.');
    const bytes = new Uint8Array(asset.bytes),
      mime = sniffImageMime(bytes);
    const header = mime && validateRasterHeader(bytes, mime as SupportedImageMime);
    if (
      !mime ||
      mime !== asset.mimeType ||
      !extensions[mime] ||
      !header ||
      header.pixelWidth > limits.maxDimension ||
      header.pixelHeight > limits.maxDimension ||
      header.pixelWidth * header.pixelHeight > limits.maxPixels
    )
      throw Error('The image attachment cannot be decoded safely.');
    const hash = sha256Bytes(bytes).replace(/^sha256:/, '');
    if (asset.id !== hash || asset.path !== `media/${hash}.${extensions[mime]}`)
      throw Error('The image attachment identity does not match its bytes.');
    byPath.set(asset.path, { bytes, mime: mime as SupportedImageMime });
    total += bytes.length;
  }
  return async (reference, signal) => {
    if (signal?.aborted) throw Error('Image resource preparation was aborted.');
    const asset = byPath.get(reference);
    if (!asset) throw Error('The requested image attachment is missing.');
    return { bytes: new Uint8Array(asset.bytes), mime: asset.mime };
  };
}
