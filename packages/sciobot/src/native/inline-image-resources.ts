import { insertedAtomsOfEdit } from '../diff/inserted-cell-atoms';
import {
  validateRasterHeader,
  sniffImageMime,
  resolveImageResourceLimits,
  type ImageDecodePort,
  type SupportedImageMime,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { decodeEditPlan, serializeEditPlan, type CanonicalEditPlan } from '../diff/plan';
import type { InsertedInlineAtom } from '../diff/inserted-atoms';
export interface InlineImageSource {
  readonly bytes: Uint8Array;
  readonly mime: SupportedImageMime;
  readonly widthPoints?: number;
  readonly heightPoints?: number;
}
export type InlineImageResolver = (
  reference: string,
  signal?: AbortSignal
) => Promise<InlineImageSource>;
export interface PreparedInlineImages {
  readonly kind: 'prepared-inline-images';
}
export interface ResolvedInlineImage {
  readonly bytes: Uint8Array;
  readonly mime: SupportedImageMime;
  readonly widthPoints: number;
  readonly heightPoints: number;
  readonly contentHash: string;
}
const stores = new WeakMap<
  PreparedInlineImages,
  { plan: string; resources: ReadonlyMap<string, ResolvedInlineImage> }
>();
export function planInlineDrawings(
  plan: CanonicalEditPlan
): readonly Extract<InsertedInlineAtom, { kind: 'drawing' }>[] {
  return plan.operations
    .flatMap(insertedAtomsOfEdit)
    .filter(
      (atom): atom is Extract<InsertedInlineAtom, { kind: 'drawing' }> => atom.kind === 'drawing'
    );
}
export function preparedInlineImageResources(
  handle: PreparedInlineImages | undefined,
  plan: CanonicalEditPlan
): ReadonlyMap<string, ResolvedInlineImage> | null {
  const stored = handle && stores.get(handle);
  return stored && stored.plan === serializeEditPlan(plan) ? stored.resources : null;
}
/** Executor preparation only. Planning retains references and never calls this resolver. */
export async function prepareInlineImageResources(
  input: CanonicalEditPlan,
  resolve: InlineImageResolver,
  decodePort: ImageDecodePort,
  signal?: AbortSignal
): Promise<{ readonly success: true; readonly images: PreparedInlineImages } | NativeEditFailure> {
  try {
    const decoded = decodeEditPlan(serializeEditPlan(input));
    if (!decoded.success) return decoded;
    const plan = decoded.plan,
      resources = new Map<string, ResolvedInlineImage>();
    let total = 0;
    for (const atom of planInlineDrawings(plan)) {
      if (signal?.aborted) throw Error('Image resource preparation was aborted.');
      if (resources.has(atom.resource.reference)) continue;
      const source = await resolve(atom.resource.reference, signal);
      if (
        !(source.bytes instanceof Uint8Array) ||
        source.bytes.byteLength > 64 * 1024 * 1024 - total
      )
        throw Error('Image resources exceed the preparation byte budget.');
      const mime = source.mime,
        requestedWidth = source.widthPoints,
        requestedHeight = source.heightPoints;
      const bytes = new Uint8Array(source.bytes),
        limits = resolveImageResourceLimits();
      const header = validateRasterHeader(bytes, mime);
      if (
        !header ||
        !bytes.length ||
        bytes.length > limits.maxEncodedBytes ||
        sniffImageMime(bytes) !== mime ||
        header.pixelWidth > limits.maxDimension ||
        header.pixelHeight > limits.maxDimension ||
        header.pixelWidth * header.pixelHeight > limits.maxPixels
      )
        throw Error('The image resource cannot be decoded safely.');
      const decoded = await decodePort.decode(new Uint8Array(bytes), mime, limits, signal);
      if (decoded.pixelWidth !== header.pixelWidth || decoded.pixelHeight !== header.pixelHeight)
        throw Error('The image decoder dimensions disagree with its native header.');
      const widthPoints = requestedWidth ?? (header.pixelWidth * 72) / (header.dpiX ?? 96);
      const heightPoints = requestedHeight ?? (header.pixelHeight * 72) / (header.dpiY ?? 96);
      if (
        ![widthPoints, heightPoints].every(
          (v) => Number.isFinite(v) && Math.round(v * 12700) > 0 && v <= 158400
        )
      )
        throw Error('The image dimensions exceed the native drawing bounds.');
      const hash = await crypto.subtle.digest('SHA-256', new Uint8Array(bytes));
      if (signal?.aborted) throw Error('Image resource preparation was aborted.');
      resources.set(atom.resource.reference, {
        bytes,
        mime,
        widthPoints,
        heightPoints,
        contentHash: [...new Uint8Array(hash)].map((n) => n.toString(16).padStart(2, '0')).join(''),
      });
      total += bytes.byteLength;
    }
    const images: PreparedInlineImages = Object.freeze({ kind: 'prepared-inline-images' });
    stores.set(images, { plan: serializeEditPlan(plan), resources });
    return { success: true, images };
  } catch (error) {
    return {
      success: false,
      code: 'invalid-input',
      error: error instanceof Error ? error.message : String(error),
    };
  }
}
