import type { PaginatedSurface } from '@docx-editor.dev/core/editor';

/** Tie local roots, measurements, or listeners to the native document's lifetime. */
export function retainSurfaceResources(surface: PaginatedSurface, release: () => void): () => void {
  return surface.observePaint({ afterPaint() {}, destroyed: release });
}
