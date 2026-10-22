import type { SurfacePaintFrame } from '@docx-editor.dev/core/editor';

/** CSS viewport height inside the remaining native page content, at the current zoom. */
export function blockViewportHeight(
  frame: SurfacePaintFrame,
  pageIndex: number,
  top: number
): number {
  const page = frame.layout.pages[pageIndex];
  return Math.max(1, ((page?.contentBox.height ?? 0) - top) * frame.scale);
}
