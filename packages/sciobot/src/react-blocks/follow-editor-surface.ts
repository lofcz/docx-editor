import type { DocxEditorInstance, PaginatedSurface } from '@docx-editor.dev/core/editor';

/** A document load replaces the surface; detach and window moves retain it. */
export function followEditorSurface<T extends { dispose(): void }>(
  editor: DocxEditorInstance,
  mount: () => T
) {
  let surface: PaginatedSurface | null = null;
  let host: T | null = null;
  let disposed = false;
  function follow() {
    if (disposed || editor.surface === surface) return;
    host?.dispose();
    host = null;
    surface = editor.surface;
    if (surface) host = mount();
  }
  const stops = [
    editor.on('change', follow),
    editor.on('selectionChange', follow),
    editor.on('hostChange', follow),
  ];
  function dispose() {
    if (disposed) return;
    disposed = true;
    stops.forEach((stop) => stop());
    host?.dispose();
    host = null;
    surface = null;
  }
  try {
    follow();
  } catch (error) {
    dispose();
    throw error;
  }
  return {
    current() {
      follow();
      return host;
    },
    dispose,
  };
}
