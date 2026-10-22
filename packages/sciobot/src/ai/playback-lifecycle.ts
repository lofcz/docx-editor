import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';

/** Stop presentation when its document or host disappears. This never changes document state. */
export function watchPlaybackLifecycle(
  editor: DocxEditorInstance,
  cancel: () => void,
  followReducedMotion: boolean
): { current(): boolean; dispose(): void } {
  const surface = editor.surface!;
  const root = surface.hostElement();
  const document = root.ownerDocument;
  const win = document.defaultView;
  const current = () =>
    editor.surface === surface &&
    !surface.parked() &&
    root.isConnected &&
    root.ownerDocument === document;
  const check = () => {
    if (!current()) cancel();
  };
  const stops = [
    editor.on('change', check),
    editor.on('selectionChange', check),
    editor.on('hostChange', check),
  ];
  // Destruction clears editor event listeners. Its DOM disposal still notifies this observer.
  const observer = win?.MutationObserver ? new win.MutationObserver(check) : null;
  observer?.observe(root, { childList: true });
  if (root.parentNode) observer?.observe(root.parentNode, { childList: true });
  const preference = followReducedMotion
    ? win?.matchMedia?.('(prefers-reduced-motion: reduce)')
    : null;
  const changed = () => {
    if (preference?.matches) cancel();
  };
  preference?.addEventListener('change', changed);
  return {
    current,
    dispose() {
      for (const stop of stops) stop();
      observer?.disconnect();
      preference?.removeEventListener('change', changed);
    },
  };
}
