import { caretAt } from '@docx-editor.dev/core/layout';
import type { PaginatedSurface } from './paginated-surface-contract.ts';
import { isInHostUi } from './dom-selection.ts';

/** Restore a native input caret before an unclaimed text key reaches the browser. */
export function prepareCollapsedBrowserInput(
  surface: PaginatedSurface,
  event: KeyboardEvent
): void {
  if (
    event.defaultPrevented ||
    event.isComposing ||
    event.metaKey ||
    // Ctrl+Alt can produce AltGraph text even when the platform omits AltGraph state.
    (event.ctrlKey && !event.altKey && !event.getModifierState?.('AltGraph')) ||
    isInHostUi(event.target as Node | null) ||
    !(Array.from(event.key).length === 1 || event.key === 'Dead' || event.key === 'Process') ||
    !surface.collapsedContentControls().length
  )
    return;
  const selection = surface.state().selection;
  const layout = surface.layout();
  if (caretAt(layout, selection.anchor) && caretAt(layout, selection.head)) return;
  // The existing selection path reveals local folds and mirrors the canonical range into
  // painted DOM. The browser then emits beforeinput; this preflight never inserts text.
  surface.setSelection(selection);
}
