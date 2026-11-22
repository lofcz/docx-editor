import type { RetainedPage } from './semantic-paint-paragraph-roots.ts';

/** Place each sheet once while keeping connected sheets in document order. */
export function orderPaintedPages(container: HTMLElement, pages: readonly RetainedPage[]): void {
  // Keyed reconcile instead of `replaceChildren`: retained elements stay where they are —
  // keeping the browser's style and layout for them, and the DOM selection anchored inside
  // them — while changed pages are placed in order and anything else is dropped.
  const kept = new Set<HTMLElement>(pages.map((entry) => entry.element));
  let child = container.firstChild;
  while (child) {
    const next = child.nextSibling;
    // Drop stale pages first. Leaving them in front of retained virtual shells makes the
    // ordering pass move every shell out and back on each keystroke.
    if (!kept.has(child as HTMLElement)) (child as ChildNode).remove();
    child = next;
  }
  let cursor = container.firstChild;
  for (const entry of pages) {
    if (entry.element === cursor) {
      cursor = cursor.nextSibling;
      continue;
    }
    container.insertBefore(entry.element, cursor);
  }
}
