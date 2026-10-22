import { createElement } from 'react';

/** Moving native nodes can clear browser selection before the engine mirrors it. */
export function preserveNativeSelection<T>(pages: HTMLElement, action: () => T): T {
  const selection = pages.ownerDocument.defaultView?.getSelection();
  const anchor = selection?.anchorNode,
    focus = selection?.focusNode;
  const native = (node: Node | null | undefined) =>
    node && pages.contains(node) && node.parentElement?.closest('.docx-paragraph-fragment');
  const saved =
    native(anchor) && native(focus)
      ? { anchor: anchor!, focus: focus!, a: selection!.anchorOffset, f: selection!.focusOffset }
      : null;
  try {
    return action();
  } finally {
    if (saved?.anchor.isConnected && saved.focus.isConnected)
      selection!.setBaseAndExtent(saved.anchor, saved.a, saved.focus, saved.f);
  }
}

/** A lease moves the original paint elements. Comments retain their exact sibling positions. */
export class NativeSlotLease {
  private slot: HTMLDivElement | null = null;
  private elements: readonly HTMLElement[] = [];
  private borrowed: { element: HTMLElement; marker: Comment }[] = [];

  constructor(
    private readonly key: string,
    private readonly pages: HTMLElement
  ) {}

  private restore(): void {
    for (const { element, marker } of this.borrowed) {
      if (marker.parentNode) marker.replaceWith(element);
    }
    this.borrowed = [];
  }

  /** Restore before paint or disposal. A delayed React commit cannot adopt an old plan. */
  reset(): void {
    this.restore();
    this.elements = [];
  }

  private adopt(): void {
    if (!this.slot) return;
    for (const element of this.elements) {
      if (!element.parentNode || element === this.slot || element.contains(this.slot)) continue;
      const marker = element.ownerDocument.createComment(`native slot ${this.key}`);
      element.before(marker);
      this.borrowed.push({ element, marker });
      this.slot.append(element);
    }
  }

  set(elements: readonly HTMLElement[]): void {
    preserveNativeSelection(this.pages, () => {
      this.restore();
      this.elements = elements;
      this.adopt();
    });
  }

  // Stable across renders. React owns the empty slot; the native engine owns its children.
  private readonly ref = (slot: HTMLDivElement | null): void => {
    preserveNativeSelection(this.pages, () => {
      this.restore();
      this.slot = slot;
      this.adopt();
    });
  };

  content() {
    return createElement('div', {
      key: this.key,
      ref: this.ref,
      'data-sciobot-native-slot': this.key,
      style: { display: 'contents' },
    });
  }
}
