import type { PaginatedSurface } from '@docx-editor.dev/core/editor';
import { contentControlsOfLayout } from '@docx-editor.dev/core/layout';

interface FoldTransition {
  button: HTMLButtonElement;
  targetOpen: boolean;
  reverse(open: boolean, change: () => boolean): boolean;
  cancel(): void;
}

// A document owns one browser snapshot transition. Different windows have independent transitions.
const active = new WeakMap<Document, FoldTransition>();

function visiblePane(pages: HTMLElement, win: Window): HTMLElement {
  for (let el = pages.parentElement; el && el !== pages.ownerDocument.body; el = el.parentElement) {
    if (/(auto|scroll|hidden|clip)/.test(win.getComputedStyle(el).overflowY) && el.clientHeight)
      return el;
  }
  return pages;
}

/** The last native fragment can end on a later page, including one outside the scroll pane. */
function bottomEdge(surface: PaginatedSurface, pages: HTMLElement, id: string, pane: DOMRect) {
  const layout = surface.publishedLayout();
  const end = contentControlsOfLayout(layout)
    .find((record) => record.id === id)
    ?.fragments.at(-1);
  if (!end) return undefined;
  const page = pages.children.item(end.pageIndex);
  const sheet = layout.pages[end.pageIndex];
  if (!page || !sheet) return undefined;
  const box = page.getBoundingClientRect();
  const scale = box.width / sheet.box.width;
  const bottom = box.top + (sheet.contentBox.y - sheet.box.y + end.box.y + end.box.height) * scale;
  return Math.max(0, Math.min(pane.height, bottom - pane.top));
}

function followAnimation(doc: Document): Animation | undefined {
  return doc
    .getAnimations?.()
    .find((animation) => (animation as CSSAnimation).animationName === 'docx-mdx-move-following');
}

/** A second press can arrive before the browser finishes presenting the previous state. */
export function pendingDisclosureState(button: HTMLButtonElement): boolean | undefined {
  const pending = active.get(button.ownerDocument);
  return pending?.button === button ? pending.targetOpen : undefined;
}

export function cancelDisclosureMotion(button: HTMLButtonElement): void {
  const previous = active.get(button.ownerDocument);
  if (previous?.button === button) previous.cancel();
}

/** Animate the visible boundary and following Word content over the actual final page paint. */
export function changeDisclosure(
  surface: PaginatedSurface,
  controlId: string,
  button: HTMLButtonElement,
  opening: boolean,
  change: () => boolean
): void {
  const doc = button.ownerDocument;
  const win = doc.defaultView;
  const previous = active.get(doc);
  const motion = win?.matchMedia?.('(prefers-reduced-motion: no-preference)');
  const pages = button.closest<HTMLElement>('.docx-pages');
  const pane = pages && win ? visiblePane(pages, win) : null;
  function apply() {
    const top = button.getBoundingClientRect().top;
    if (!change()) return false;
    surface.flushPendingInput();
    if (pane && pane !== pages && button.isConnected) {
      const delta = button.getBoundingClientRect().top - top;
      if (Math.abs(delta) > 0.5) pane.scrollTop += delta;
    }
    return true;
  }
  if (previous?.button === button && previous.reverse(opening, apply)) return;
  previous?.cancel();
  if (!motion?.matches || !doc.startViewTransition || !pages || !pane) {
    apply();
    return;
  }
  const rect = pane.getBoundingClientRect();
  const before = bottomEdge(surface, pages, controlId, rect);
  if (before === undefined || !rect.height || !rect.width) {
    apply();
    return;
  }
  const root = doc.documentElement;
  const paneName = pane.style.viewTransitionName;
  const glyph = button.querySelector('svg');
  const glyphName = glyph?.style.viewTransitionName ?? '';
  pane.style.viewTransitionName = 'docx-mdx-paper';
  if (glyph) glyph.style.viewTransitionName = 'docx-mdx-arrow';
  const motionClass = opening ? 'docx-mdx-opening' : 'docx-mdx-closing';
  root.classList.add(motionClass);
  let pending = true;
  let accepted = false;
  let ready = false;
  let reverseChange: (() => boolean) | undefined;
  function applyPending() {
    if (!pending) return;
    pending = false;
    accepted = apply();
  }
  const transition = doc.startViewTransition(() => {
    applyPending();
    const after = bottomEdge(surface, pages, controlId, rect);
    if (
      !accepted ||
      active.get(doc) !== state ||
      after === undefined ||
      Math.abs(after - before) < 1
    ) {
      transition.skipTransition();
      return;
    }
    root.style.setProperty('--docx-mdx-closed-edge', Math.min(after, before) + 'px');
    root.style.setProperty('--docx-mdx-travel', Math.abs(after - before) + 'px');
  });
  const finish = () => {
    motion.removeEventListener('change', cancel);
    pane.removeEventListener('wheel', cancel);
    pane.removeEventListener('touchmove', cancel);
    if (active.get(doc) !== state) return;
    active.delete(doc);
    root.classList.remove(motionClass);
    root.style.removeProperty('--docx-mdx-closed-edge');
    root.style.removeProperty('--docx-mdx-travel');
    pane.style.viewTransitionName = paneName;
    if (glyph) glyph.style.viewTransitionName = glyphName;
    reverseChange?.();
  };
  function cancel() {
    transition.skipTransition();
    applyPending();
    finish();
  }
  const state: FoldTransition = {
    button,
    targetOpen: opening,
    cancel,
    reverse(open, update) {
      const animation = ready && followAnimation(doc);
      if (!animation || animation.playState !== 'running') return false;
      // The incoming browser snapshot is live. Keep its Word layout until the reversed
      // boundary reaches the original position, or the following text jumps immediately.
      reverseChange = open === opening ? undefined : update;
      state.targetOpen = open;
      button.setAttribute('aria-expanded', String(open));
      animation.reverse();
      for (const glyphMotion of doc.getAnimations()) {
        if (
          (glyphMotion.effect as KeyframeEffect | null)?.pseudoElement ===
          '::view-transition-group(docx-mdx-arrow)'
        )
          glyphMotion.reverse();
      }
      return true;
    },
  };
  active.set(doc, state);
  motion.addEventListener('change', cancel);
  pane.addEventListener('wheel', cancel, { once: true, passive: true });
  pane.addEventListener('touchmove', cancel, { once: true, passive: true });
  void transition.ready.then(
    () => {
      ready = true;
    },
    () => {}
  );
  void transition.finished.then(finish, finish);
}
