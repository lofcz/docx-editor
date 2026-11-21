import type { SurfacePaintFrame } from '@docx-editor.dev/core/editor';
import { paragraphFragmentsOf, type SemanticLayout } from '@docx-editor.dev/core/layout';

const painted = new WeakMap<SurfacePaintFrame, ReadonlyMap<number, HTMLElement>>();

/** Read only the page shells and their immediate children. Both MDX observers share the frame. */
export function paintedPageContents(frame: SurfacePaintFrame): ReadonlyMap<number, HTMLElement> {
  const cached = painted.get(frame);
  if (cached) return cached;
  const pages = new Map<number, HTMLElement>();
  for (const page of frame.pagesLayer.children) {
    if (!page.classList.contains('docx-page')) continue;
    for (const child of page.children) {
      if (child.classList.contains('docx-page-content')) {
        pages.set(Number(page.getAttribute('data-page-index')), child as HTMLElement);
        break;
      }
    }
  }
  painted.set(frame, pages);
  return pages;
}

/** Keep anchors in retained paragraphs. Search only a page that needs a new native anchor. */
export function disclosureAnchors(
  frame: SurfacePaintFrame,
  summaries: ReadonlyMap<string, unknown>,
  buttons: ReadonlyMap<string, HTMLButtonElement>
): ReadonlyMap<string, HTMLElement> {
  const anchors = new Map<string, HTMLElement>();
  for (const [index, content] of paintedPageContents(frame)) {
    const page = frame.layout.pages[index];
    if (!page) continue;
    let missing = false;
    for (const fragment of paragraphFragmentsOf(page, true)) {
      const id = fragment.paragraphId;
      if (fragment.fragmentIndex !== 0 || !summaries.has(id)) continue;
      const parent = buttons.get(id)?.parentElement;
      if (parent?.dataset.paragraphId !== id || !content.contains(parent)) {
        missing = true;
        break;
      }
      anchors.set(id, parent);
    }
    if (missing) {
      for (const element of content.querySelectorAll<HTMLElement>(
        '.docx-paragraph-fragment[data-fragment-index="0"]'
      )) {
        const id = element.dataset.paragraphId;
        if (id && summaries.has(id)) anchors.set(id, element);
      }
    }
  }
  return anchors;
}

/** Preserve document-wide fragment indices while visiting only the displayed Word pages. */
export function* paintedComponentFragments(
  layout: SemanticLayout,
  pages: ReadonlyMap<number, HTMLElement>
) {
  for (const pageIndex of pages.keys()) {
    const first = layout.pages[pageIndex]?.contentControls?.[0];
    if (first && first.firstFragmentIndex === undefined) {
      // Earlier core packages publish complete fragment indices only on the document list.
      for (const control of layout.contentControls ?? []) {
        for (const [fragmentIndex, fragment] of control.fragments.entries()) {
          const page = pages.get(fragment.pageIndex);
          if (page) yield { control, fragmentIndex, fragment, page };
        }
      }
      return;
    }
  }
  for (const [pageIndex, page] of pages) {
    for (const control of layout.pages[pageIndex]?.contentControls ?? []) {
      for (const [offset, fragment] of control.fragments.entries()) {
        yield { control, fragmentIndex: control.firstFragmentIndex! + offset, fragment, page };
      }
    }
  }
}
