import type {
  ContentControlCollapseTarget,
  DocxEditorInstance,
  SurfacePaintFrame,
  SurfaceContentControlCollapse,
} from '@docx-editor.dev/core/editor';
import { paragraphTextOf, type OoxmlPackage } from '@docx-editor.dev/core/store';
import { createDisclosureReader, type MdxDocumentComponent } from './document';
import {
  cancelDisclosureMotion,
  changeDisclosure,
  pendingDisclosureState,
} from './document-motion';
import { disclosureAnchors } from './document-painted-pages';

/** The local reading preferences of a mounted editor. @public */
export interface MdxDisclosureOptions {
  /** Override each component's authored `open` attribute on first display. */
  readonly initiallyOpen?: (component: MdxDocumentComponent) => boolean;
  /** Announce an unsupported collapse, for example a section break inside a disclosure. */
  readonly onRefusal?: (component: MdxDocumentComponent, reason: string) => void;
}

function currentFolds(view: SurfaceContentControlCollapse) {
  const values = new Map<string, ContentControlCollapseTarget>();
  for (const item of view.collapsedContentControls())
    values.set(
      item.controlId,
      item.placeholderHeightPt === undefined
        ? item.summaryParagraphId
        : { placeholderHeightPt: item.placeholderHeightPt }
    );
  return values;
}

function initiallyOpen(item: MdxDocumentComponent, options: MdxDisclosureOptions): boolean {
  const value = item.attributes.open;
  return (
    options.initiallyOpen?.(item) ??
    (value === true ||
      typeof value === 'string' ||
      (value !== null && typeof value === 'object' && 'value' in value && value.value === 'true'))
  );
}

function applyFolds(
  view: SurfaceContentControlCollapse,
  items: readonly MdxDocumentComponent[],
  options: MdxDisclosureOptions
): void {
  if (!items.length) return;
  const next = currentFolds(view);
  const pending = items.filter((item) => next.get(item.controlId) !== item.summaryParagraphId);
  if (!pending.length) return;
  for (const item of pending) next.set(item.controlId, item.summaryParagraphId);
  if (view.setCollapsedContentControls(next).ok) return;
  for (const item of pending) {
    const single = currentFolds(view);
    single.set(item.controlId, item.summaryParagraphId);
    const attempt = view.setCollapsedContentControls(single);
    if (!attempt.ok) options.onRefusal?.(item, attempt.reason);
  }
}

/** Apply authored disclosure states before first layout. Existing host folds are preserved. @public */
export function configureMdxDisclosures(
  pkg: OoxmlPackage,
  view: Pick<
    NonNullable<DocxEditorInstance['surface']>,
    'collapsedContentControls' | 'setCollapsedContentControls'
  >,
  options: MdxDisclosureOptions = {}
): void {
  const items = createDisclosureReader()(pkg);
  applyFolds(
    view,
    items.filter((item) => item.summaryParagraphId && !initiallyOpen(item, options)),
    options
  );
}

/**
 * Mount keyboard-accessible disclosure buttons beside native summary paragraphs. Open and
 * closed states belong to this editor instance. Saving and printing the DOCX includes each
 * complete section. Call the returned function when the host unmounts.
 * @public
 */
export function mountMdxDisclosures(
  editor: DocxEditorInstance,
  options: MdxDisclosureOptions = {}
): () => void {
  const surface = editor.surface;
  if (!surface) throw new Error('Attach the document editor before mounting disclosures.');
  const buttons = new Map<string, HTMLButtonElement>();
  let focusedSummary: string | undefined;
  let ownerDocument: Document | undefined;
  const known = new Set<string>();
  const readDisclosures = createDisclosureReader();
  let summaries = new Map<string, MdxDocumentComponent>();
  let previousPackage: OoxmlPackage | undefined;
  let disposed = false;
  const folds = () => currentFolds(surface);
  function refresh() {
    const pkg = surface!.session.currentPackage();
    if (pkg === previousPackage) return;
    previousPackage = pkg;
    const next = new Map<string, MdxDocumentComponent>();
    const initial: MdxDocumentComponent[] = [];
    for (const item of readDisclosures(pkg)) {
      if (!item.summaryParagraphId) continue;
      next.set(item.summaryParagraphId, item);
      if (!known.has(item.controlId)) {
        known.add(item.controlId);
        if (!initiallyOpen(item, options)) initial.push(item);
      }
    }
    summaries = next;
    applyFolds(surface!, initial, options);
  }
  function makeButton(doc: Document, summary: string) {
    const button = doc.createElement('button');
    button.type = 'button';
    button.className = 'docx-mdx-disclosure-toggle';
    button.contentEditable = 'false';
    button.setAttribute('data-docx-host-ui', '');
    button.setAttribute('data-docx-mdx-disclosure', '');
    button.dataset.mdxSummaryId = summary;
    const icon = doc.createElementNS('http://www.w3.org/2000/svg', 'svg');
    icon.setAttribute('viewBox', '0 0 16 16');
    icon.setAttribute('aria-hidden', 'true');
    const path = doc.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', 'm6 3.5 4.5 4.5L6 12.5');
    path.setAttribute('fill', 'none');
    path.setAttribute('stroke', 'currentColor');
    path.setAttribute('stroke-width', '1.6');
    path.setAttribute('stroke-linecap', 'round');
    icon.append(path);
    button.append(icon);
    button.addEventListener('mousedown', (event) => event.preventDefault());
    button.addEventListener('click', (event) => {
      event.stopPropagation();
      const item = summaries.get(summary);
      if (!item || disposed) return;
      const pending = pendingDisclosureState(button);
      const opening = pending === undefined ? folds().has(item.controlId) : !pending;
      const keyboardFocus = doc.activeElement === button;
      changeDisclosure(surface!, item.controlId, button, opening, () => {
        if (disposed || !summaries.has(summary)) return false;
        const next = folds();
        if (opening) next.delete(item.controlId);
        else next.set(item.controlId, summary);
        if (keyboardFocus) focusedSummary = summary;
        const result = surface!.setCollapsedContentControls(next);
        if (!result.ok) {
          options.onRefusal?.(item, result.reason);
          return false;
        }
        button.setAttribute('aria-expanded', String(opening));
        if (keyboardFocus) buttons.get(summary)?.focus({ preventScroll: true });
        return true;
      });
    });
    return button;
  }
  function paint(frame: SurfacePaintFrame) {
    if (disposed) return;
    ownerDocument = frame.pagesLayer.ownerDocument;
    refresh();
    if (!summaries.size && !buttons.size) return;
    const closed = folds();
    const seen = new Set<string>();
    for (const [summary, element] of disclosureAnchors(frame, summaries, buttons)) {
      const item = summaries.get(summary)!;
      const button = buttons.get(summary) ?? makeButton(element.ownerDocument, summary);
      const label = paragraphTextOf(frame.part, summary) || item.name;
      const expanded = String(pendingDisclosureState(button) ?? !closed.has(item.controlId));
      if (button.getAttribute('aria-label') !== label) button.setAttribute('aria-label', label);
      if (button.getAttribute('aria-expanded') !== expanded)
        button.setAttribute('aria-expanded', expanded);
      if (button.parentElement !== element) element.append(button);
      buttons.set(summary, button);
      seen.add(summary);
    }
    for (const [id, button] of buttons)
      if (!seen.has(id)) {
        button.remove();
        buttons.delete(id);
      }
    if (focusedSummary) {
      const doc = frame.pagesLayer.ownerDocument;
      const active = doc.activeElement;
      const button = buttons.get(focusedSummary);
      if (button && (active === doc.body || active === frame.pagesLayer || !active))
        button.focus({ preventScroll: true });
      focusedSummary = undefined;
    }
  }
  function clear() {
    disposed = true;
    for (const button of buttons.values()) {
      cancelDisclosureMotion(button);
      button.remove();
    }
    buttons.clear();
  }
  const foldsBeforeMount = surface.collapsedContentControls();
  refresh();
  const stop = surface.observePaint({
    replayCurrentFrame: foldsBeforeMount === surface.collapsedContentControls(),
    beforePaint() {
      const active = ownerDocument?.activeElement as HTMLElement | null;
      const summary = active?.dataset.mdxSummaryId;
      if (summary && buttons.get(summary) === active) focusedSummary = summary;
    },
    afterPaint: paint,
    destroyed: clear,
  });
  return () => {
    if (disposed) return;
    clear();
    stop();
    const remaining = folds();
    for (const id of known) remaining.delete(id);
    surface.setCollapsedContentControls(remaining);
  };
}
