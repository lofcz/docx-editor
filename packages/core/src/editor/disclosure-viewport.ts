import type { CollapsedContentControl } from '../contracts/content-control-collapse.ts';
import type { SemanticLayout } from '../layout/semantic-records.ts';
import { surfaceScroller, viewportInLayout } from './surface-pages.ts';

/** Only one explicitly changed disclosure can own a visible position. */
export function changedDisclosure(
  before: readonly CollapsedContentControl[],
  after: readonly CollapsedContentControl[]
): string | undefined {
  const previous = new Map(before.map((control) => [control.controlId, control]));
  let changed: string | undefined;
  for (const control of after) {
    const old = previous.get(control.controlId);
    previous.delete(control.controlId);
    if (
      old &&
      old.summaryParagraphId === control.summaryParagraphId &&
      old.placeholderHeightPt === control.placeholderHeightPt
    )
      continue;
    if (changed) return undefined;
    changed = control.controlId;
  }
  if (!previous.size) return changed;
  return !changed && previous.size === 1 ? previous.keys().next().value : undefined;
}

function headerTop(layout: SemanticLayout, id: string): number | undefined {
  const fragment = layout.contentControls?.find((control) => control.id === id)?.fragments[0];
  const page = fragment && layout.pages[fragment.pageIndex];
  if (page && fragment) return page.contentBox.y + fragment.box.y;
}

/** Keep a locally toggled header visible through pagination and later widget measurements. */
export function createDisclosureViewport(container: () => HTMLElement) {
  let anchor:
    | {
        id: string;
        scroller: HTMLElement;
        scale: number;
        revision: number;
        y: number;
        scrollTop: number;
        viewportTop: number;
      }
    | undefined;
  return {
    capture(id: string | undefined, layout: SemanticLayout, scale: number, revision: number) {
      anchor = undefined;
      if (!id) return;
      const host = container();
      const scroller = surfaceScroller(host);
      if (!scroller || !scroller.scrollTop || !scroller.clientHeight) return;
      const y = headerTop(layout, id);
      if (y === undefined) return;
      const view = viewportInLayout(host, scroller, scale, scroller.clientHeight);
      if (y < view.top || y >= view.top + view.height) return;
      anchor = {
        id,
        scroller,
        scale,
        revision,
        y,
        scrollTop: scroller.scrollTop,
        viewportTop: view.top,
      };
    },
    plan(layout: SemanticLayout, scale: number, revision: number) {
      if (!anchor) return;
      const current = anchor;
      if (
        scale !== current.scale ||
        revision !== current.revision ||
        surfaceScroller(container()) !== current.scroller ||
        current.scroller.scrollTop !== current.scrollTop
      ) {
        anchor = undefined;
        return;
      }
      const y = headerTop(layout, current.id);
      if (y === undefined) {
        anchor = undefined;
        return;
      }
      const delta = y - current.y;
      if (Math.abs(delta * scale) < 0.5) return;
      const top = current.viewportTop + delta;
      return {
        top,
        restore() {
          if (current !== anchor) return;
          current.scroller.scrollTop = current.scrollTop + delta * scale;
          const actual = current.scroller.scrollTop;
          current.viewportTop += (actual - current.scrollTop) / scale;
          current.scrollTop = actual;
          current.y = y;
        },
      };
    },
  };
}
