import type { CollectedControl, CollectedControlIndex } from './content-control-boundary-layout.ts';
import type { LayoutSession } from './layout-session.ts';
import type { LayoutBox, SemanticLayout } from './semantic-records.ts';

export interface PlacedBlockBox {
  readonly pageIndex: number;
  readonly blockId: string;
  readonly box: LayoutBox;
}

export interface PlacedSpanBox {
  readonly pageIndex: number;
  readonly paragraphId: string;
  readonly start: number;
  readonly end: number;
  /**
   * Ordinal of the line this span sits on, so inline-control fragments can union per LINE.
   * Uniting per page gave a wrapped control one rectangle covering everything between its
   * first and last line, including neighbouring words. Composed as page index × 2^20 plus
   * the line's ordinal WITHIN its page, which keeps document order sortable while letting
   * an unchanged page's contribution be reused verbatim when other pages move.
   */
  readonly line: number;
  /**
   * The span's TEXT extent: the raw span box dropped by the line's leading. Span boxes sit
   * at the line-box top, but non-single `w:spacing` puts the whole leading ABOVE the glyphs,
   * so a boundary built from raw boxes tints the gap over the text and misses the text
   * itself.
   */
  readonly box: LayoutBox;
}

export interface PageGeometryContribution {
  /** Version of the needed-id sets this contribution was filtered under. */
  readonly neededStamp: number;
  readonly blocks: readonly PlacedBlockBox[];
  readonly spans: readonly PlacedSpanBox[];
}


interface PublishedBoundaries {
  readonly index: CollectedControlIndex;
  readonly contributions: readonly PageGeometryContribution[];
  readonly raw: SemanticLayout;
  readonly published: SemanticLayout;
}

const recent = new WeakMap<LayoutSession, PublishedBoundaries>();

function sameBox(left: LayoutBox, right: LayoutBox): boolean {
  return left === right || (left.x === right.x && left.y === right.y &&
    left.width === right.width && left.height === right.height);
}

function sameContribution(left: PageGeometryContribution, right: PageGeometryContribution): boolean {
  if (left === right) return true;
  if (left.blocks.length !== right.blocks.length || left.spans.length !== right.spans.length) return false;
  for (let i = 0; i < left.blocks.length; i++) {
    const a = left.blocks[i]!, b = right.blocks[i]!;
    if (a.blockId !== b.blockId || a.pageIndex !== b.pageIndex || !sameBox(a.box,b.box)) return false;
  }
  for (let i = 0; i < left.spans.length; i++) {
    const a = left.spans[i]!, b = right.spans[i]!;
    if (a.paragraphId !== b.paragraphId || a.pageIndex !== b.pageIndex || a.line !== b.line ||
      a.start !== b.start || a.end !== b.end || !sameBox(a.box,b.box)) return false;
  }
  return true;
}

/** A block w:sdt (CT_SdtBlock, ECMA-376 Part 1 §17.5.2.29) follows its blocks.
 * An inline w:sdt (CT_SdtRun, §17.5.2.31) follows its text range. Reuse their
 * records only while the collected ranges and placed geometry still agree.
 */
export function retainControlBoundaries(
  session: LayoutSession,
  layout: SemanticLayout,
  index: CollectedControlIndex,
  contributions: readonly PageGeometryContribution[],
  token: string,
  sameControl: (left: CollectedControl, right: CollectedControl) => boolean,
  build: () => SemanticLayout
): SemanticLayout {
  const previous = recent.get(session);
  let published: SemanticLayout;
  if (previous && previous.published.controlContextToken === token &&
    contributions.length === previous.contributions.length &&
    (index === previous.index || (index.controls.length === previous.index.controls.length &&
      index.controls.every((entry, i) => sameControl(entry, previous.index.controls[i]!)))) &&
    contributions.every((entry, i) => sameContribution(entry, previous.contributions[i]!))) {
    const pages = layout.pages.map((page, i) => {
      const prior = previous.published.pages[i]!;
      if (page === previous.raw.pages[i] || page === prior) return prior;
      const controls = prior.contentControls;
      return controls === page.contentControls ? page : { ...page, contentControls: controls };
    });
    published = { ...layout, pages: pages.every((page, i) => page === layout.pages[i]) ? layout.pages : pages,
      contentControls: previous.published.contentControls, controlContextToken: token };
  } else {
    published = build();
  }
  recent.set(session, { raw: layout, published, contributions, index });
  return published;
}
