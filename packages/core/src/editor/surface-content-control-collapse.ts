import type {
  CollapsedContentControl,
  ContentControlCollapseResult,
  ContentControlCollapseTarget,
} from '../contracts/content-control-collapse.ts';
export type {
  CollapsedContentControl,
  ContentControlCollapseResult,
  ContentControlCollapsePlaceholder,
  ContentControlCollapseTarget,
} from '../contracts/content-control-collapse.ts';

/** Local disclosure state for native block content controls. @public */
export interface SurfaceContentControlCollapse {
  /**
   * A frozen snapshot of this surface's collapsed controls. Its reference stays stable until
   * the collapse state changes. Native paint observers notify disclosure chrome after a change.
   */
  collapsedContentControls(): readonly CollapsedContentControl[];
  /**
   * Replace the local collapse view with canonical control-to-summary IDs. An empty map opens
   * every control. Validation refuses the whole request before changing the existing view.
   *
   * A summary must be the first direct paragraph of a block content control in the main story.
   * A null summary hides the whole body behind a blank 24-point local header. Its first flow
   * paragraph can sit within nested controls. A leading table refuses this placeholder mode.
   * Pass `{ placeholderHeightPt: 0 }` to hide all body flow without a header. Positive heights
   * up to 1584 points are rounded to a twip. Invalid heights refuse the complete request.
   * The supported body contains paragraphs, tables, and nested block content controls. Section
   * boundaries and unsupported wrappers refuse collapse. Table cells use the same local view.
   * The input and traversal are bounded.
   *
   * Collapse changes layout only. It never writes canonical content, enters history, reaches a
   * collaboration peer, or affects save and export. It survives host moves and resets on load.
   * Removing an owner or its summary clears that entry. Local navigation into hidden content
   * opens its enclosing controls; remote changes keep this reader's remaining entries closed.
   * Local text edits reveal the selected hidden content before planning their operations.
   */
  setCollapsedContentControls(
    controls: ReadonlyMap<string, ContentControlCollapseTarget>
  ): ContentControlCollapseResult;
}
