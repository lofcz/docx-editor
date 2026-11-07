/** One locally collapsed block content control and its visible summary. @public */
export interface CollapsedContentControl {
  /** Canonical content-control node ID in the main document story. */
  readonly controlId: string;
  /** First direct paragraph ID, or null for a blank local header hiding the whole body. */
  readonly summaryParagraphId: string | null;
  /** Explicit local header height. Zero hides all body flow. Omitted means 24 points. */
  readonly placeholderHeightPt?: number;
}

/** Local placeholder size, independent of document content. @public */
export interface ContentControlCollapsePlaceholder {
  /** Finite height from 0 to 1584 points, rounded to a twip. Zero publishes no header. */
  readonly placeholderHeightPt: number;
}

/** A canonical summary, the default blank header, or an explicitly sized header. @public */
export type ContentControlCollapseTarget = string | null | ContentControlCollapsePlaceholder;

/** Validation outcome for a complete replacement of the local collapse view. @public */
export type ContentControlCollapseResult =
  | { readonly ok: true }
  | {
      readonly ok: false;
      readonly reason:
        | 'document-not-ready'
        | 'invalid-owner'
        | 'invalid-summary'
        | 'invalid-placeholder-height'
        | 'unsupported-content'
        | 'limit-exceeded';
    };
