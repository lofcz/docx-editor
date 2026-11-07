/** One locally collapsed block content control and its visible summary. @public */
export interface CollapsedContentControl {
  /** Canonical content-control node ID in the main document story. */
  readonly controlId: string;
  /** First direct paragraph ID, or null for a blank local header hiding the whole body. */
  readonly summaryParagraphId: string | null;
}

/** Validation outcome for a complete replacement of the local collapse view. @public */
export type ContentControlCollapseResult =
  | { readonly ok: true }
  | {
      readonly ok: false;
      readonly reason:
        | 'document-not-ready'
        | 'invalid-owner'
        | 'invalid-summary'
        | 'unsupported-content'
        | 'limit-exceeded';
    };
