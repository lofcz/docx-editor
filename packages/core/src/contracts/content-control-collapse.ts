/** One locally collapsed block content control and its visible summary. @public */
export interface CollapsedContentControl {
  /** Canonical content-control node ID in the main document story. */
  readonly controlId: string;
  /** Canonical paragraph ID of the control's first direct content child. */
  readonly summaryParagraphId: string;
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
