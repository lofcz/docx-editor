/** Stable application identity, separate from a mounted editor instance. */
export interface DocumentIdentity {
  readonly id: string;
  readonly path: string;
  readonly kind: 'document' | 'lesson_plan' | 'curriculum';
}

/** The epoch changes whenever document bytes replace a live editor instance. */
export interface RevisionToken {
  readonly documentId: string;
  readonly epoch: string;
  readonly revision: number;
}

export type DocumentView = 'committed' | 'working';
export type EditMode = 'direct' | 'suggestion';

export function sameRevision(left: RevisionToken, right: RevisionToken): boolean {
  return (
    left.documentId === right.documentId &&
    left.epoch === right.epoch &&
    left.revision === right.revision
  );
}

export type EditFailureCode =
  | 'stale-revision'
  | 'unsupported-operation'
  | 'invalid-input'
  | 'ambiguous-target'
  | 'invalid-document'
  | 'document-not-ready'
  | 'apply-failed';

/** Internal failures retain codes. The legacy tool adapter formats their error. */
export interface NativeEditFailure {
  readonly success: false;
  readonly code: EditFailureCode;
  readonly error: string;
  readonly expectedRevision?: RevisionToken;
  readonly actualRevision?: RevisionToken;
}
