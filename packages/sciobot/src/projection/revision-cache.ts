import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  sameRevision,
  type DocumentIdentity,
  type DocumentView,
  type RevisionToken,
} from '../contracts/document';

export interface ProjectionCacheKey {
  readonly identity: DocumentIdentity;
  readonly revision: RevisionToken;
  readonly view: DocumentView;
}

/** One document revision, with at most its two read views. It owns no editor state. */
export class RevisionProjectionCache<Value> {
  private key?: Omit<ProjectionCacheKey, 'view'>;
  private readonly values = new Map<DocumentView, Value>();

  get(key: ProjectionCacheKey, build: () => Value): Value {
    if (key.identity.id !== key.revision.documentId)
      throw new TypeError('Projection document identity does not match its revision');
    const previous = this.key;
    if (
      !previous ||
      previous.identity.id !== key.identity.id ||
      previous.identity.path !== key.identity.path ||
      previous.identity.kind !== key.identity.kind ||
      !sameRevision(previous.revision, key.revision)
    ) {
      this.values.clear();
      this.key = { identity: { ...key.identity }, revision: { ...key.revision } };
    }
    if (this.values.has(key.view)) return this.values.get(key.view)!;
    const value = build();
    this.values.set(key.view, value);
    return value;
  }

  invalidate(): void {
    this.key = undefined;
    this.values.clear();
  }
}

/** Settle user typing before reading the package revision, as the native host does. */
export function readEditorRevision(
  editor: DocxEditorInstance,
  documentId: string,
  epoch: string
): RevisionToken | null {
  const surface = editor.surface;
  if (!surface) return null;
  surface.flushPendingInput();
  return { documentId, epoch, revision: surface.session.packageRevision() };
}
