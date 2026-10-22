import type * as Y from 'yjs';
import type { EditorCollaborationSession } from '@docx-editor.dev/core/collaboration';
import { REVIEW_HISTORY_CAPABILITY } from '../review/group-history';
import {
  attachOwnedTextAttributes,
  type OwnedTextHandle,
  type OwnedTextOptions,
} from './owned-text';
import { attachOwnedReviewFlags, type OwnedReviewFlags } from './review-flags';
import { attachOwnedReviewFields } from './review-fields';

export interface OwnedReviewHistory extends OwnedTextHandle {
  /** Bind the session created over this scope's document. All native methods keep their receiver. */
  bindSession(session: EditorCollaborationSession): EditorCollaborationSession;
}
export interface OwnedReviewHistoryOptions extends OwnedTextOptions {
  readonly documentId: string;
}

/** Attach before bootstrap or room snapshot loading. Dispose after detaching/destroying the session. */
export function attachOwnedReviewHistory(
  doc: Y.Doc,
  options: OwnedReviewHistoryOptions
): OwnedReviewHistory {
  if (
    !options ||
    typeof options.documentId !== 'string' ||
    !options.documentId.trim() ||
    options.documentId.length > 4096 ||
    /[\u0000-\u001f]/.test(options.documentId)
  )
    throw new TypeError('Owned review history needs the current room document ID.');
  const documentId = options.documentId,
    owned = attachOwnedTextAttributes(doc, { onError: options.onError }),
    sessions = new WeakMap<EditorCollaborationSession, EditorCollaborationSession>();
  let flags: OwnedReviewFlags;
  try {
    flags = attachOwnedReviewFlags(doc, { onError: options.onError, protocol: '2' });
  } catch (cause) {
    owned.dispose();
    throw cause;
  }
  let fields: ReturnType<typeof attachOwnedReviewFields>;
  try {
    fields = attachOwnedReviewFields(doc, { onError: options.onError });
  } catch (cause) {
    flags.dispose();
    owned.dispose();
    throw cause;
  }
  let disposed = false;
  const assertHealthy = Object.assign(
    () => {
      if (disposed || doc.isDestroyed)
        throw new Error('The owned review history scope is no longer active.');
      owned.assertHealthy();
      flags.assertHealthy();
      fields.assertHealthy();
    },
    {
      capture: <T>(run: () => T, flagWrites: number, fieldWrites = 0): T => {
        assertHealthy();
        return flags.capture(() => fields.capture(run, fieldWrites), flagWrites);
      },
    }
  );
  return Object.freeze({
    protocol: owned.protocol,
    assertHealthy,
    bindSession(session: EditorCollaborationSession) {
      assertHealthy();
      if (session.documentId !== documentId)
        throw new TypeError('Review history document identity does not match the session.');
      fields.enable();
      const previous = sessions.get(session);
      if (previous) return previous;
      const attached = (session as unknown as Record<symbol, unknown>)[REVIEW_HISTORY_CAPABILITY];
      if (attached !== undefined && attached !== assertHealthy)
        throw new TypeError('The session already belongs to another review history scope.');
      const methods = new Map<PropertyKey, { value: unknown; bound: unknown }>();
      // A separate extensible facade also supports frozen transport objects and private receivers.
      const bound = new Proxy(
        Object.create(Object.getPrototypeOf(session)) as EditorCollaborationSession,
        {
          get(_target, key) {
            if (key === REVIEW_HISTORY_CAPABILITY) return assertHealthy;
            const value = Reflect.get(session, key, session);
            if (typeof value !== 'function') return value;
            const cached = methods.get(key);
            if (cached && cached.value === value) return cached.bound;
            const method = value.bind(session);
            methods.set(key, { value, bound: method });
            return method;
          },
        }
      );
      sessions.set(session, bound);
      sessions.set(bound, bound);
      return bound;
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      fields.dispose();
      flags.dispose();
      owned.dispose();
    },
  });
}
