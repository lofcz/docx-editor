import {
  createPropertyReviewPrototype,
  PROPERTY_REVIEW_PROTOTYPE,
  type PropertyReviewPrototype,
} from './property-review-prototype';
import type { ReviewFieldReceipt } from './review-fields';
import type { ReviewFlagReceipt } from './review-flag-codec';
import { encodeStructuralPolicy } from './owned-structure';
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
// A local wrapper capability; it never enters the shared ownership protocol.
const NATIVE_UNDO_AUTHORITY = Symbol.for('@docx-editor.dev/undo-history-authority');

export interface OwnedReviewHistory extends Omit<OwnedTextHandle, 'setStructuralAttributes'> {
  /** Bind the session created over this scope's document. All native methods keep their receiver. */
  bindSession(session: EditorCollaborationSession): EditorCollaborationSession;
}
export interface OwnedReviewHistoryOptions extends OwnedTextOptions {
  readonly documentId: string;
  readonly bootstrap: 'create' | 'join';
  /** Obtain the provider's advertised native structural keys before room bootstrap. */
  readonly structuralTextAttributeKeys: readonly string[];
}

/** Attach before bootstrap or room snapshot loading. Dispose after detaching/destroying the session. */
export function attachOwnedReviewHistory(
  doc: Y.Doc,
  options: OwnedReviewHistoryOptions
): OwnedReviewHistory {
  return createReviewHistory(doc, options, false);
}
/** Internal proof entry. It is absent from package exports and production attachment. */
export function attachPropertyReviewHistoryPrototype(
  doc: Y.Doc,
  options: OwnedReviewHistoryOptions
): OwnedReviewHistory {
  return createReviewHistory(doc, options, true);
}
function createReviewHistory(
  doc: Y.Doc,
  options: OwnedReviewHistoryOptions,
  experimentalProperties: boolean
): OwnedReviewHistory {
  if (
    !options ||
    typeof options.documentId !== 'string' ||
    !options.documentId.trim() ||
    options.documentId.length > 4096 ||
    /[\u0000-\u001f]/.test(options.documentId)
  )
    throw new TypeError('Owned review history needs the current room document ID.');
  if (options.bootstrap !== 'create' && options.bootstrap !== 'join')
    throw new TypeError('Owned review history needs a create or join bootstrap mode.');
  encodeStructuralPolicy(options.structuralTextAttributeKeys);
  const documentId = options.documentId,
    owned = attachOwnedTextAttributes(doc, {
      onError: options.onError,
      protocol: 2,
      bootstrap: options.bootstrap,
    }),
    sessions = new WeakMap<EditorCollaborationSession, EditorCollaborationSession>();
  try {
    owned.setStructuralAttributes(options.structuralTextAttributeKeys);
  } catch (cause) {
    owned.dispose();
    throw cause;
  }
  let flags: OwnedReviewFlags;
  try {
    flags = attachOwnedReviewFlags(doc, {
      onError: options.onError,
      protocol: experimentalProperties ? '3' : '2',
      bootstrap: options.bootstrap,
    });
  } catch (cause) {
    owned.dispose();
    throw cause;
  }
  let fields: ReturnType<typeof attachOwnedReviewFields>;
  try {
    fields = attachOwnedReviewFields(doc, {
      onError: options.onError,
      protocol: experimentalProperties ? '2' : '1',
    });
  } catch (cause) {
    flags.dispose();
    owned.dispose();
    throw cause;
  }
  const properties = experimentalProperties
    ? createPropertyReviewPrototype(doc, fields, flags)
    : null;
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
        if (!properties) return flags.capture(() => fields.capture(run, fieldWrites), flagWrites);
        let fieldReceipt: readonly ReviewFieldReceipt[] = [],
          flagReceipt: readonly ReviewFlagReceipt[] = [];
        const value = flags.capture(
          () =>
            fields.capture(run, fieldWrites, (items) => {
              fieldReceipt = items;
            }),
          flagWrites,
          (items) => {
            flagReceipt = items;
          }
        );
        properties.note({ fields: fieldReceipt, flags: flagReceipt });
        return value;
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
      const previous = sessions.get(session);
      if (previous) return previous;
      const attached = (session as unknown as Record<symbol, unknown>)[REVIEW_HISTORY_CAPABILITY];
      if (attached !== undefined && attached !== assertHealthy)
        throw new TypeError('The session already belongs to another review history scope.');
      if (!Array.isArray(session.structuralTextAttributeKeys))
        throw new TypeError('Owned history requires a provider structural text attribute policy.');
      owned.setStructuralAttributes(session.structuralTextAttributeKeys);
      fields.enable();
      const methods = new Map<PropertyKey, { value: unknown; bound: unknown }>();
      // A separate extensible facade also supports frozen transport objects and private receivers.
      const bound = new Proxy(
        Object.create(Object.getPrototypeOf(session)) as EditorCollaborationSession,
        {
          get(_target, key) {
            if (key === PROPERTY_REVIEW_PROTOTYPE)
              return properties as PropertyReviewPrototype | null;
            if (key === NATIVE_UNDO_AUTHORITY) return session;
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
