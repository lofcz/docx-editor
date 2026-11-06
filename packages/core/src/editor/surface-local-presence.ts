import type {
  CollaborationLocalSelection,
  EditorCollaborationSession,
} from '../collaboration/index.ts';
import { hostWindowOf } from './host-realm.ts';
import type { SurfaceHostListeners } from './surface-host-listeners.ts';

/** A local draft has no canonical selection until its owner commits it. */
const DRAFT_ATTRIBUTE = 'data-docx-local-draft';

/** Publish through one cache, including temporary withdrawal for explicitly owned local drafts. */
export function createLocalCollaborationPresence(deps: {
  readonly session: EditorCollaborationSession | undefined;
  readonly listeners: SurfaceHostListeners;
  host(): HTMLElement;
  read(): CollaborationLocalSelection | null;
}) {
  let published: string | null = null;
  let disposed = false;
  let observer: MutationObserver | null = null;
  let observedDocument: Document | null = null;
  const ownsDraft = (): boolean => {
    const host = deps.host();
    const marker = host.ownerDocument.activeElement?.closest(`[${DRAFT_ATTRIBUTE}]`);
    if (!marker?.isConnected) return false;
    const owner = marker.getAttribute(DRAFT_ATTRIBUTE);
    return owner === '' ? host.contains(marker) : owner === deps.session?.sessionId;
  };
  function observeDraft(active: boolean) {
    const document = deps.host().ownerDocument;
    if (!active || observedDocument !== document) {
      observer?.disconnect();
      observer = null;
      observedDocument = null;
    }
    if (!active || observer) return;
    // Removal of a focused field need not dispatch blur. Observe only while a draft owns focus.
    observer = new (hostWindowOf(deps.host()).MutationObserver)(publish);
    observer.observe(document.documentElement, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: [DRAFT_ATTRIBUTE],
    });
    observedDocument = document;
  }
  function publish(): void {
    if (disposed || !deps.session) return;
    const draft = ownsDraft();
    observeDraft(draft);
    const next = draft ? null : deps.read();
    const key = next
      ? `${next.anchor.paragraphId}:${next.anchor.offset}|${next.head.paragraphId}:${next.head.offset}|${next.kind ?? ''}`
      : '';
    if (published === key) return;
    published = key;
    deps.session.setLocalSelection(next);
  }
  // The registry rebinds these listeners when the same surface moves to another window.
  const offFocus = deps.session ? deps.listeners.onDocument('focusin', publish, true) : () => {};
  const offBlur = deps.session
    ? deps.listeners.onDocument('focusout', () => queueMicrotask(publish), true)
    : () => {};
  return {
    publish,
    destroy() {
      disposed = true;
      offFocus();
      offBlur();
      observer?.disconnect();
      observer = null;
      observedDocument = null;
    },
  };
}
