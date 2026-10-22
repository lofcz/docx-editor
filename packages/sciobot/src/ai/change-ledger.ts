import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  contentControlPropertiesOf,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { blockIdFromTag } from '../blocks/payload-opc-codec';

/**
 * - `new`: highlighted until acknowledged.
 * - `seen`: acknowledged; fading out, then forgotten.
 */
export type DirectChangeStatus = 'new' | 'seen';

/** One AI-authored change relative to the request-session baseline. */
export interface DirectChange {
  readonly id: string;
  /** Persistent registered block that owns the change, or null for an ordinary body block. */
  readonly ownerBlockId: string | null;
  /** Native story node that groups the change: the registered control or top-level block. */
  readonly ownerNodeId: string;
  readonly kind: 'insert' | 'text';
  readonly status: DirectChangeStatus;
  /** Hidden while edit playback has not reached it yet. */
  readonly pending: boolean;
  /** Current native paragraph and table IDs that were inserted or changed, in document order. */
  readonly nodeIds: readonly string[];
}

/** How long an acknowledged change fades before it is forgotten. */
export const DIRECT_CHANGE_FADE = 1200;

interface Leaf {
  readonly id: string;
  readonly fingerprint: string;
  readonly ownerNodeId: string;
  readonly ownerBlockId: string | null;
}
interface Entry {
  id: string;
  ownerNodeId: string;
  ownerBlockId: string | null;
  kind: DirectChange['kind'];
  status: DirectChangeStatus;
  pending: boolean;
  nodeIds: string[];
  /** Fingerprints when last computed; a user write to one of them drops the change. */
  prints: Map<string, string>;
}

/** Paragraphs are leaves; tables recurse to their cell paragraphs. Ownership is the nearest registered control. */
function leaves(part: OoxmlPart): Leaf[] {
  const root = bodyStoryRoot(part);
  if (!root || root.kind === 'textValue') return [];
  const out: Leaf[] = [];
  const visit = (node: OoxmlNode, top: string, owner: { id: string; block: string } | null) => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'paragraph') {
      out.push({
        id: node.id,
        fingerprint: canonicalOoxmlFingerprint(node),
        ownerNodeId: owner?.id ?? top,
        ownerBlockId: owner?.block ?? null,
      });
      return;
    }
    let next = owner;
    if (node.kind === 'contentControl') {
      const tag = contentControlPropertiesOf(node).tag,
        block = tag ? blockIdFromTag(tag) : null;
      if (block) next = { id: node.id, block };
    }
    for (const child of node.children) visit(child, top, next);
  };
  for (const child of root.children) visit(child, child.id, null);
  return out;
}

/**
 * Per-editor record of content a direct AI edit wrote, highlighted until acknowledged.
 *
 * Within one request session, changes are recomputed against the session baseline, so repeated
 * edits of one paragraph or one registered block merge into one change. Acknowledged changes fold
 * into the baseline. A non-AI commit (typing, undo, review decision, remote revision) seals the
 * baseline and drops changes it touched. A new session settles all earlier changes. The ledger
 * reads canonical fingerprints only; it never accepts or rejects Word review items.
 */
export class DirectChangeLedger {
  private entries: Entry[] = [];
  private baseline: Map<string, string> | undefined;
  private session: string | undefined;
  private expectedRevision: number | null = null;
  private sequence = 0;
  private muted = 0;
  private version = 0;
  private readonly fades = new Map<string, ReturnType<typeof setTimeout>>();
  private readonly listeners = new Set<() => void>();
  private readonly stop: () => void;

  constructor(
    private readonly editor: DocxEditorInstance,
    private readonly fadeMs = DIRECT_CHANGE_FADE
  ) {
    const session = editor.surface?.session;
    if (!session) throw new Error('The native editor is not attached.');
    this.stop = session.subscribe(() => this.onCommit());
  }

  get changes(): readonly DirectChange[] {
    return this.entries.map((entry) =>
      Object.freeze({
        id: entry.id,
        ownerBlockId: entry.ownerBlockId,
        ownerNodeId: entry.ownerNodeId,
        kind: entry.kind,
        status: entry.status,
        pending: entry.pending,
        nodeIds: Object.freeze([...entry.nodeIds]),
      })
    );
  }
  getVersion = () => this.version;
  subscribe = (listener: () => void) => {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  };

  /** Run AI-authored writes that must not count as user edits; call `record` afterwards. */
  mute<T>(write: () => T): T {
    this.muted++;
    try {
      return write();
    } finally {
      this.muted--;
    }
  }

  /**
   * Record one committed direct edit. `before` is the canonical part read right before the
   * commit; a different `session` first settles every earlier change.
   */
  record(before: OoxmlPart, options: { session: string; pending?: boolean }) {
    if (!options.session || options.session !== this.session) {
      this.clear();
      this.session = options.session;
    }
    if (!this.baseline)
      this.baseline = new Map(leaves(before).map((leaf) => [leaf.id, leaf.fingerprint]));
    this.expectedRevision = this.surface().session.packageRevision();
    this.recompute(options.pending ?? false);
  }

  /** Reveal pending changes for the given native nodes, or all of them. */
  reveal(nodeIds?: readonly string[]) {
    let changed = false;
    for (const entry of this.entries)
      if (entry.pending && (!nodeIds || entry.nodeIds.some((id) => nodeIds.includes(id)))) {
        entry.pending = false;
        changed = true;
      }
    if (changed) this.notify();
  }

  /** Fade a change out, then forget it. */
  acknowledge(id: string) {
    const entry = this.entries.find((item) => item.id === id);
    if (!entry || entry.status === 'seen') return;
    entry.status = 'seen';
    this.fades.set(
      id,
      setTimeout(() => {
        this.fades.delete(id);
        this.settle(id);
      }, this.fadeMs)
    );
    this.notify();
  }

  /** Forget one change and keep its content, folding it into the session baseline. */
  settle(id: string) {
    const entry = this.entries.find((item) => item.id === id);
    if (!entry) return;
    this.stopFade(id);
    if (this.baseline) {
      for (const [nodeId, print] of entry.prints) this.baseline.set(nodeId, print);
      this.entries = this.entries.filter((item) => item !== entry);
      this.recompute(false);
      return;
    }
    this.entries = this.entries.filter((item) => item !== entry);
    this.notify();
  }

  /** Forget every change and keep content. Review items are never decided here. */
  settleAll() {
    this.clear();
    this.notify();
  }

  dispose() {
    this.stop();
    this.clear();
    this.listeners.clear();
  }

  private surface() {
    const surface = this.editor.surface;
    if (!surface) throw new Error('The native editor is not attached.');
    return surface;
  }

  private onCommit() {
    if (this.muted) return;
    const revision = this.surface().session.packageRevision();
    if (revision === this.expectedRevision) return;
    // A user, undo, review, or remote commit cannot be expressed against the request baseline.
    this.expectedRevision = null;
    this.baseline = undefined;
    if (!this.entries.length) return;
    const current = new Map(
      leaves(this.surface().session.part()).map((leaf) => [leaf.id, leaf.fingerprint])
    );
    const touched = this.entries.filter((entry) =>
      [...entry.prints].some(([id, print]) => current.get(id) !== print)
    );
    if (!touched.length) return;
    for (const entry of touched) this.stopFade(entry.id);
    this.entries = this.entries.filter((entry) => !touched.includes(entry));
    this.notify();
  }

  private recompute(pending: boolean) {
    const baseline = this.baseline;
    if (!baseline) return;
    const groups = new Map<string, { leaves: Leaf[]; inserted: boolean }>();
    for (const leaf of leaves(this.surface().session.part())) {
      const old = baseline.get(leaf.id);
      if (old === leaf.fingerprint) continue;
      const group = groups.get(leaf.ownerNodeId) ?? { leaves: [], inserted: true };
      group.leaves.push(leaf);
      group.inserted &&= old === undefined;
      groups.set(leaf.ownerNodeId, group);
    }
    const previous = new Map(this.entries.map((entry) => [entry.ownerNodeId, entry]));
    const next: Entry[] = [];
    for (const [ownerNodeId, group] of groups) {
      const prints = new Map(group.leaves.map((leaf) => [leaf.id, leaf.fingerprint])),
        prior = previous.get(ownerNodeId),
        same =
          !!prior &&
          prior.prints.size === prints.size &&
          [...prints].every(([id, print]) => prior.prints.get(id) === print);
      if (prior && !same) this.stopFade(prior.id);
      next.push({
        id: prior?.id ?? `ai-change-${++this.sequence}`,
        ownerNodeId,
        ownerBlockId: group.leaves[0]!.ownerBlockId,
        kind: group.inserted ? 'insert' : 'text',
        status: same ? prior!.status : 'new',
        pending: same ? prior!.pending : pending,
        nodeIds: group.leaves.map((leaf) => leaf.id),
        prints,
      });
    }
    for (const entry of this.entries)
      if (!next.some((item) => item.id === entry.id)) this.stopFade(entry.id);
    this.entries = next;
    this.notify();
  }

  private clear() {
    for (const entry of this.entries) this.stopFade(entry.id);
    this.entries = [];
    this.baseline = undefined;
    this.expectedRevision = null;
  }

  private stopFade(id: string) {
    clearTimeout(this.fades.get(id));
    this.fades.delete(id);
  }

  private notify() {
    this.version++;
    for (const listener of this.listeners) listener();
  }
}
