// The growing prefix a large open lays out in separate tasks (`bodyBlockLimit`).
//
// Each pass lays out the first N body blocks and resumes from the last pass through the
// session's checkpoints, so later passes only place the blocks the limit added. The count
// recorded here tells the caller when a limit no longer cuts anything.

import type { LayoutSession } from './layout-session.ts';

/** A sliced open's cut on a session: the limit it took, and the first block it left out. */
interface Cut {
  readonly limit: number;
  readonly firstOutId: string | undefined;
}
interface BodyBlockPass {
  readonly all: readonly unknown[];
  readonly cut: Cut | undefined;
}
const passes = new WeakMap<LayoutSession, BodyBlockPass>();

/** Put back the visible pass's cut when the host discards an unpublished layout. */
export function snapshotBodyBlockPass(session: LayoutSession): () => void {
  const previous = passes.get(session);
  return () => {
    if (previous) passes.set(session, previous);
    else passes.delete(session);
  };
}

function idOf(block: unknown): string | undefined {
  return typeof block === 'object' && block !== null && 'id' in block
    ? String((block as { readonly id: unknown }).id)
    : undefined;
}

function indexOfId(blocks: readonly unknown[], id: string): number {
  for (let index = 0; index < blocks.length; index += 1) {
    if (idOf(blocks[index]) === id) return index;
  }
  return -1;
}

/**
 * `all` cut to `bodyBlockLimit`, recording the uncut count on the session. List numbering
 * still resolves over `all`: a number depends only on the paragraphs before it, so the items
 * are the same, and one resolution serves every prefix instead of a new one per pass.
 *
 * A pass asking for the same limit as the last one keeps the same CUT: before the block the
 * last pass left out, wherever edits since moved it. A paragraph split or pasted before the cut
 * stays laid out, and a deletion after it leaves the cut where it was. When that block is gone,
 * the cut moves by the change in the block count instead.
 */
export function limitBodyBlocks<T>(
  all: readonly T[],
  options: { readonly bodyBlockLimit?: number; readonly session?: LayoutSession }
): { readonly blocks: readonly T[]; readonly all: readonly T[] } {
  const session = options.session;
  const requested = options.bodyBlockLimit;
  let limit = requested;
  const previous = session ? passes.get(session) : undefined;
  const cut = previous?.cut;
  if (requested !== undefined && cut?.limit === requested && previous) {
    const at = cut.firstOutId === undefined ? all.length : indexOfId(all, cut.firstOutId);
    limit = at >= 0 ? at : Math.max(1, requested + all.length - previous.all.length);
  }
  if (session) {
    passes.set(session, {
      all,
      cut: limit === undefined ? cut : { limit, firstOutId: idOf(all[limit]) },
    });
  }
  return { blocks: limit === undefined || limit >= all.length ? all : all.slice(0, limit), all };
}

/**
 * The cut a sliced open keeps after a pass on `session`: `current` as the edits that pass saw
 * moved it, so the next pass cuts at the same place. Null stays null: the open has finished.
 */
export function bodyBlockLimitAfter(session: LayoutSession, current: number | null): number | null {
  return current === null ? null : (passes.get(session)?.cut?.limit ?? current);
}

/** The body block count the last pass on `session` saw, or undefined before any pass. */
export function bodyBlockCountOf(session: LayoutSession): number | undefined {
  return passes.get(session)?.all.length;
}

/** The body blocks the last pass on `session` cut its prefix from, or undefined before any. */
export function bodyBlocksOf(session: LayoutSession): readonly unknown[] | undefined {
  return passes.get(session)?.all;
}
