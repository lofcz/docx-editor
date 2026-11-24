// The internal seam a large open uses to finish its first layout in slices.
//
// The surface mounts with a layout of only the first body blocks and takes input from then
// on. Each step lays out more of the body, resuming from the previous pass, until the limit no
// longer cuts anything; the last step publishes the normal full layout. Kept off the public
// surface contract, like `surface-measurement.ts`.

import { drainLayoutSteps, type LayoutSteps } from '../layout/layout-steps.ts';
import { bodyBlockCountOf, bodyBlocksOf } from '../layout/body-block-limit.ts';
import type { LayoutSession } from '../layout/layout-session.ts';
import { snapshotLayoutSession } from '../layout/layout-session-snapshot.ts';
import { caretAt } from '../layout/semantic-interaction.ts';
import type { SemanticPosition } from '../layout/semantic-interaction.ts';
import type { SemanticLayout } from '../layout/semantic-records.ts';
import { warmLayoutNodeDigests } from '../layout/layout-cache.ts';
import type { OoxmlNode } from '@docx-editor.dev/core/store';
import { showOpeningIndicator } from './surface-opening-indicator.ts';
import type { PaginatedSurfaceRuntimeOptions } from './surface-runtime-options.ts';

/** Body blocks a sliced opening lays out at mount, before its first yield. */
export const PROGRESSIVE_OPEN_FIRST_BLOCKS = 64;

/** A surface's opening step; `force` runs it even while an IME composition is open. */
type OpeningStep = (budgetMs: number, force?: boolean) => boolean;
interface OpeningWork {
  run: OpeningStep;
  shortPrefix: boolean;
  atEnd?: boolean;
  wake?: () => void;
}
const steps = new WeakMap<object, OpeningWork>();

/** A reveal waiting for the opening to lay out its target; see {@link revealWhenLaidOut}. */
interface PendingReveal {
  covers(layout: unknown): boolean;
  /** Runs the reveal, unless the reader scrolled away since asking (scrollbar, keyboard). */
  reveal(): void;
}
const pendingReveals = new WeakMap<object, PendingReveal>();

/**
 * The slice length while a reveal waits. Longer slices finish sooner, and a slice still ends
 * early for input, so the page stays responsive while the reader waits for the target.
 */
const REVEAL_SLICE_MS = 48;
let stepCount = 0;

/** @internal How many opening steps have run, for tests that must see the slicing. */
export function progressiveOpenStepCount(): number {
  return stepCount;
}

/** Register `surface`'s opening step. */
function registerProgressiveOpen(surface: object, step: OpeningStep, shortPrefix: boolean): void {
  steps.set(surface, { run: step, shortPrefix });
}

/** A prepared paragraph prefix can yield to input until the reader requests a later page. */
export function openingUsesIdleTime(surface: object): boolean {
  const work = steps.get(surface);
  return work?.shortPrefix === true && !work.atEnd && !pendingReveals.has(surface);
}

/** A reader can ask for more pages while an idle callback is waiting for the browser. */
export function onOpeningDemand(surface: object, wake: () => void): void {
  const work = steps.get(surface);
  if (work?.shortPrefix) work.wake = wake;
}

/** Scrolling to the bottom of the visible prefix requests its next pages. */
export function openingViewportMoved(surface: object, scroller: HTMLElement): void {
  const work = steps.get(surface);
  if (!work?.shortPrefix) return;
  const atEnd =
    scroller.scrollTop > 0 &&
    scroller.scrollHeight - scroller.scrollTop - scroller.clientHeight <= 1;
  if (work.atEnd === atEnd) return;
  work.atEnd = atEnd;
  if (atEnd) work.wake?.();
}

/**
 * Advance `surface`'s opening layout for about `budgetMs`. True once the full layout is
 * published, and for a surface that opened without slicing.
 */
export function continueProgressiveOpen(surface: object, budgetMs: number): boolean {
  const step = steps.get(surface);
  if (!step) return true;
  stepCount += 1;
  const done = step.run(budgetMs);
  if (done) steps.delete(surface);
  return done;
}

/** Whether `surface` is still laying out its opening slices. */
export function progressiveOpenPending(surface: object): boolean {
  return steps.has(surface);
}

/**
 * The body text one prefix step adds, in characters, and its bounds in blocks.
 *
 * Sized by CONTENT, never by the time a pass took. Each prefix is its own footnote-reserve
 * problem and seeds the next one, so a schedule that followed the clock published a different
 * pagination on a slower or busier machine. Content is the same on every machine and tracks the
 * work closely enough: a footnoted body lays a step out in one task, so the step bounds the
 * longest pause, and every prefix also walks the blocks before it, so steps must not be tiny.
 */
const PREFIX_STEP_CHARACTERS = 60_000;

/**
 * The shortest time between two shown prefixes. Showing one repaints the visible pages and the
 * comment highlights; doing it after every step took a fifth of the opening. Only what the
 * screen shows follows the clock: every prefix pass still runs, so the layout is the same.
 */
const PREFIX_PUBLISH_INTERVAL_MS = 300;
const MIN_PREFIX_STEP = 16;
const MAX_PREFIX_STEP = 4096;

/**
 * Count paragraphs in a prepared w:body (CT_Body, §17.2.2). Each w:p is a CT_P (§17.3.1.22).
 * Final w:sectPr (CT_SectPr, §17.6.17) follows the paragraphs and does not add a body block.
 */
export function preparedParagraphPrefix(root: OoxmlNode): number | null {
  if (root.kind !== 'document') return null;
  const body = root.children.find((node) => node.kind === 'body');
  if (!body || body.children.length <= PROGRESSIVE_OPEN_FIRST_BLOCKS) return null;
  const last = body.children.at(-1)!;
  const hasSection = last.namespaceUri === body.namespaceUri && last.localName === 'sectPr';
  const count = body.children.length - Number(hasSection);
  if (count <= PROGRESSIVE_OPEN_FIRST_BLOCKS) return null;
  for (let index = 0; index < count; index += 1) {
    if (body.children[index]!.kind !== 'paragraph') return null;
  }
  return MIN_PREFIX_STEP;
}

/**
 * The body block limit after `limit`: at least {@link MIN_PREFIX_STEP} more blocks, then more
 * until the step holds {@link PREFIX_STEP_CHARACTERS} of text, at most {@link MAX_PREFIX_STEP}.
 * A function of the content alone.
 */
export function nextPrefixLimit(
  limit: number,
  total: number,
  blockAt: (index: number) => unknown
): number {
  let characters = 0;
  let next = limit;
  while (
    next < total &&
    next - limit < MAX_PREFIX_STEP &&
    (next - limit < MIN_PREFIX_STEP || characters < PREFIX_STEP_CHARACTERS)
  ) {
    characters += bodyBlockWeight(blockAt(next));
    next += 1;
  }
  return next;
}

/**
 * The tables among body blocks `from` to `to`, rows first: a table's own digest then only joins
 * its rows'. Hashed before the prefix pass that lays them out, in slices.
 */
function tableNodesBetween(
  from: number,
  to: number,
  blockAt: (index: number) => unknown
): OoxmlNode[] {
  const nodes: OoxmlNode[] = [];
  for (let index = from; index < to; index += 1) {
    const block = blockAt(index) as OoxmlNode | undefined;
    if (block?.kind !== 'table') continue;
    for (const row of block.children) nodes.push(row);
    nodes.push(block);
  }
  return nodes;
}

const blockWeights = new WeakMap<object, number>();

/** A body block's text length, the measure a prefix step is sized by. */
export function bodyBlockWeight(block: unknown): number {
  if (typeof block !== 'object' || block === null) return 1;
  const cached = blockWeights.get(block);
  if (cached !== undefined) return cached;
  let weight = 1;
  const stack: unknown[] = [block];
  while (stack.length > 0) {
    const node = stack.pop() as { kind?: string; value?: string; children?: readonly unknown[] };
    if (node.kind === 'textValue') weight += node.value?.length ?? 0;
    else if (node.children) for (const child of node.children) stack.push(child);
  }
  blockWeights.set(block, weight);
  return weight;
}

/** What the slices need from the surface that is opening. */
export interface OpeningSlicesHost {
  now(): number;
  destroyed(): boolean;
  /** The body block limit of the opening, or null once a full layout ran. */
  limit(): number | null;
  setLimit(limit: number): void;
  bodyBlockCount(): number | undefined;
  /** A prefix layout to `limit` body blocks, as steps that may pause inside a long table. */
  prefixPass(limit: number): LayoutSteps<unknown>;
  /** Save the layout session that produced the visible prefix, before extending it. */
  snapshot(): () => void;
  /** The normal full layout, resumed from the last prefix; it ends the opening. */
  finish(): void;
  /** The body block at `index` of the last pass, for sizing the next step. */
  bodyBlock(index: number): unknown;
  /** Show a finished prefix layout, so the pages grow while the reader works. */
  publish(layout: unknown): void;
  /** The pages layer: it takes the reader's input, and moves with the surface. */
  pages(): HTMLElement;
  /** Whether input waits; a slice then ends early so the input runs first. */
  inputPending(): boolean;
  /** Whether an edit's pass is still finishing the prefix; the next step waits for it. */
  busy(): boolean;
  /** Finish that pass now, for a step that may not wait. */
  finishBusy(): void;
  /**
   * Whether an IME composition is open. The slices wait: a shown prefix repaints the pages,
   * which would replace the text the composition is writing into.
   */
  composing(): boolean;
}

/** Own the opening limit and its continuations from the first layout until the surface closes. */
export function createSurfaceOpening(
  options: Pick<PaginatedSurfaceRuntimeOptions, 'progressiveOpen' | 'openedSession'>,
  root: () => OoxmlNode,
  session: () => LayoutSession
) {
  const first = options.progressiveOpen
    ? PROGRESSIVE_OPEN_FIRST_BLOCKS
    : options.openedSession
      ? preparedParagraphPrefix(root())
      : null;
  let slices: ReturnType<typeof startOpeningSlices> | undefined;
  const opening = {
    limit: first,
    finishPaused: () => slices?.finishPaused(),
    abandonForEdit: () => slices?.abandonForEdit(),
    stop: () => slices?.stop(),
    start(
      surface: object,
      container: () => HTMLElement,
      host: Omit<
        OpeningSlicesHost,
        'limit' | 'setLimit' | 'bodyBlockCount' | 'bodyBlock' | 'snapshot'
      >
    ): void {
      slices = startOpeningSlices(
        surface,
        container,
        {
          ...host,
          limit: () => opening.limit,
          setLimit: (limit) => (opening.limit = limit),
          bodyBlockCount: () => bodyBlockCountOf(session()),
          bodyBlock: (index) => bodyBlocksOf(session())?.[index],
          snapshot: () => snapshotLayoutSession(session()),
        },
        first ?? PROGRESSIVE_OPEN_FIRST_BLOCKS
      );
    },
  };
  return opening;
}

/**
 * Lay a mounted surface's body out in growing prefixes (see {@link PREFIX_STEP_CHARACTERS}):
 * each extends the prefix and resumes from the previous one, and a prefix pass pauses between the
 * rows of a long table instead of laying it out in one task. Marks `container` while it runs,
 * so the loading overlay steps aside, and shows the loading indicator below the laid-out pages.
 *
 * Answers `finishPaused`, which finishes a paused prefix pass at once: any other layout of the
 * same session must call it first, because the paused pass holds that session mid-flow. And
 * `stop`, which a destroyed surface calls to drop its opening state.
 */
export function startOpeningSlices(
  surface: object,
  container: () => HTMLElement,
  host: OpeningSlicesHost,
  firstGrowth: number
): { finishPaused(): void; abandonForEdit(): void; stop(): void } {
  let growth = firstGrowth;
  const shortPrefix = firstGrowth < PROGRESSIVE_OPEN_FIRST_BLOCKS;
  let publishedAt = host.now();
  let pass: LayoutSteps<unknown> | null = null;
  let restorePublished: (() => void) | null = null;
  /** The next limit, chosen before its tables are hashed; see `tableNodesBetween`. */
  let nextLimit: number | null = null;
  let hashQueue: readonly OoxmlNode[] = [];
  let hashed = 0;
  const finishPaused = (): void => {
    restorePublished = null;
    const paused = pass;
    pass = null;
    if (paused) drainLayoutSteps(paused);
  };
  const abandonForEdit = (): void => {
    if (!restorePublished || host.limit() === null) return finishPaused();
    // A foreground edit owns the published pages. The background pass can resume from
    // their checkpoints after the edit, without finishing or painting its unpublished pages.
    restorePublished();
    restorePublished = null;
    pass = null;
    nextLimit = null;
    hashQueue = [];
    hashed = 0;
  };
  const step = (budgetMs: number, force = false): boolean => {
    if (!force && host.composing()) return false;
    const deadline =
      host.now() + (pendingReveals.has(surface) ? Math.max(budgetMs, REVEAL_SLICE_MS) : budgetMs);
    while (host.limit() !== null && !host.destroyed()) {
      if (!pass) {
        if (host.busy()) {
          if (!force) return false;
          host.finishBusy();
        }
        const limit = host.limit()!;
        const total = host.bodyBlockCount();
        if (total !== undefined && limit >= total) {
          host.finish();
          // The final flush can itself leave later pages to a background pass; a reader of
          // the whole pagination needs those too.
          if (force && host.busy()) host.finishBusy();
          return true;
        }
        if (nextLimit === null) {
          nextLimit =
            total === undefined
              ? limit + growth
              : nextPrefixLimit(limit, total, (index) => host.bodyBlock(index));
          hashQueue = tableNodesBetween(limit, nextLimit, (index) => host.bodyBlock(index));
          hashed = 0;
        }
        // A table's key hashes its whole content at once; one row at a time here, it fits in
        // a slice, and the pass then reads the digests it needs from the memo.
        // At least one row per slice, so a slice always makes progress.
        for (let rows = 0; hashed < hashQueue.length; rows += 1) {
          if (rows > 0 && (host.now() >= deadline || host.inputPending())) return false;
          warmLayoutNodeDigests([hashQueue[hashed++]!]);
        }
        // An edit during the hashing can have moved the cut past the limit chosen before it.
        if (nextLimit <= host.limit()!) {
          nextLimit = null;
          continue;
        }
        if (shortPrefix && !restorePublished) {
          const restore = host.snapshot();
          const savedGrowth = growth;
          restorePublished = () => {
            restore();
            host.setLimit(limit);
            growth = savedGrowth;
          };
        }
        host.setLimit(nextLimit);
        pass = host.prefixPass(nextLimit);
        nextLimit = null;
      }
      let result = pass.next();
      while (!result.done) {
        if (result.value?.restorable === false) restorePublished = null;
        if (host.now() >= deadline || host.inputPending()) return false;
        result = pass.next();
      }
      pass = null;
      const waiting = pendingReveals.get(surface);
      const reached = waiting?.covers(result.value) === true;
      const work = steps.get(surface);
      if (
        host.limit() !== null &&
        (reached || work?.atEnd || host.now() - publishedAt >= PREFIX_PUBLISH_INTERVAL_MS)
      ) {
        restorePublished = null;
        if (work) work.atEnd = false;
        host.publish(result.value);
        publishedAt = host.now();
      }
      if (reached) {
        pendingReveals.delete(surface);
        waiting!.reveal();
      }
      // Before the first pass reports the block count, steps double; see the constants above.
      growth = Math.min(MAX_PREFIX_STEP, growth * 2);
      if (host.now() >= deadline || host.inputPending()) return false;
    }
    return true;
  };
  container().removeAttribute('data-docx-opening-preview');
  if (host.limit() === null) return { finishPaused, abandonForEdit, stop: () => undefined };
  // The later pages are still to come, as after an edit whose pass finishes in the background.
  // The host element can change when the surface moves, so it is read at each use.
  container().setAttribute('data-docx-opening-preview', '');
  container().setAttribute('data-docx-layout-pending', '');
  const indicator = showOpeningIndicator(surface, container());
  // The reader moved on: a reveal still waiting must not pull the view back later. The pages
  // layer takes the input and moves with the surface.
  const dropReveal = (): void => void pendingReveals.delete(surface);
  const inputs = ['pointerdown', 'keydown', 'wheel'] as const;
  const target = host.pages();
  for (const type of inputs)
    target.addEventListener(type, dropReveal, { capture: true, passive: true });
  let stopped = false;
  const stop = (): void => {
    if (stopped) return;
    stopped = true;
    steps.delete(surface);
    // The surface keeps `finishPaused` for its lifetime: drop what it closes over, or the last
    // table rows hashed and a paused pass keep old revisions of the document alive.
    pass = null;
    restorePublished = null;
    nextLimit = null;
    hashQueue = [];
    for (const type of inputs) target.removeEventListener(type, dropReveal, true);
    container().removeAttribute('data-docx-opening-preview');
    if (!host.busy()) container().removeAttribute('data-docx-layout-pending');
    indicator.remove();
  };
  registerProgressiveOpen(
    surface,
    (budgetMs, force) => {
      const done = step(budgetMs, force);
      if (!done) {
        indicator.place();
        return false;
      }
      stop();
      const waiting = pendingReveals.get(surface);
      pendingReveals.delete(surface);
      if (!host.destroyed()) waiting?.reveal();
      return true;
    },
    shortPrefix
  );
  return { finishPaused, abandonForEdit, stop };
}

/**
 * Finish `surface`'s opening layout now, for a reader of the whole pagination (a page count,
 * TOC page numbers), even while an IME composition is open. True when there was one to finish.
 */
export function completeProgressiveOpen(surface: object): boolean {
  if (!steps.has(surface)) return false;
  while (!steps.get(surface)?.run(Number.POSITIVE_INFINITY, true) && steps.has(surface));
  return true;
}

/** Forget a waiting reveal: a newer one went ahead at once. */
export function dropPendingReveal(surface: object): void {
  pendingReveals.delete(surface);
}

/**
 * Reveal a target the opening has not laid out yet, once it has: a page index, or a position
 * in a body paragraph that `exists`. `reveal` runs after the first prefix that holds the target
 * is shown, or after the full layout when none did, and only while the scroller is where it
 * was when asked. A newer reveal replaces a waiting one.
 * False when `surface` is not opening or the paragraph does not exist, so the caller answers
 * at once.
 */
export function revealWhenLaidOut(
  surface: object,
  target: number | SemanticPosition,
  reveal: () => void,
  exists: (paragraphId: string) => boolean,
  scrollTop: () => number
): boolean {
  if (!steps.has(surface)) return false;
  if (typeof target !== 'number' && !exists(target.paragraphId)) return false;
  const covers = (layout: unknown): boolean =>
    typeof target === 'number'
      ? (layout as SemanticLayout).pages.length > target
      : caretAt(layout as SemanticLayout, target) !== null;
  const asked = scrollTop();
  pendingReveals.set(surface, {
    covers,
    reveal: () => {
      if (Math.abs(scrollTop() - asked) < 1) reveal();
    },
  });
  steps.get(surface)?.wake?.();
  return true;
}

/** What the reveal helpers below read from their surface. */
export interface OpeningRevealHost {
  flushLayout(): void;
  paragraphIds(): readonly string[];
  scroller(): HTMLElement | null;
}

/**
 * A surface's reveal helpers during an opening. `begin` starts every reveal: it lands pending
 * layout and drops a waiting reveal, which a newer one replaces even when it goes at once.
 * `later` makes a reveal of a target not laid out yet wait for it; see
 * {@link revealWhenLaidOut}. It needs a scroller, as every reveal does.
 */
export function openingReveals(surface: () => object, host: OpeningRevealHost) {
  return {
    begin(): void {
      host.flushLayout();
      dropPendingReveal(surface());
    },
    later(target: number | SemanticPosition, reveal: () => void): boolean {
      return (
        host.scroller() !== null &&
        revealWhenLaidOut(
          surface(),
          target,
          reveal,
          (id) => host.paragraphIds().includes(id),
          () => host.scroller()?.scrollTop ?? 0
        )
      );
    },
  };
}
