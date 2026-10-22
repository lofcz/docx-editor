import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { findNode } from '@docx-editor.dev/core/store';
import type { DirectChangeLedger } from './change-ledger';
import { watchPlaybackLifecycle } from './playback-lifecycle';

/** Playback never exceeds this budget, including scrolls, across a burst of edits. */
export const PLAYBACK_CAP_MS = 4000;
const DEFAULT_DURATION = 2400;
const STYLE_ID = 'sciobot-ai-playback-style';
const CSS = `
[data-sciobot-ai-sweep]{position:fixed;pointer-events:none;z-index:40;border-radius:2px;
background:linear-gradient(90deg,rgba(124,58,237,.18),rgba(124,58,237,.08));}
[data-sciobot-ai-caret]{position:fixed;pointer-events:none;z-index:41;width:2px;background:#7c3aed;
transition:left .12s linear,top .12s linear,height .12s linear;}
[data-sciobot-ai-caret]>span{position:absolute;left:4px;top:-1.35em;white-space:nowrap;
font:600 11px/1.4 system-ui,sans-serif;color:#fff;background:#7c3aed;border-radius:3px;padding:0 4px;}`;

export interface DirectPlaybackOptions {
  /** Collaborator label beside the AI caret. Default: AI. */
  readonly label?: string;
  /** Total budget in milliseconds; clamped to 0..4000. Default: 2400. */
  readonly duration?: number;
  /** Disable motion explicitly. Otherwise follows prefers-reduced-motion. */
  readonly reducedMotion?: boolean;
  /** Aborts playback only; the committed edit stays. */
  readonly signal?: AbortSignal;
}

export interface DirectPlayback {
  /** Whether any motion plays. Without motion every change is revealed immediately. */
  readonly motion: boolean;
  /** Resolves when playback ends or is cancelled. */
  readonly finished: Promise<void>;
  /** Stop playback, remove overlays, and reveal remaining changes. Never writes the document. */
  cancel(): void;
}

const active = new WeakMap<DocxEditorInstance, { cancel(): void; startedAt: number }>();

const clamp = (value: number, min: number, max: number) => Math.max(min, Math.min(max, value));

const CSS_ESCAPE = (globalThis as { CSS?: { escape?: (value: string) => string } }).CSS?.escape;

/** Painted line rectangles of one native paragraph, from the surface's text fragments. */
function paragraphRects(root: HTMLElement, paragraphId: string): DOMRect[] {
  const escaped = CSS_ESCAPE ? CSS_ESCAPE(paragraphId) : paragraphId.replace(/["\\]/g, '\\$&');
  return [...root.querySelectorAll<HTMLElement>(`[data-paragraph-id="${escaped}"][data-start]`)]
    .map((element) => element.getBoundingClientRect())
    .filter((rect) => rect.width > 0 && rect.height > 0);
}

function union(rects: readonly DOMRect[]) {
  const top = Math.min(...rects.map((rect) => rect.top)),
    bottom = Math.max(...rects.map((rect) => rect.bottom)),
    left = Math.min(...rects.map((rect) => rect.left)),
    right = Math.max(...rects.map((rect) => rect.right));
  return { top, bottom, left, right };
}

/**
 * Play the pending changes of one committed direct edit. The edit is already canonical and
 * fully painted, so every target stays readable throughout: overlays only tint the changed
 * paragraphs in document order while a labelled caret sweeps across them, and each change is
 * revealed in the ledger as its step ends. Changes deleted during playback (by the user, undo,
 * or a peer) lose their overlay at the next frame and leave no highlight. A newer playback on
 * the same editor cancels this one and inherits its deadline, so a burst shares one budget.
 */
export function playDirectChanges(
  editor: DocxEditorInstance,
  ledger: DirectChangeLedger,
  options: DirectPlaybackOptions = {}
): DirectPlayback {
  const previous = active.get(editor),
    startedAt = previous?.startedAt ?? Date.now();
  previous?.cancel();
  const requested = options.duration ?? DEFAULT_DURATION,
    budget = clamp(Number.isFinite(requested) ? requested : DEFAULT_DURATION, 0, PLAYBACK_CAP_MS),
    deadline = startedAt + budget,
    surface = editor.surface,
    root = surface && !surface.parked() ? surface.hostElement() : null,
    doc = root?.ownerDocument,
    win = doc?.defaultView ?? null;
  const pendingIds = () => ledger.changes.filter((change) => change.pending);
  const motion = !!(
    root &&
    win &&
    budget > 0 &&
    Date.now() < deadline &&
    !options.signal?.aborted &&
    pendingIds().length &&
    !(options.reducedMotion ?? win.matchMedia?.('(prefers-reduced-motion: reduce)').matches)
  );
  if (!motion) {
    ledger.reveal();
    return { motion: false, finished: Promise.resolve(), cancel: () => {} };
  }
  let cancelled = false,
    frame = 0,
    timer: ReturnType<typeof setTimeout> | undefined,
    wake: (() => void) | undefined;
  const overlays: HTMLElement[] = [];
  let releaseLifecycle = () => {};
  if (!doc!.getElementById(STYLE_ID)) {
    const style = doc!.createElement('style');
    style.id = STYLE_ID;
    style.textContent = CSS;
    doc!.head.append(style);
  }
  const caret = doc!.createElement('div');
  caret.dataset.sciobotAiCaret = '';
  caret.setAttribute('aria-hidden', 'true');
  const label = doc!.createElement('span');
  label.textContent = `✦ ${options.label ?? 'AI'}`;
  caret.append(label);
  caret.style.display = 'none';
  doc!.body.append(caret);
  const clear = () => {
    for (const overlay of overlays.splice(0)) overlay.remove();
  };
  const stop = () => {
    if (cancelled) return;
    cancelled = true;
    releaseLifecycle();
    if (frame) win!.cancelAnimationFrame(frame);
    if (timer) clearTimeout(timer);
    clear();
    caret.remove();
    options.signal?.removeEventListener('abort', stop);
    if (active.get(editor)?.cancel === stop) active.delete(editor);
    if (editor.surface !== surface) ledger.settleAll();
    else ledger.reveal();
    wake?.();
  };
  options.signal?.addEventListener('abort', stop, { once: true });
  active.set(editor, { cancel: stop, startedAt });
  const lifecycle = watchPlaybackLifecycle(editor, stop, options.reducedMotion === undefined);
  releaseLifecycle = () => lifecycle.dispose();

  const wait = (ms: number) =>
    new Promise<void>((resolve) => {
      wake = resolve;
      timer = setTimeout(resolve, Math.max(0, ms));
    });
  const order = (ids: readonly string[]) => {
    const all = surface!.session.paragraphIds();
    return [...ids].sort((a, b) => all.indexOf(a) - all.indexOf(b));
  };

  const finished = (async () => {
    // Each step owns one pending change; the ledger drops changes that later commits touched.
    const seen = new Set<string>();
    for (;;) {
      if (cancelled) return;
      const remaining = pendingIds().filter((change) => !seen.has(change.id));
      const now = Date.now();
      if (!remaining.length || now >= deadline) break;
      const part = surface!.session.part(),
        positions = surface!.session.paragraphIds();
      remaining.sort(
        (a, b) =>
          Math.min(...a.nodeIds.map((id) => positions.indexOf(id))) -
          Math.min(...b.nodeIds.map((id) => positions.indexOf(id)))
      );
      const change = remaining[0]!;
      seen.add(change.id);
      const live = order(change.nodeIds.filter((id) => findNode(part, id)?.kind === 'paragraph'));
      if (!live.length) {
        ledger.reveal(change.nodeIds);
        continue;
      }
      const slice = (deadline - now) / remaining.length;
      // Hosts without the native scroll-container class fall back to the painted element.
      if (!surface!.revealParagraph(live[0]!, { block: 'centerIfNeeded', behavior: 'auto' }))
        root!
          .querySelector<HTMLElement>(
            `[data-paragraph-id="${CSS_ESCAPE ? CSS_ESCAPE(live[0]!) : live[0]!}"][data-start]`
          )
          ?.scrollIntoView({ block: 'nearest' });
      const sweep = doc!.createElement('div');
      sweep.dataset.sciobotAiSweep = change.id;
      sweep.setAttribute('aria-hidden', 'true');
      doc!.body.append(sweep);
      overlays.push(sweep);
      caret.style.display = '';
      const begin = performance.now();
      const paint = () => {
        frame = 0;
        if (cancelled) return;
        if (!lifecycle.current()) {
          stop();
          return;
        }
        const current = surface!.session.part(),
          rects = live
            .filter((id) => findNode(current, id)?.kind === 'paragraph')
            .flatMap((id) => paragraphRects(root!, id));
        if (!rects.length) {
          // Deleted or moved off-screen: no stale overlay.
          sweep.style.display = 'none';
          caret.style.display = 'none';
        } else {
          const box = union(rects),
            progress = clamp((performance.now() - begin) / Math.max(1, slice), 0, 1),
            scaled = progress * rects.length,
            index = Math.min(rects.length - 1, Math.floor(scaled)),
            line = rects[index]!,
            x = line.left + (line.right - line.left) * (progress >= 1 ? 1 : scaled - index);
          Object.assign(sweep.style, {
            display: '',
            top: `${box.top}px`,
            left: `${box.left}px`,
            height: `${box.bottom - box.top}px`,
            width: `${(box.right - box.left) * progress}px`,
          });
          Object.assign(caret.style, {
            display: '',
            top: `${line.top}px`,
            left: `${x}px`,
            height: `${line.bottom - line.top}px`,
          });
        }
        frame = win!.requestAnimationFrame(paint);
      };
      paint();
      await wait(slice);
      if (frame) win!.cancelAnimationFrame(frame);
      frame = 0;
      sweep.remove();
      overlays.splice(overlays.indexOf(sweep), 1);
      ledger.reveal(change.nodeIds);
    }
    stop();
  })();
  return { motion: true, finished, cancel: stop };
}
