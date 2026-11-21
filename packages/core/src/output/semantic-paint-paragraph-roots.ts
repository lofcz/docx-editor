import type {
  BlockFragmentRecord,
  PageRecord,
  ParagraphFragmentRecord,
} from '../layout/semantic-records.ts';
import { changedPageParts } from './semantic-paint-page-diff.ts';
import { sameDataWithin } from './semantic-paint-record-equality.ts';

/**
 * Reserve a page only if the host can keep its loans across these child replacements.
 * @public
 */
export type PreserveParagraphRoots = (
  pageIndex: number,
  accepts: (changedRoots: ReadonlySet<HTMLElement>) => boolean
) => boolean;

/**
 * One painted page, retained so an unchanged page never has to be rebuilt.
 *
 * Incremental layout keeps the RECORD of an untouched page identical across revisions —
 * same object, by design — so record identity is a complete reuse test: same record, same
 * materialization, same paint parameters means the element in hand is already exactly what
 * this pass would build. Rebuilding every sheet per commit made the browser restyle the
 * whole document on each keystroke, which at several hundred pages cost more than layout
 * and paint together.
 */
export interface RetainedPage {
  readonly record: PageRecord;
  readonly materialized: boolean;
  readonly element: HTMLElement;
  /**
   * The body content layer and its block elements by RECORD, kept so the next pass can
   * repaint the one block that moved instead of the sheet it sits on. Absent on a page
   * that was never materialized.
   */
  readonly content?: HTMLElement;
  readonly blocks?: ReadonlyMap<BlockFragmentRecord, HTMLElement>;
}

/** Keep the block roots only when each painted paragraph can keep its existing outer element. */
function planParagraphRootPaint(
  prior: RetainedPage,
  page: PageRecord,
  paint: (fragment: ParagraphFragmentRecord) => HTMLElement
) {
  const { content, blocks } = prior;
  if (
    !content ||
    !blocks ||
    content.parentElement !== prior.element ||
    prior.record.fragments.length !== page.fragments.length
  )
    return null;
  const nextBlocks = new Map<BlockFragmentRecord, HTMLElement>();
  const updates = new Map<HTMLElement, DocumentFragment>();
  let previous: HTMLElement | undefined;
  for (let i = 0; i < page.fragments.length; i++) {
    const from = prior.record.fragments[i]!;
    const to = page.fragments[i]!;
    if (from.id !== to.id || from.kind !== to.kind) return null;
    const root = blocks.get(from);
    if (
      !root ||
      !content.contains(root) ||
      !root.classList.contains(
        to.kind === 'paragraph' ? 'docx-paragraph-fragment' : 'docx-table-fragment'
      ) ||
      (previous && previous.compareDocumentPosition(root) !== root.DOCUMENT_POSITION_FOLLOWING)
    )
      return null;
    previous = root;
    const identity = to.kind === 'paragraph' ? root.dataset.paragraphId : root.dataset.tableId;
    if (
      identity !== (to.kind === 'paragraph' ? to.paragraphId : to.tableId) ||
      root.dataset.fragmentIndex !== String(to.fragmentIndex)
    )
      return null;
    if (from !== to && !sameDataWithin(from, to, 20000)) {
      if (to.kind !== 'paragraph') return null;
      const fresh = paint(to);
      if (
        root.attributes.length !== fresh.attributes.length ||
        Array.from(fresh.attributes).some((attr) => root.getAttribute(attr.name) !== attr.value)
      )
        return null;
      const children = fresh.ownerDocument.createDocumentFragment();
      while (fresh.firstChild) children.append(fresh.firstChild);
      updates.set(root, children);
    }
    nextBlocks.set(to, root);
  }
  return {
    changedRoots: new Set(updates.keys()),
    apply() {
      for (const [root, children] of updates) root.replaceChildren(children);
      return nextBlocks;
    },
  };
}

/** Offer whole pages to every observer. Only a host that requests roots pays for a partial-page plan. */
export function prepareParagraphRootPaint(
  container: HTMLElement,
  pages: readonly PageRecord[],
  materialize: ReadonlySet<number> | undefined,
  reusable: ReadonlyMap<PageRecord, RetainedPage> | null,
  previousByIndex: ReadonlyMap<number, RetainedPage> | null,
  beforePaint: (retained: ReadonlySet<number>, preserve: PreserveParagraphRoots) => void,
  paint: (page: PageRecord, fragment: ParagraphFragmentRecord) => HTMLElement
) {
  type Plan = NonNullable<ReturnType<typeof planParagraphRootPaint>>;
  const reserved = new Map<number, Plan>();
  const retained = new Set<number>();
  for (const page of pages) {
    if (reusable?.get(page)?.materialized && (materialize?.has(page.index) ?? true))
      retained.add(page.index);
  }
  const plans = new Map<number, Plan | null>();
  let accepting = true;
  try {
    beforePaint(retained, (index, accepts) => {
      if (!accepting || retained.has(index)) return false;
      let plan = plans.get(index);
      if (plan === undefined) {
        const prior = previousByIndex?.get(index),
          page = pages[index];
        if (
          page?.index !== index ||
          !prior?.materialized ||
          prior.element.parentElement !== container ||
          !(materialize?.has(index) ?? true)
        )
          return false;
        // Chrome may change while the sheet keeps every body paragraph and table.
        // Only hosts that reserve paragraph roots can keep their loans in this case.
        if (prior.record === page) return accepts(new Set());
        const delta = changedPageParts(prior.record, page);
        if (!delta?.blocks || delta.furniture || delta.controls) return false;
        plan = planParagraphRootPaint(prior, page, (fragment) => paint(page, fragment));
        plans.set(index, plan);
      }
      if (!plan || !accepts(plan.changedRoots)) return false;
      reserved.set(index, plan);
      return true;
    });
  } finally {
    accepting = false;
  }
  return reserved;
}
