import type { BlockFragmentRecord, PageRecord } from '../layout/semantic-records.ts';
import type { RetainedPage } from './semantic-paint-paragraph-roots.ts';
import {
  canRetainPaintedParagraph,
  retainPaintedParagraph,
} from './semantic-paint-paragraph-children.ts';
import { sameDataWithin } from './semantic-paint-record-equality.ts';

/** Insert and remove fragments at their canonical page positions, including host-owned markers. */
export function planStructuralRoots(
  prior: RetainedPage,
  page: PageRecord,
  paint: (fragment: BlockFragmentRecord) => HTMLElement,
  homeOf: (element: HTMLElement) => ChildNode
) {
  const { content, blocks } = prior;
  if (!content || !blocks || content.parentElement !== prior.element) return null;
  const existing = new Map<
    string,
    { record: BlockFragmentRecord; root: HTMLElement; home: ChildNode }
  >();
  let previous: ChildNode | undefined;
  for (const [record, root] of blocks) {
    const home = homeOf(root);
    if (
      !content.contains(root) ||
      (home !== root && home.nodeType !== 8) ||
      home.parentNode !== content ||
      !root.classList.contains(
        record.kind === 'paragraph' ? 'docx-paragraph-fragment' : 'docx-table-fragment'
      ) ||
      root.dataset.fragmentIndex !== String(record.fragmentIndex) ||
      (record.kind === 'paragraph'
        ? root.dataset.paragraphId !== record.paragraphId
        : root.dataset.tableId !== record.tableId) ||
      (previous && previous.compareDocumentPosition(home) !== home.DOCUMENT_POSITION_FOLLOWING) ||
      existing.has(record.id)
    )
      return null;
    existing.set(record.id, { record, root, home });
    previous = home;
  }
  const planned: { fragment: BlockFragmentRecord; root?: HTMLElement }[] = [];
  const updates = new Map<HTMLElement, HTMLElement>();
  const survivors = new Map<HTMLElement, ChildNode>();
  previous = undefined;
  for (const fragment of page.fragments) {
    const old = existing.get(fragment.id);
    const same = old?.record.kind === fragment.kind && sameDataWithin(old.record, fragment, 20000);
    // Paint added blocks and replacement tables only after the host accepts this plan.
    const fresh =
      !same && old?.record.kind === 'paragraph' && fragment.kind === 'paragraph'
        ? paint(fragment)
        : undefined;
    if (old && (same || (fresh && canRetainPaintedParagraph(old.root, fresh)))) {
      // A reorder can change a host's ownership. Use the regular reconciliation for that page.
      if (
        previous &&
        previous.compareDocumentPosition(old.home) !== old.home.DOCUMENT_POSITION_FOLLOWING
      )
        return null;
      previous = old.home;
      survivors.set(old.root, old.home);
      planned.push({ fragment, root: old.root });
      if (fresh) updates.set(old.root, fresh);
    } else planned.push({ fragment, root: fresh });
  }
  if (!survivors.size) return null;
  const changedRoots = new Set(updates.keys());
  for (const { root } of existing.values()) if (!survivors.has(root)) changedRoots.add(root);
  return {
    changedRoots,
    apply() {
      for (const { root, home } of existing.values()) {
        if (survivors.has(root)) continue;
        if (home !== root) home.remove();
        root.remove();
      }
      for (const [root, fresh] of updates) retainPaintedParagraph(root, fresh);
      let before: ChildNode | null = null;
      const nextBlocks = new Map(
        planned.map(({ fragment, root }) => [fragment, root ?? paint(fragment)])
      );
      const ordered = [...nextBlocks.values()];
      for (let i = ordered.length - 1; i >= 0; i--) {
        const root = ordered[i]!;
        const home = survivors.get(root);
        // Another observer can return its own loan after this plan was accepted.
        // The restored root then occupies the original marker's position.
        if (home) before = home.parentNode === content ? home : root;
        else {
          content.insertBefore(root, before);
          before = root;
        }
      }
      return nextBlocks;
    },
  };
}
