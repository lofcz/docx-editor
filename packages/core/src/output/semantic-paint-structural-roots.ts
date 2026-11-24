import { planParagraphReflow, moveRelativeBox, type ParagraphReflowPaint } from './semantic-paint-paragraph-reflow.ts';
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
  paint: (fragment: BlockFragmentRecord, previousTable?: HTMLElement) => HTMLElement,
  homeOf: (element: HTMLElement) => ChildNode,
  paragraphPaint: ParagraphReflowPaint
) {
  const { content, blocks } = prior;
  if (!content || !blocks || content.parentElement !== prior.element) return null;
  // A continuation can become the first fragment on this sheet. Its table still owns the rows.
  const keyOf = (record: BlockFragmentRecord) =>
    record.kind === 'table' ? `table:${record.tableId}` : `block:${record.id}`;
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
      existing.has(keyOf(record))
    )
      return null;
    existing.set(keyOf(record), { record, root, home });
    previous = home;
  }
  const planned: {
    fragment: BlockFragmentRecord;
    root?: HTMLElement;
    table?: { root: HTMLElement; home: ChildNode };
  }[] = [];
  const seen = new Set<string>();
  const updates = new Map<HTMLElement, () => void>();
  const survivors = new Map<HTMLElement, ChildNode>();
  previous = undefined;
  for (const fragment of page.fragments) {
    const key = keyOf(fragment);
    if (seen.has(key)) return null;
    seen.add(key);
    const old = existing.get(key);
    const same = old?.record.kind === fragment.kind && sameDataWithin(old.record, fragment, 20000);
    // Paint added blocks and replacement tables only after the host accepts this plan.
    // Preserve a translated w:p (CT_P, ECMA-376 Part 1 §17.3.1.22) after host acceptance.
    const translate =
      !same && old?.record.kind === 'paragraph' && fragment.kind === 'paragraph'
        ? planParagraphReflow(old.root, old.record, fragment, paragraphPaint)
        : undefined;
    const fresh =
      !translate && !same && old?.record.kind === 'paragraph' && fragment.kind === 'paragraph'
        ? paint(fragment)
        : undefined;
    const table = old?.record.kind === 'table' && fragment.kind === 'table' && !same;
    if (old && (same || translate || table || (fresh && canRetainPaintedParagraph(old.root, fresh)))) {
      // A reorder can change a host's ownership. Use the regular reconciliation for that page.
      if (
        previous &&
        previous.compareDocumentPosition(old.home) !== old.home.DOCUMENT_POSITION_FOLLOWING
      )
        return null;
      previous = old.home;
      if (table) {
        // Borrowers must release descendants that the ordinary table painter can replace.
        planned.push({ fragment, table: old });
      } else {
        survivors.set(old.root, old.home);
        planned.push({ fragment, root: old.root });
        if (translate) {
          updates.set(old.root, () => {
            translate();
            const origin = { x: 0, y: 0 };
            moveRelativeBox(old.root.style, old.record.box, origin, fragment.box, origin, paragraphPaint.scale);
          });
        } else if (fresh) updates.set(old.root, () => retainPaintedParagraph(old.root, fresh));
      }
    } else planned.push({ fragment, root: fresh });
  }
  if (!survivors.size && !planned.some((entry) => entry.table)) return null;
  const changedRoots = new Set(updates.keys());
  for (const { root } of existing.values()) if (!survivors.has(root)) changedRoots.add(root);
  return {
    changedRoots,
    apply() {
      // Reflow accepted table rows while the old table is still connected.
      const nextBlocks = new Map(
        planned.map(({ fragment, root, table }) => {
          const element = root ?? paint(fragment, table?.root);
          if (table?.root === element) survivors.set(element, table.home);
          return [fragment, element] as const;
        })
      );
      for (const { root, home } of existing.values()) {
        if (survivors.has(root)) continue;
        if (home !== root) home.remove();
        root.remove();
      }
      for (const apply of updates.values()) apply();
      let before: ChildNode | null = null;
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
