import type { SurfacePaintFrame } from '@docx-editor.dev/core/editor';
import { contentControlContentNodeOf, findNode, type OoxmlNode } from '@docx-editor.dev/core/store';

/** Animate newly opened Word fragments without changing their measured positions. */
export function revealDisclosureBody(frame: SurfacePaintFrame, id: string, summary: string): void {
  const win = frame.pagesLayer.ownerDocument.defaultView;
  if (!win?.matchMedia?.('(prefers-reduced-motion: no-preference)').matches) return;
  const control = findNode(frame.part, id);
  const content = control && contentControlContentNodeOf(control);
  if (!content) return;
  const blocks = new Set<string>();
  function collect(node: OoxmlNode): void {
    if (node.kind === 'paragraph' || node.kind === 'table') {
      if (node.id !== summary) blocks.add(node.id);
    } else if (node.kind !== 'textValue') {
      for (const child of node.children) collect(child);
    }
  }
  // The entire leading block stays visible when the summary spans several paragraphs.
  for (const child of content.children.slice(1)) collect(child);
  for (const element of frame.pagesLayer.querySelectorAll<HTMLElement>(
    '.docx-paragraph-fragment, .docx-table-fragment'
  )) {
    const block = element.dataset.tableId ?? element.dataset.paragraphId;
    if (block && blocks.has(block))
      element.animate?.(
        [
          { opacity: 0.35, translate: '0 -3px' },
          { opacity: 1, translate: '0 0' },
        ],
        { duration: 140, easing: 'ease-out' }
      );
  }
}
