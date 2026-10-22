import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { mountBlockClipboard, type BlockClipboardOptions } from './clipboard';
import { inlineLayoutBoxes } from './inline-layout-boxes';
import { isWidgetInputEvent } from './input-boundary';

export interface InlineInput {
  dispose(): void;
}

/**
 * Bind native atom input and payload clipboard handling. Native projected caret stops handle
 * arrows; the engine handles whole-control deletion. Expand stale partial selections before
 * copying, cutting, pasting, or replacing an admitted inline widget.
 */
export function mountInlineInput(
  editor: DocxEditorInstance,
  options: BlockClipboardOptions = {}
): InlineInput {
  const surface = editor.surface;
  if (!surface) throw new Error('Attach the editor before mounting inline input.');
  const host = surface.hostElement();
  const target = host.ownerDocument;
  let indexedPackage: OoxmlPackage | null = null;
  let valid = new Set<string>();
  function normalize(event: Event): void {
    const node = event.target as Node | null;
    if (!node || !host.contains(node)) return;
    if (isWidgetInputEvent(event)) return;
    if (
      event.type === 'keydown' &&
      !['Backspace', 'Delete', 'Enter'].includes((event as KeyboardEvent).key)
    )
      return;
    if (event.type === 'beforeinput') {
      const input = event as InputEvent;
      if (
        input.isComposing ||
        !/^(?:insert|delete)/.test(input.inputType) ||
        input.inputType.includes('Composition')
      )
        return;
    }
    surface!.flushPendingInput();
    const selection = surface!.state().selection;
    if (!selection) return;
    const layout = surface!.layout(),
      boxes = inlineLayoutBoxes(layout);
    if (!boxes.size) return;
    const pkg = surface!.session.currentPackage();
    if (pkg !== indexedPackage) {
      indexedPackage = pkg;
      valid = new Set([...buildBlockIndex(pkg).byId.values()].map((block) => block.controlId));
    }
    const order = new Map(surface!.session.paragraphIds().map((id, ordinal) => [id, ordinal]));
    const compare = (a: typeof selection.anchor, b: typeof selection.anchor) => {
      const ai = order.get(a.paragraphId),
        bi = order.get(b.paragraphId);
      return ai === undefined || bi === undefined ? null : ai - bi || a.offset - b.offset;
    };
    const direction = compare(selection.anchor, selection.head);
    if (direction === null) return;
    let from = direction <= 0 ? selection.anchor : selection.head;
    let to = direction <= 0 ? selection.head : selection.anchor;
    // Nested controls can extend a range into another enclosing atom. Repeat to a fixed point.
    let changed = true;
    while (changed) {
      changed = false;
      for (const control of layout.contentControls ?? []) {
        if (
          !valid.has(control.id) ||
          !control.range ||
          !control.fragments.some((fragment) => boxes.has(`${fragment.pageIndex}:${control.id}`))
        )
          continue;
        const { paragraphId, start, end } = control.range;
        const first = { paragraphId, offset: start },
          last = { paragraphId, offset: end };
        const left = compare(from, last),
          right = compare(to, first);
        if (left === null || right === null || left >= 0 || right <= 0) continue;
        if ((compare(from, first) ?? 0) > 0) {
          from = first;
          changed = true;
        }
        if ((compare(to, last) ?? 0) < 0) {
          to = last;
          changed = true;
        }
      }
    }
    const next = direction <= 0 ? { anchor: from, head: to } : { anchor: to, head: from };
    if (next.anchor !== selection.anchor || next.head !== selection.head)
      surface!.setSelection(next);
  }
  // Document capture precedes an existing surface clipboard lease, regardless of mounting order.
  for (const type of ['copy', 'cut', 'paste', 'beforeinput', 'keydown'])
    target.addEventListener(type, normalize, true);
  let clipboard: ReturnType<typeof mountBlockClipboard>;
  try {
    clipboard = mountBlockClipboard(editor, options);
  } catch (error) {
    for (const type of ['copy', 'cut', 'paste', 'beforeinput', 'keydown'])
      target.removeEventListener(type, normalize, true);
    throw error;
  }
  let disposed = false;
  return {
    dispose() {
      if (disposed) return;
      disposed = true;
      for (const type of ['copy', 'cut', 'paste', 'beforeinput', 'keydown'])
        target.removeEventListener(type, normalize, true);
      clipboard.dispose();
      indexedPackage = null;
      valid.clear();
    },
  };
}
