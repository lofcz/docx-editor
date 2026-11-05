import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  contentControlPropertiesOf,
  findNode,
  paragraphTextOf,
  parentNodeOf,
} from '@docx-editor.dev/core/store';
import { checklistTag } from '../registry/adapter-names';
import type { BlockCommandContext } from './block-type';
import { setManualListChecked } from './lists';

export interface TaskCheckboxOptions {
  readonly onError?: (message: string) => void;
  readonly label?: () => string;
  readonly disabledReason?: (reason: string) => string | undefined;
}
/** Owned task widgets use the same canonical command as keyboard and agent list edits. */
export function installTaskCheckboxEditing(
  editor: DocxEditorInstance,
  context: BlockCommandContext,
  options: TaskCheckboxOptions
) {
  const surface = editor.surface!,
    host = surface.hostElement();
  const inspect = (widget: Element | null) => {
    if (!widget || editor.surface !== surface) return null;
    const id = widget.getAttribute('data-docx-cc-id');
    if (!id) return null;
    const part = surface.session.part(),
      control = findNode(part, id);
    if (!control || control.kind === 'textValue') return null;
    const properties = contentControlPropertiesOf(control);
    if (properties.tag !== checklistTag() || !properties.checkbox) return null;
    let parent = parentNodeOf(part, id);
    for (let depth = 0; parent && parent.kind !== 'paragraph' && depth < 64; depth++)
      parent = parentNodeOf(part, parent.id);
    return parent?.kind === 'paragraph'
      ? { paragraphId: parent.id, checked: properties.checkbox.checked }
      : null;
  };
  const activate = (event: Event) => {
    const widget = (event.target as Element | null)?.closest?.('[data-docx-cc-widget="checkbox"]');
    const task = inspect(widget ?? null);
    if (!task) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    if (widget!.hasAttribute('disabled') || widget!.getAttribute('data-disabled-reason')) return;
    const result = setManualListChecked(editor, !task.checked, context, task.paragraphId);
    if (!result.success) options.onError?.(result.error);
    else if (event.type === 'keydown' || event.type === 'click')
      queueMicrotask(() => {
        for (const current of host.querySelectorAll<HTMLElement>(
          '[data-docx-cc-widget="checkbox"]'
        ))
          if (inspect(current)?.paragraphId === task.paragraphId) {
            current.focus({ preventScroll: true });
            break;
          }
      });
  };
  const pointer = (event: PointerEvent) => {
    if (event.button === 0) activate(event);
  };
  const key = (event: KeyboardEvent) => {
    if (
      !event.isComposing &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.altKey &&
      [' ', 'Enter'].includes(event.key)
    )
      activate(event);
  };
  const click = (event: MouseEvent) => {
    if (event.detail === 0) activate(event);
  };
  const decorate = (root: Element) => {
    const widgets = [...root.querySelectorAll('[data-docx-cc-widget="checkbox"]')];
    if (root.matches('[data-docx-cc-widget="checkbox"]')) widgets.push(root);
    for (const widget of widgets) {
      const task = inspect(widget);
      if (!task) continue;
      widget.setAttribute('tabindex', '0');
      const text = paragraphTextOf(surface.session.part(), task.paragraphId)?.slice(1).trim();
      const label = text || options.label?.();
      if (label) widget.setAttribute('aria-label', label);
      const reason = widget.getAttribute('data-disabled-reason');
      const title = reason ? options.disabledReason?.(reason) : undefined;
      if (title) widget.setAttribute('title', title);
      else widget.removeAttribute('title');
    }
  };
  // Existing widgets can survive a text-only paint. Their accessible names still change.
  const stopPaint = surface.observePaint({ afterPaint: (frame) => decorate(frame.pagesLayer) });
  decorate(host);
  host.addEventListener('pointerdown', pointer, true);
  host.addEventListener('keydown', key, true);
  host.addEventListener('click', click, true);
  return () => {
    stopPaint();
    host.removeEventListener('pointerdown', pointer, true);
    host.removeEventListener('keydown', key, true);
    host.removeEventListener('click', click, true);
  };
}
