import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { ExecResult } from '@docx-editor.dev/core/contracts/editor';

/** Production Plate choices, expressed as Word multiples. */
export const LINE_HEIGHT_CHOICES = [1, 1.2, 1.5, 2, 3] as const;
export type LineHeightChoice = (typeof LINE_HEIGHT_CHOICES)[number];

/** Keep imported exact and minimum spacing visible without converting it to a multiple. */
export function lineHeightState(editor: DocxEditorInstance) {
  return editor.surface?.formatting().lineSpacing ?? null;
}

/** Native spacing merges line settings while preserving before/after spacing and other properties. */
export function setLineHeight(editor: DocxEditorInstance, value: LineHeightChoice): ExecResult {
  if (!LINE_HEIGHT_CHOICES.includes(value))
    return { ok: false, code: 'invalidArgs', reason: 'Choose an application line height.' };
  editor.surface?.flushPendingInput();
  return editor.exec({ type: 'setLineSpacing', rule: 'multiple', value });
}

/** Host controls preserve the native selection and subscribe to canonical formatting state. */
export function createLineHeightToolbar(editor: DocxEditorInstance) {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const element = surface.hostElement().ownerDocument.createElement('div');
  element.setAttribute('role', 'toolbar');
  element.setAttribute('aria-label', 'Line height');
  const current = element.ownerDocument.createElement('output');
  current.setAttribute('aria-label', 'Current line height');
  element.append(current);
  const buttons = LINE_HEIGHT_CHOICES.map((value) => {
    const button = element.ownerDocument.createElement('button');
    button.type = 'button';
    button.textContent = String(value);
    button.addEventListener('mousedown', (event) => {
      if (event.button === 0) event.preventDefault();
    });
    button.addEventListener('click', () => {
      if (setLineHeight(editor, value).ok) surface.focus();
    });
    element.append(button);
    return { button, value };
  });
  const update = () => {
    const spacing = lineHeightState(editor);
    current.textContent = spacing
      ? spacing.rule === 'multiple'
        ? String(spacing.value)
        : `${spacing.rule === 'exact' ? 'Exactly' : 'At least'} ${spacing.value} pt`
      : 'Mixed';
    for (const { button, value } of buttons) {
      const allowed = editor.can({ type: 'setLineSpacing', rule: 'multiple', value });
      button.disabled = !allowed.ok;
      button.title = allowed.ok ? '' : allowed.reason;
      button.setAttribute(
        'aria-pressed',
        String(spacing?.rule === 'multiple' && spacing.value === value)
      );
    }
  };
  const stops = [editor.on('change', update), editor.on('selectionChange', update)];
  update();
  return { element, dispose: () => stops.forEach((stop) => stop()) };
}
