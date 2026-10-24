import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { EditorCommand } from '@docx-editor.dev/core/contracts/editor';

export type ManualTableRowAction = 'above' | 'below' | 'delete';
const command = (action: ManualTableRowAction): EditorCommand =>
  action === 'delete' ? { type: 'deleteRow' } : { type: 'insertRow', where: action };

/** Native row commands resolve the current cell or rectangle and preserve neighboring properties. */
export function changeManualTableRow(editor: DocxEditorInstance, action: ManualTableRowAction) {
  editor.surface?.flushPendingInput();
  return editor.exec(command(action));
}

export function createTableRowToolbar(editor: DocxEditorInstance) {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const element = surface.hostElement().ownerDocument.createElement('div');
  element.setAttribute('role', 'toolbar');
  element.setAttribute('aria-label', 'Table rows');
  const buttons = (['above', 'below', 'delete'] as const).map((action) => {
    const button = element.ownerDocument.createElement('button');
    button.type = 'button';
    button.textContent = action === 'delete' ? 'Delete row' : `Insert row ${action}`;
    button.addEventListener('mousedown', (event) => {
      if (event.button === 0) event.preventDefault();
    });
    button.addEventListener('click', () => {
      if (changeManualTableRow(editor, action).ok) surface.focus();
    });
    element.append(button);
    return { button, action };
  });
  const update = () => {
    for (const { button, action } of buttons) {
      const allowed = editor.can(command(action));
      button.disabled = !allowed.ok;
      button.title = allowed.ok ? '' : allowed.reason;
    }
  };
  const stops = [editor.on('change', update), editor.on('selectionChange', update)];
  update();
  return { element, dispose: () => stops.forEach((stop) => stop()) };
}
