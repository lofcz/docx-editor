import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { EditorCommand } from '@docx-editor.dev/core/contracts/editor';

export type ManualTableColumnAction = 'left' | 'right' | 'delete';
const command = (action: ManualTableColumnAction): EditorCommand =>
  action === 'delete' ? { type: 'deleteColumn' } : { type: 'insertColumn', where: action };

/** Native column commands resolve the current cell or rectangle and preserve the Word grid. */
export function changeManualTableColumn(
  editor: DocxEditorInstance,
  action: ManualTableColumnAction
) {
  editor.surface?.flushPendingInput();
  return editor.exec(command(action));
}

export function createTableColumnToolbar(editor: DocxEditorInstance) {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const element = surface.hostElement().ownerDocument.createElement('div');
  element.setAttribute('role', 'toolbar');
  element.setAttribute('aria-label', 'Table columns');
  const buttons = (['left', 'right', 'delete'] as const).map((action) => {
    const button = element.ownerDocument.createElement('button');
    button.type = 'button';
    button.textContent = action === 'delete' ? 'Delete column' : `Insert column ${action}`;
    button.addEventListener('mousedown', (event) => {
      if (event.button === 0) event.preventDefault();
    });
    button.addEventListener('click', () => {
      if (changeManualTableColumn(editor, action).ok) surface.focus();
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
