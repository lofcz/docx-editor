import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { setBlockType, type BlockCommandContext, type TextBlockType } from './block-type';
import { insertBlock } from './insert-block';

/** Host-owned controls preserve the native model selection when the pointer enters the toolbar. */
export function createBlockToolbar(editor: DocxEditorInstance, context: BlockCommandContext) {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const toolbar = surface.hostElement().ownerDocument.createElement('div');
  toolbar.setAttribute('role', 'toolbar');
  toolbar.setAttribute('aria-label', 'Paragraph type');
  const labels: Readonly<Record<TextBlockType | 'hr', string>> = {
    p: 'Paragraph',
    h1: 'Heading 1',
    h2: 'Heading 2',
    h3: 'Heading 3',
    blockquote: 'Blockquote',
    hr: 'Horizontal rule',
  };
  for (const type of Object.keys(labels) as (TextBlockType | 'hr')[]) {
    const button = toolbar.ownerDocument.createElement('button');
    button.type = 'button';
    button.textContent = labels[type];
    button.addEventListener('mousedown', (event) => {
      if (event.button === 0) event.preventDefault();
    });
    button.addEventListener('click', () => {
      const result =
        type === 'hr' ? insertBlock(editor, type) : setBlockType(editor, type, context);
      if (result.success) surface.focus();
    });
    toolbar.append(button);
  }
  return toolbar;
}
