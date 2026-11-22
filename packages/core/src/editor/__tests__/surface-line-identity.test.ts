import { expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor.ts';
import { docx, p } from '../../automation/__tests__/support/protocol.ts';

test('typing keeps the connected line, text node, and paragraph control', () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: docx(p('Alpha') + p('Beta')) });
  try {
    const surface = editor.surface!;
    const root = container.querySelector<HTMLElement>('.docx-paragraph-fragment')!;
    const position = { paragraphId: root.dataset.paragraphId!, offset: 3 };
    surface.setSelection({ anchor: position, head: position });
    const line = root.querySelector('.docx-line')!;
    const run = line.querySelector('.layout-run-text')!;
    const text = run.firstChild!;
    const button = root.appendChild(document.createElement('button'));
    button.setAttribute('contenteditable', 'false');
    surface.insertPlainText('Z');
    surface.flushPendingInput();
    expect(text.textContent).toBe('AlpZha');
    expect(line.isConnected).toBe(true);
    expect(run.firstChild).toBe(text);
    expect(button.isConnected).toBe(true);
    expect(container.querySelector('.docx-line')).toBe(line);
  } finally {
    editor.destroy();
    container.remove();
  }
});
