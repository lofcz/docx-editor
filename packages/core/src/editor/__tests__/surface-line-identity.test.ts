import { expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor.ts';
import { docx, p } from '../../automation/__tests__/support/protocol.ts';

test('an edit before a line break preserves the following text row at its new offset', () => {
  const container = document.createElement('div');
  document.body.append(container);
  const body = '<w:p><w:r><w:t>Alpha</w:t><w:br/><w:t>Beta</w:t></w:r></w:p>';
  const editor = createDocxEditor({ container, document: docx(body) });
  try {
    const surface = editor.surface!;
    const rows = () => container.querySelectorAll<HTMLElement>('.docx-line');
    const second = rows()[1]!;
    const before = second.dataset.lineId;
    const text = second.querySelector('.layout-run-text')!.firstChild!;
    const point = { paragraphId: surface.session.paragraphIds()[0]!, offset: 1 };
    surface.setSelection({ anchor: point, head: point });
    surface.insertPlainText('bc');
    surface.flushPendingInput();
    expect(rows()[1]).toBe(second);
    expect(second.dataset.lineId).not.toBe(before);
    expect(text.isConnected).toBe(true);
    expect(text.textContent).toBe('Beta');
  } finally {
    editor.destroy();
    container.remove();
  }
});

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

test('inserting a word before a link keeps the link and its text attached', () => {
  const container = document.createElement('div');
  document.body.append(container);
  const body =
    '<w:p><w:r><w:t xml:space="preserve">Read </w:t></w:r>' +
    '<w:hyperlink w:anchor="target"><w:r><w:t>here</w:t></w:r></w:hyperlink></w:p>';
  const editor = createDocxEditor({ container, document: docx(body) });
  try {
    const surface = editor.surface!;
    const link = container.querySelector('a.docx-hyperlink')!;
    const run = link.querySelector('.layout-run-text')!;
    const text = run.firstChild!;
    const point = { paragraphId: surface.session.paragraphIds()[0]!, offset: 0 };
    surface.setSelection({ anchor: point, head: point });
    surface.insertPlainText('Now ');
    surface.flushPendingInput();
    expect(container.querySelector('a.docx-hyperlink')).toBe(link);
    expect(text.isConnected).toBe(true);
    expect(run.getAttribute('data-start')).toBe('9');
    expect(link.getAttribute('href')).toBe('#target');
  } finally {
    editor.destroy();
    container.remove();
  }
});
