import { expect, test } from 'bun:test';
import { act } from 'react';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, paragraphTextOf } from '@docx-editor.dev/core/store';
import { nativeSlotDocument } from '../../test/support/react-block-document';
import { mountReactBlockHost, type ReactBlockViewProps } from './host';

const environment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean };
function Container({ block, children }: ReactBlockViewProps) {
  return (
    <div data-container={block.id} style={{ display: 'contents' }}>
      {children}
    </div>
  );
}
const host = (editor: DocxEditorInstance) =>
  mountReactBlockHost(
    editor,
    new Map([['native_container', { view: Container, nativeContent: true }]])
  );
function paragraph(editor: DocxEditorInstance, text: string) {
  const session = editor.surface!.session;
  const id = session.paragraphIds().find((id) => paragraphTextOf(session.part(), id) === text);
  if (!id) throw new Error(`Missing paragraph: ${text}`);
  return id;
}
function structure(container: HTMLElement) {
  const own = (id: string) => container.querySelector(`[data-container="${id}"]`)!;
  expect(container.querySelectorAll('[data-sciobot-native-slot]')).toHaveLength(7);
  expect(own('outer').contains(own('inner'))).toBe(true);
  expect(own('outer').contains(own('only-child'))).toBe(true);
  expect(own('only-child').contains(own('leaf'))).toBe(true);
  expect(own('outer').querySelectorAll('.docx-table-fragment')).toHaveLength(1);
  expect(own('inside-cell').closest('.docx-table-cell') !== null).toBe(true);
  expect(own('inside-cell').textContent).toBe('Cell text');
  expect(own('row').querySelectorAll('.docx-table-row')).toHaveLength(1);
  expect(own('cell').querySelectorAll('.docx-table-cell')).toHaveLength(1);
  expect(own('row').contains(own('cell'))).toBe(true);
  expect(
    [...container.querySelectorAll('.docx-paragraph-fragment')].map((p) => p.textContent)
  ).toEqual([
    'Intro',
    'Outer start',
    'Inner text',
    'Outer middle',
    'Cell text',
    'Cell tail',
    'Other cell',
    'Leaf text',
    'Outer end',
    'Row cell text',
    'Row other text',
    'Tail',
  ]);
}

test('native slots wrap original paragraphs, tables, nested blocks, row controls, and cell controls without a document edit', async () => {
  const previous = environment.IS_REACT_ACT_ENVIRONMENT;
  environment.IS_REACT_ACT_ENVIRONMENT = true;
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: nativeSlotDocument() });
  const originals = [...container.querySelectorAll('.docx-paragraph-fragment')];
  const fingerprint = canonicalOoxmlFingerprint(editor.surface!.session.part());
  let blocks: ReturnType<typeof host> | undefined;
  try {
    await act(async () => {
      blocks = host(editor);
    });
    structure(container);
    expect(
      [...container.querySelectorAll('.docx-paragraph-fragment')].every(
        (p, i) => p === originals[i]
      )
    ).toBe(true);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(fingerprint);
    const id = paragraph(editor, 'Inner text');
    await act(async () => {
      editor.surface!.setSelection({
        anchor: { paragraphId: id, offset: 3 },
        head: { paragraphId: id, offset: 3 },
      });
    });
    const current = [...container.querySelectorAll('.docx-paragraph-fragment')];
    const selected = container.ownerDocument.defaultView!.getSelection()!;
    const anchor = selected.anchorNode,
      offset = selected.anchorOffset;
    await act(async () => {
      blocks!.refresh();
    });
    structure(container);
    expect(selected.anchorNode === anchor).toBe(true);
    expect(selected.anchorOffset).toBe(offset);
    expect(editor.surface!.state().selection?.head).toEqual({ paragraphId: id, offset: 3 });
    await act(async () => {
      blocks!.dispose();
      await Promise.resolve();
    });
    expect(container.querySelectorAll('[data-sciobot-block-host]')).toHaveLength(0);
    expect(container.querySelectorAll('[data-sciobot-native-slot]')).toHaveLength(0);
    expect(
      [...container.querySelectorAll('.docx-paragraph-fragment')].every((p, i) => p === current[i])
    ).toBe(true);
    expect(
      container.querySelectorAll('.docx-page-content > .docx-paragraph-fragment')
    ).toHaveLength(7);
    expect(container.querySelectorAll('.docx-page-content > .docx-table-fragment')).toHaveLength(2);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(fingerprint);
  } finally {
    await act(async () => {
      blocks?.dispose();
      await Promise.resolve();
      editor.destroy();
    });
    container.remove();
    environment.IS_REACT_ACT_ENVIRONMENT = previous;
  }
});

test('native edits repaint nested slots, retain model selection, and undo without duplicate paragraphs', async () => {
  const previous = environment.IS_REACT_ACT_ENVIRONMENT;
  environment.IS_REACT_ACT_ENVIRONMENT = true;
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: nativeSlotDocument() });
  let blocks: ReturnType<typeof host> | undefined;
  try {
    await act(async () => {
      blocks = host(editor);
    });
    for (const text of ['Inner text', 'Cell text', 'Row cell text']) {
      const id = paragraph(editor, text);
      await act(async () => {
        editor.surface!.setSelection({
          anchor: { paragraphId: id, offset: text.length },
          head: { paragraphId: id, offset: text.length },
        });
        editor.surface!.insertPlainText('!');
        editor.surface!.flushPendingInput();
      });
      expect(editor.surface!.state().selection?.head).toEqual({
        paragraphId: id,
        offset: text.length + 1,
      });
      expect(
        container.querySelectorAll(`.docx-paragraph-fragment[data-paragraph-id="${id}"]`)
      ).toHaveLength(1);
      expect(
        container.querySelector(`.docx-paragraph-fragment[data-paragraph-id="${id}"]`)?.textContent
      ).toBe(text + '!');
      await act(async () => {
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      });
      structure(container);
      await act(async () => {
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      });
      expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe(text + '!');
      await act(async () => {
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      });
    }
    await act(async () => {
      blocks!.dispose();
      await Promise.resolve();
    });
    const text = [...container.querySelectorAll('.docx-paragraph-fragment')].map(
      (p) => p.textContent
    );
    expect(text).toHaveLength(12);
    expect(text.filter((value) => value === 'Inner text')).toHaveLength(1);
  } finally {
    await act(async () => {
      blocks?.dispose();
      await Promise.resolve();
      editor.destroy();
    });
    container.remove();
    environment.IS_REACT_ACT_ENVIRONMENT = previous;
  }
});
