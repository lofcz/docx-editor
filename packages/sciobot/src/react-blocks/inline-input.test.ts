import { expect, test } from 'bun:test';
import { createDocxEditor, setSurfaceInlineWidgetMetrics } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, paragraphTextOf } from '@docx-editor.dev/core/store';
import { inlineWidgetDocument } from '../../test/support/react-block-document';
import { buildBlockIndex } from '../blocks/block-index';
import { mountBlockClipboard, SCIOBOT_BLOCKS_MIME } from './clipboard';
import { mountInlineInput } from './inline-input';

function mounted() {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: inlineWidgetDocument() });
  const surface = editor.surface!,
    paragraphId = surface.session.paragraphIds()[0]!;
  setSurfaceInlineWidgetMetrics(surface, {
    measure: () => ({ width: 72, ascent: 18, descent: 2 }),
    revision: '1',
  });
  return {
    editor,
    surface,
    container,
    pages: container.querySelector<HTMLElement>('.docx-pages')!,
    select: (anchor: number, head = anchor) =>
      surface.setSelection({
        anchor: { paragraphId, offset: anchor },
        head: { paragraphId, offset: head },
      }),
    text: () => paragraphTextOf(surface.session.part(), paragraphId)!,
    fingerprint: () => canonicalOoxmlFingerprint(surface.session.part().root),
    destroy() {
      editor.destroy();
      container.remove();
    },
  };
}
function transfer(initial: Record<string, string> = {}) {
  const data = new Map(Object.entries(initial));
  return {
    getData: (type: string) => data.get(type) ?? '',
    setData: (type: string, value: string) => {
      data.set(type, value);
    },
    get types() {
      return [...data.keys()];
    },
  };
}
function clipboard(type: string, target: HTMLElement, data = transfer()) {
  const event = new Event(type, { bubbles: true, cancelable: true });
  Object.defineProperty(event, 'clipboardData', { value: data });
  target.dispatchEvent(event);
  return { event, data };
}

test('partial copy expands to one atom; shared clipboard leases paste once and dispose independently', () => {
  const m = mounted();
  let sequence = 0,
    globalPastes = 0,
    inlinePastes = 0;
  const createId = () => `copied-${++sequence}`;
  const global = mountBlockClipboard(m.editor, {
    createId,
    onPaste: () => {
      globalPastes++;
    },
  });
  const input = mountInlineInput(m.editor, {
    createId,
    onPaste: () => {
      inlinePastes++;
    },
  });
  try {
    const before = m.fingerprint();
    m.select(8, 11);
    const copy = clipboard('copy', m.pages).data;
    expect(copy.getData('text/plain')).toBe('Source');
    expect(JSON.parse(copy.getData(SCIOBOT_BLOCKS_MIME)).blocks[0].payload.attributes).toEqual({
      name: 'Initial',
    });
    expect(m.surface.state().selection).toMatchObject({
      anchor: { offset: 7 },
      head: { offset: 13 },
    });
    expect(m.fingerprint()).toBe(before);
    m.select(m.text().length);
    clipboard('paste', m.pages, copy);
    expect(m.text()).toBe('Before Source afterSource');
    expect([globalPastes, inlinePastes, sequence]).toEqual([1, 1, 1]);
    expect(
      buildBlockIndex(m.surface.session.currentPackage()).byId.get('copied-1')!.payload.attributes
    ).toEqual({ name: 'Initial' });
    expect(() => mountBlockClipboard(m.editor, { createId: () => 'conflict' })).toThrow(
      'different block ID factory'
    );
    global.dispose();
    global.dispose();
    m.select(m.text().length);
    clipboard('paste', m.pages, copy);
    expect(m.text()).toBe('Before Source afterSourceSource');
    expect([globalPastes, inlinePastes, sequence]).toEqual([1, 2, 2]);
    input.dispose();
    input.dispose();
    m.select(8, 11);
    const ordinary = clipboard('copy', m.pages).data;
    expect(ordinary.getData('text/plain')).toBe('our');
    expect(ordinary.getData(SCIOBOT_BLOCKS_MIME)).toBe('');
  } finally {
    input.dispose();
    global.dispose();
    m.destroy();
  }
});

test('widget form cut and rich paste keep clipboard metadata and native content isolated', () => {
  const m = mounted(),
    input = mountInlineInput(m.editor);
  try {
    m.select(7, 13);
    const copy = clipboard('copy', m.pages).data;
    const ui = document.createElement('input');
    ui.dataset.sciobotBlockUi = '';
    ui.dataset.docxHostUi = '';
    ui.value = 'local';
    m.pages.append(ui);
    ui.focus();
    const before = m.fingerprint(),
      revision = m.surface.session.packageRevision();
    expect(clipboard('cut', ui).data.getData(SCIOBOT_BLOCKS_MIME)).toBe('');
    expect(clipboard('paste', ui, copy).event.defaultPrevented).toBe(false);
    expect(m.fingerprint()).toBe(before);
    expect(m.surface.session.packageRevision()).toBe(revision);
    expect(ui.value).toBe('local');
  } finally {
    input.dispose();
    m.destroy();
  }
});

test('partial native replacement and deletion cannot split a measured payload control', () => {
  const m = mounted(),
    input = mountInlineInput(m.editor);
  try {
    const before = m.fingerprint();
    m.select(8, 11);
    m.pages.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Backspace', bubbles: true, cancelable: true })
    );
    m.surface.flushPendingInput();
    expect(m.text()).toBe('Before  after');
    expect(buildBlockIndex(m.surface.session.currentPackage()).byId.has('source-1')).toBe(false);
    expect(m.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(m.fingerprint()).toBe(before);
    m.select(11, 8);
    m.pages.dispatchEvent(
      new InputEvent('beforeinput', {
        inputType: 'insertText',
        data: 'X',
        bubbles: true,
        cancelable: true,
      })
    );
    m.surface.flushPendingInput();
    expect(m.text()).toBe('Before X after');
    expect(buildBlockIndex(m.surface.session.currentPackage()).byId.has('source-1')).toBe(false);
    expect(m.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(m.fingerprint()).toBe(before);
  } finally {
    input.dispose();
    m.destroy();
  }
});
