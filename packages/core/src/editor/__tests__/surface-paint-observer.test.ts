import { expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor.ts';
import { docx, p } from '../../automation/__tests__/support/protocol.ts';
import { paragraphTextOf } from '../../store/store/tree-op-apply.ts';

test('a host can wrap editable native content between paints and unsubscribe cleanly', () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: docx(p('Alpha') + p('Beta')) });
  const surface = editor.surface!;
  let node: HTMLElement | undefined, parent: HTMLElement | undefined;
  const calls: string[] = [];
  const stop = surface.observePaint({
    beforePaint() {
      calls.push('before');
      if (node && parent) parent.append(node);
    },
    afterPaint(frame) {
      calls.push('after');
      node = frame.pagesLayer.querySelector<HTMLElement>('[data-paragraph-id]')!;
      parent = node.parentElement!;
      const wrapper = document.createElement('section');
      wrapper.dataset.testReactBlock = '';
      parent.append(wrapper);
      wrapper.append(node);
    },
  });
  try {
    expect(calls).toEqual(['before', 'after']);
    const id = node!.dataset.paragraphId!;
    surface.setSelection({
      anchor: { paragraphId: id, offset: 5 },
      head: { paragraphId: id, offset: 5 },
    });
    surface.insertPlainText('!');
    surface.flushPendingInput();
    expect(paragraphTextOf(surface.session.part(), id)).toBe('Alpha!');
    expect(node!.parentElement!.dataset.testReactBlock).toBe('');
    expect(node!.textContent).toContain('Alpha!');
    stop();
    const count = calls.length;
    surface.insertPlainText('?');
    surface.flushPendingInput();
    expect(calls).toHaveLength(count);
    expect(paragraphTextOf(surface.session.part(), id)).toBe('Alpha!?');
  } finally {
    editor.destroy();
    container.remove();
  }
});
