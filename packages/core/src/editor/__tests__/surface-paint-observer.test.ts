import { expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor.ts';
import { docx, p } from '../../automation/__tests__/support/protocol.ts';
import { paragraphTextOf } from '../../store/store/tree-op-apply.ts';
import { destroyPaintObservers } from '../surface-paint-observer.ts';

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

test('document replacement and editor destruction release each surface observer once', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docx(p('Old')),
  });
  const calls: string[] = [];
  const old = editor.surface!;
  const stop = old.observePaint({
    afterPaint() {},
    beforePaint: () => calls.push('restore'),
    destroyed: () => calls.push('destroy'),
  });
  calls.length = 0;
  await editor.load(docx(p('New')));
  expect(editor.surface).not.toBe(old);
  expect(calls).toEqual(['restore', 'destroy']);
  stop();
  stop();
  expect(calls).toEqual(['restore', 'destroy']);
  let destroyed = 0;
  editor.surface!.observePaint({
    afterPaint() {},
    destroyed: () => {
      destroyed++;
    },
  });
  editor.destroy();
  editor.destroy();
  expect(destroyed).toBe(1);
});

test('stopping a paint observer restores borrowed content only once and ends its lifetime', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docx(p('Text')),
  });
  let restored = 0,
    destroyed = 0;
  const stop = editor.surface!.observePaint({
    afterPaint() {},
    beforePaint: () => {
      restored++;
    },
    destroyed: () => {
      destroyed++;
    },
  });
  const baseline = restored;
  stop();
  stop();
  editor.destroy();
  expect(restored).toBe(baseline + 1);
  expect(destroyed).toBe(0);
});

test('one failing observer still releases all other surface resources', () => {
  const calls: string[] = [],
    errors: unknown[] = [],
    report = console.error;
  console.error = (...error) => {
    errors.push(error);
  };
  try {
    const observers = new Set([
      {
        afterPaint() {},
        beforePaint() {
          throw new Error('restore');
        },
        destroyed() {
          calls.push('first');
          throw new Error('dispose');
        },
      },
      {
        afterPaint() {},
        destroyed() {
          calls.push('second');
        },
      },
    ]);
    destroyPaintObservers(observers);
    destroyPaintObservers(observers);
    expect(calls).toEqual(['first', 'second']);
    expect(errors).toHaveLength(2);
    expect(observers.size).toBe(0);
  } finally {
    console.error = report;
  }
});
