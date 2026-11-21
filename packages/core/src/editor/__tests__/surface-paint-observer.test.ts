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

test('passive observers can subscribe to the current frame without repainting other hosts', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docx(p('Text')),
  });
  try {
    const calls: string[] = [];
    const surface = editor.surface!;
    surface.observePaint({ beforePaint() {}, afterPaint: () => calls.push('host') });
    surface.observePaint({ afterPaint: () => calls.push('disclosure') });
    surface.observePaint({ afterPaint: () => calls.push('image') });
    expect(calls).toEqual(['host', 'disclosure', 'image']);
  } finally {
    editor.destroy();
  }
});

test('the surface reads its viewport before a host prepares its widgets for paint', () => {
  const scroller = document.createElement('div');
  scroller.className = 'docx-editor__scroll-container';
  const container = document.createElement('div');
  scroller.append(container);
  document.body.append(scroller);
  let preparing = false;
  const reads: boolean[] = [];
  Object.defineProperty(scroller, 'clientHeight', {
    get: () => {
      reads.push(preparing);
      return 640;
    },
  });
  const editor = createDocxEditor({ container, document: docx(p('Text')) });
  try {
    editor.surface!.observePaint({
      beforePaint: () => {
        preparing = true;
      },
      afterPaint: () => {
        preparing = false;
      },
    });
    expect(reads).toContain(false);
    expect(reads).not.toContain(true);
  } finally {
    editor.destroy();
    scroller.remove();
  }
});

test('a host can retain its paragraph roots while paint replaces edited text', () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: docx(p('Alpha') + p('Beta')) });
  const surface = editor.surface!;
  const root = () => container.querySelector<HTMLElement>('.docx-paragraph-fragment')!;
  let marker: Comment | null = null;
  let slot: HTMLElement | null = null;
  let borrowed: HTMLElement | null = null;
  const accepted: HTMLElement[][] = [];
  const retainedPages: boolean[] = [];
  let decline = false;
  let expired: (() => boolean) | undefined;
  const stop = surface.observePaint({
    beforePaint(kept, preserve) {
      if (preserve) expired = () => preserve(0, () => true);
      if (
        preserve?.(0, (changed) => {
          if (!borrowed || !changed.has(borrowed) || decline) return false;
          accepted.push([...changed]);
          retainedPages.push(kept?.has(0) ?? false);
          return true;
        })
      )
        return;
      if (borrowed && marker?.parentNode) marker.replaceWith(borrowed);
      slot?.remove();
      marker = slot = borrowed = null;
    },
    afterPaint() {
      if (borrowed) return;
      borrowed = root();
      marker = document.createComment('paragraph');
      slot = document.createElement('section');
      borrowed.before(marker, slot);
      slot.append(borrowed);
    },
  });
  try {
    const original = root(),
      originalSlot = slot;
    const id = surface.session.paragraphIds()[0]!;
    surface.setSelection({
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 0 },
    });
    surface.type('x');
    expect(accepted).toContainEqual([original]);
    expect(retainedPages).not.toContain(true);
    expect(root()).toBe(original);
    expect(root().parentElement).toBe(originalSlot);
    expect(root().textContent).toContain('xAlpha');
    expect(expired?.()).toBe(false);
    decline = true;
    surface.type('y');
    expect(root()).not.toBe(original);
    expect(root().textContent).toContain('xyAlpha');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(paragraphTextOf(surface.session.part(), id)).toBe('xAlpha');
  } finally {
    stop();
    editor.destroy();
    container.remove();
  }
});

test('field visibility keeps an opted-in body loan and its field widget connected', () => {
  const container = document.createElement('div');
  document.body.append(container);
  const field = `<w:sdt><w:sdtPr><w:alias w:val="Choice"/><w:dropDownList/></w:sdtPr><w:sdtContent>${p('Alpha')}</w:sdtContent></w:sdt>`;
  const editor = createDocxEditor({ container, document: docx(field + p('Beta')) });
  const surface = editor.surface!;
  const paragraphId = surface.session.paragraphIds()[1]!;
  const point = { paragraphId, offset: 2 };
  surface.setSelection({ anchor: point, head: point });
  const decisions: boolean[] = [];
  const whole: boolean[] = [];
  const root = container.querySelector<HTMLElement>('.docx-paragraph-fragment')!;
  const widget = container.querySelector<HTMLElement>('.docx-content-control-widget')!;
  const chrome = widget.closest<HTMLElement>('.docx-content-control-chrome')!;
  const slot = document.createElement('section');
  const marker = document.createComment('borrowed paragraph');
  const stop = surface.observePaint({
    beforePaint(retained, reserve) {
      if (root.parentNode !== slot) return;
      whole.push(retained?.has(0) === true);
      const accepted = reserve?.(0, (changed) => changed.size === 0) === true;
      decisions.push(accepted);
      if (!accepted) marker.replaceWith(root);
    },
    afterPaint() {
      if (root.parentNode === slot) return;
      root.before(marker, slot);
      slot.append(root);
    },
  });
  try {
    const layout = surface.publishedLayout();
    for (const show of [true, false]) {
      surface.contentControls.setShowAll(show);
      expect(container.querySelector('.docx-content-control-widget')).toBe(widget);
      expect(chrome.hasAttribute('data-boundary-visible')).toBe(show);
      expect(root.parentElement).toBe(slot);
      expect(surface.publishedLayout()).toBe(layout);
      expect(surface.state().selection).toEqual({ anchor: point, head: point });
    }
    expect(whole).toEqual([false, false]);
    expect(decisions).toEqual([true, true]);
  } finally {
    stop();
    editor.destroy();
    container.remove();
  }
});

test('changed tables, new paragraphs and invalid painted roots use normal paint', () => {
  const table =
    '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>' +
    p('Cell') +
    '</w:tc></w:tr></w:tbl>';
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: docx(p('Alpha') + table) });
  const surface = editor.surface!;
  let decisions: boolean[] = [];
  const stop = surface.observePaint({
    beforePaint(_kept, preserve) {
      if (preserve) decisions.push(preserve(0, () => true));
    },
    afterPaint() {},
  });
  try {
    const [first, cell] = surface.session.paragraphIds();
    const point = { paragraphId: cell!, offset: 2 };
    surface.setSelection({ anchor: point, head: point });
    decisions = [];
    surface.type('x');
    expect(decisions).toContain(false);
    expect(decisions).not.toContain(true);
    expect(paragraphTextOf(surface.session.part(), cell!)).toBe('Cexll');
    const start = { paragraphId: first!, offset: 1 };
    surface.setSelection({ anchor: start, head: start });
    decisions = [];
    surface.splitParagraph();
    expect(decisions).toContain(false);
    expect(decisions).not.toContain(true);
    const root = container.querySelector<HTMLElement>('.docx-table-fragment')!;
    root.dataset.tableId = 'stale';
    decisions = [];
    surface.type('z');
    expect(decisions).toContain(false);
    expect(decisions).not.toContain(true);
  } finally {
    stop();
    editor.destroy();
    container.remove();
  }
});
