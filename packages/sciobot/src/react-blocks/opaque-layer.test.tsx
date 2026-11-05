import { expect, test } from 'bun:test';
import { act } from 'react';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { blockMeasurementDocument } from '../../test/support/react-block-document';
import { mountReactBlockHost } from './host';

test('opaque fragments keep connected roots through repaint and hide outside materialized pages', async () => {
  const environment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean };
  const previous = environment.IS_REACT_ACT_ENVIRONMENT;
  environment.IS_REACT_ACT_ENVIRONMENT = true;
  const scroller = document.createElement('div');
  scroller.className = 'docx-editor__scroll-container';
  Object.defineProperty(scroller, 'clientHeight', { value: 800, configurable: true });
  const container = document.createElement('div');
  scroller.append(container);
  document.body.append(scroller);
  const editor = createDocxEditor({
    container,
    document: blockMeasurementDocument(Array.from({ length: 180 }, (_, index) => `Line ${index}`)),
  });
  let host: ReturnType<typeof mountReactBlockHost> | undefined;
  const settle = async () =>
    act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 50));
    });
  try {
    await act(async () => {
      host = mountReactBlockHost(
        editor,
        new Map([['test_tall', () => <input defaultValue="kept" />]])
      );
    });
    await settle();
    const surface = editor.surface!;
    const revision = surface.session.packageRevision();
    expect(surface.layout().pages.length).toBeGreaterThan(2);
    const first = container.querySelector<HTMLElement>('[data-sciobot-block-host="tall:0"]')!;
    expect(first).not.toBeNull();
    expect(first.closest('.docx-pages')).toBeNull();
    expect(first.closest('[data-sciobot-block-layer]')?.parentElement).toBe(surface.hostElement());
    expect(first.style.clipPath).toStartWith('inset(');
    const input = first.querySelector('input')!;
    let blurs = 0;
    input.addEventListener('blur', () => blurs++);
    input.focus();
    await act(async () => {
      host!.refresh();
      editor.setZoom(0.75);
    });
    await settle();
    expect(blurs).toBe(0);
    expect(input.ownerDocument.activeElement).toBe(input);
    expect(container.querySelector('[data-sciobot-block-host="tall:0"]')).toBe(first);
    const lastPage = surface.layout().pages.at(-1)!;
    await act(async () => {
      const tail = surface.session.paragraphIds().at(-1)!;
      surface.setSelection({
        anchor: { paragraphId: tail, offset: 0 },
        head: { paragraphId: tail, offset: 0 },
      });
      scroller.scrollTop = lastPage.box.y;
      scroller.dispatchEvent(new Event('scroll'));
    });
    await settle();
    expect(first.hidden).toBe(true);
    expect(first.style.display).toBe('none');
    expect(first.isConnected).toBe(true);
    const visible = [
      ...container.querySelectorAll<HTMLElement>('[data-sciobot-block-host]'),
    ].filter((root) => !root.hidden);
    expect(visible.length).toBeGreaterThan(0);
    expect(visible.every((root) => root.parentElement === first.parentElement)).toBe(true);
    expect(surface.session.packageRevision()).toBe(revision);
    await act(async () => {
      scroller.scrollTop = 0;
      scroller.dispatchEvent(new Event('scroll'));
    });
    await settle();
    expect(first.hidden).toBe(false);
    expect(first.querySelector('input')).toBe(input);
    await act(async () => {
      host!.dispose();
    });
    expect(container.querySelector('[data-sciobot-block-layer]')).toBeNull();
    expect(first.isConnected).toBe(false);
  } finally {
    await act(async () => {
      host?.dispose();
      editor.destroy();
    });
    scroller.remove();
    environment.IS_REACT_ACT_ENVIRONMENT = previous;
  }
});
