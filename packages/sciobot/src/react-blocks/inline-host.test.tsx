import { expect, test } from 'bun:test';
import { act, createContext, useContext, useEffect, useState } from 'react';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { linesOf } from '@docx-editor.dev/core/layout';
import { canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';
import { setPaginatedSurfaceScale } from '../../../core/src/editor/surface-scale';
import { inlineWidgetDocument } from '../../test/support/react-block-document';
import { BlockUi } from './block-ui';
import { mountReactBlockHost } from './host';
import type { ReactBlockViewProps } from './host';
import { mountReactInlineHost } from './inline-host';
import type { MeasureInlineElement } from './inline-measurement';

const environment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean };
const Context = createContext('initial');

test('inline roots reuse providers, native boxes, and identity through zoom and native repaint without clearing block flow', async () => {
  const previous = environment.IS_REACT_ACT_ENVIRONMENT;
  environment.IS_REACT_ACT_ENVIRONMENT = true;
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: inlineWidgetDocument() });
  const surface = editor.surface!;
  const events = { mount: 0, unmount: 0 };
  let width = 80;
  let zoom = 1;
  const measure: MeasureInlineElement = () => ({ width, ascent: 20, descent: 4 });
  let fragmentBox: ReactBlockViewProps['fragment']['box'] | undefined;
  function Chip({ fragment }: ReactBlockViewProps) {
    fragmentBox = fragment.box;
    const context = useContext(Context),
      [count, setCount] = useState(0);
    useEffect(() => {
      events.mount++;
      return () => {
        events.unmount++;
      };
    }, []);
    return (
      <BlockUi>
        <button onClick={() => setCount((old) => old + 1)}>
          {context}:{count}
        </button>
      </BlockUi>
    );
  }
  let host: ReturnType<typeof mountReactInlineHost> | undefined;
  let blocks: ReturnType<typeof mountReactBlockHost> | undefined;
  try {
    const revision = surface.session.packageRevision(),
      fingerprint = canonicalOoxmlFingerprint(surface.session.part());
    await act(async () => {
      blocks = mountReactBlockHost(
        editor,
        new Map([
          [
            'test_tall',
            {
              view: () => <BlockUi>Tall</BlockUi>,
              reserveHeight: true,
            },
          ],
        ]),
        { measure: () => 120 * zoom }
      );
    });
    const tail = () =>
      linesOf(surface.layout()).find((line) => line.spans.some((span) => span.text === 'Tail'))!.box
        .y;
    let reservedTail = 0;
    await act(async () => {
      reservedTail = tail();
    });
    await act(async () => {
      host = mountReactInlineHost(
        editor,
        new Map([
          ['source_chip', Chip],
          ['test_tall', Chip],
        ]),
        {
          measure,
          providers: (children) => <Context.Provider value="first">{children}</Context.Provider>,
        }
      );
    });
    expect(tail()).toBeGreaterThanOrEqual(reservedTail);
    const button = container.querySelector<HTMLButtonElement>(
      '[data-sciobot-inline-measure] button'
    )!;
    const root = button.closest<HTMLElement>('[data-sciobot-block-host]')!;
    const widget = () =>
      linesOf(surface.layout())
        .flatMap((line) => line.spans)
        .find((span) => span.inlineWidget)!;
    expect(widget().inlineWidget).toMatchObject({ width: 60, ascent: 15, descent: 3 });
    expect(root.style.width).toBe('80px');
    expect(root.style.height).toBe('24px');
    expect(fragmentBox).toMatchObject({ width: 60, height: 18 });
    expect(root.style.visibility).toBe('visible');
    expect(
      container.querySelector<HTMLElement>('[data-docx-inline-widget]')!.style.visibility
    ).toBe('hidden');
    expect(surface.session.packageRevision()).toBe(revision);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(fingerprint);
    await act(async () => {
      button.click();
    });
    expect(button.textContent).toBe('first:1');
    await act(async () => {
      host!.setProviders((children) => (
        <Context.Provider value="second">{children}</Context.Provider>
      ));
      zoom = 1.5;
      setPaginatedSurfaceScale(surface, 2);
    });
    expect(container.querySelector('[data-sciobot-inline-measure] button') === button).toBe(true);
    expect(button.textContent).toBe('second:1');
    expect(widget().inlineWidget!.width).toBe(60);
    expect(root.style.width).toBe('120px');
    await act(async () => {
      width = 100;
      host!.refresh();
    });
    expect(widget().inlineWidget!.width).toBe(75);
    expect(fragmentBox).toMatchObject({ width: 75, height: 18 });
    expect(root.style.width).toBe('150px');
    expect(events).toEqual({ mount: 1, unmount: 0 });
    expect(() => mountReactInlineHost(editor)).toThrow('already has');
    await act(async () => {
      host!.dispose();
      await Promise.resolve();
    });
    expect(events).toEqual({ mount: 1, unmount: 1 });
    expect(container.querySelector('[data-docx-inline-widget]')).toBeNull();
    expect(container.textContent).toContain('Before Source after');
    await act(async () => {
      expect(tail()).toBe(reservedTail);
    });
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(fingerprint);
    expect(() => host!.register('source_chip', Chip)).toThrow('disposed');
  } finally {
    await act(async () => {
      host?.dispose();
      blocks?.dispose();
      await Promise.resolve();
      editor.destroy();
    });
    container.remove();
    environment.IS_REACT_ACT_ENVIRONMENT = previous;
  }
});

test('declined metrics retain native fallback; registration and cleanup restore it without document edits', async () => {
  const previous = environment.IS_REACT_ACT_ENVIRONMENT;
  environment.IS_REACT_ACT_ENVIRONMENT = true;
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: inlineWidgetDocument() });
  let admit = false;
  let host: ReturnType<typeof mountReactInlineHost> | undefined;
  let release: (() => void) | undefined;
  try {
    await act(async () => {
      host = mountReactInlineHost(editor, new Map(), {
        measure: () => (admit ? { width: 72, ascent: 18, descent: 2 } : null),
      });
      release = host.register('source_chip', () => <BlockUi>Chip</BlockUi>);
    });
    expect(
      container.querySelector<HTMLElement>('[data-sciobot-block-host]')!.style.visibility
    ).toBe('hidden');
    expect(container.querySelector('[data-docx-inline-widget]')).toBeNull();
    await act(async () => {
      admit = true;
      host!.refresh();
    });
    expect(
      container.querySelector<HTMLElement>('[data-docx-inline-widget]')!.style.visibility
    ).toBe('hidden');
    await act(async () => {
      release!();
      release!();
      await Promise.resolve();
    });
    expect(container.querySelector('[data-sciobot-block-host]')).toBeNull();
    expect(container.querySelector('[data-docx-inline-widget]')).toBeNull();
    expect(container.textContent).toContain('Before Source after');
  } finally {
    await act(async () => {
      host?.dispose();
      await Promise.resolve();
      editor.destroy();
    });
    container.remove();
    environment.IS_REACT_ACT_ENVIRONMENT = previous;
  }
});
