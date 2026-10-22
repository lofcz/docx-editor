import { expect, test } from 'bun:test';
import { act, useEffect } from 'react';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { reactBlockDocument, inlineWidgetDocument } from '../../test/support/react-block-document';
import { mountReactBlockHost } from './host';
import { mountReactInlineHost } from './inline-host';
import { mountInlineInput } from './inline-input';

test('twenty document switches and editor destruction unmount every retained root without host disposal', async () => {
  let mounted = 0,
    unmounted = 0;
  function View() {
    useEffect(() => {
      mounted++;
      return () => {
        unmounted++;
      };
    }, []);
    return <input aria-label="Local value" />;
  }
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: reactBlockDocument(),
  });
  const environment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean };
  const previous = environment.IS_REACT_ACT_ENVIRONMENT;
  environment.IS_REACT_ACT_ENVIRONMENT = true;
  const hosts: ReturnType<typeof mountReactBlockHost>[] = [];
  try {
    for (let i = 0; i < 20; i++) {
      await act(async () => {
        hosts.push(mountReactBlockHost(editor, new Map([['test_counter', View]])));
        mountInlineInput(editor);
      });
      expect(mounted - unmounted).toBe(1);
      await act(async () => {
        await editor.load(reactBlockDocument());
      });
      expect(mounted).toBe(i + 1);
      expect(unmounted).toBe(mounted);
    }
    await act(async () => {
      hosts.push(mountReactBlockHost(editor, new Map([['test_counter', View]])));
    });
    await act(async () => editor.destroy());
    expect(unmounted).toBe(mounted);
    await act(async () => {
      hosts.forEach((host) => host.dispose());
      editor.destroy();
    });
    expect(unmounted).toBe(21);
  } finally {
    await act(async () => editor.destroy());
    environment.IS_REACT_ACT_ENVIRONMENT = previous;
  }
});

test('inline roots and measurement owners release after twenty native document replacements', async () => {
  const environment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean };
  const previous = environment.IS_REACT_ACT_ENVIRONMENT;
  environment.IS_REACT_ACT_ENVIRONMENT = true;
  let mounted = 0,
    unmounted = 0;
  function Chip() {
    useEffect(() => {
      mounted++;
      return () => {
        unmounted++;
      };
    }, []);
    return <span>Chip</span>;
  }
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: inlineWidgetDocument(),
  });
  const hosts: ReturnType<typeof mountReactInlineHost>[] = [];
  try {
    for (let cycle = 0; cycle < 20; cycle++) {
      await act(async () => {
        hosts.push(
          mountReactInlineHost(editor, new Map([['source_chip', Chip]]), {
            measure: () => ({ width: 40, ascent: 15, descent: 3 }),
          })
        );
      });
      expect(mounted - unmounted).toBe(1);
      await act(async () => editor.load(inlineWidgetDocument()));
      expect(unmounted).toBe(mounted);
    }
    await act(async () => editor.destroy());
    await act(async () => hosts.forEach((host) => host.dispose()));
    expect(unmounted).toBe(20);
  } finally {
    await act(async () => editor.destroy());
    environment.IS_REACT_ACT_ENVIRONMENT = previous;
  }
});
