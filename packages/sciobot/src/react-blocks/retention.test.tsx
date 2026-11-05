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
  let host: ReturnType<typeof mountReactBlockHost> | undefined;
  try {
    await act(async () => {
      host = mountReactBlockHost(editor, new Map([['test_counter', View]]));
    });
    for (let i = 0; i < 20; i++) {
      await act(async () => {
        mountInlineInput(editor);
      });
      expect(mounted - unmounted).toBe(1);
      await act(async () => {
        await editor.load(reactBlockDocument());
      });
      expect(mounted).toBe(i + 2);
      expect(unmounted).toBe(i + 1);
      expect(mounted - unmounted).toBe(1);
    }
    await act(async () => editor.destroy());
    expect(unmounted).toBe(mounted);
    await act(async () => {
      host!.dispose();
      editor.destroy();
    });
    expect(unmounted).toBe(21);
  } finally {
    await act(async () => {
      host?.dispose();
      editor.destroy();
    });
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
  let host: ReturnType<typeof mountReactInlineHost> | undefined;
  try {
    await act(async () => {
      host = mountReactInlineHost(editor, new Map([['source_chip', Chip]]), {
        measure: () => ({ width: 40, ascent: 15, descent: 3 }),
      });
    });
    for (let cycle = 0; cycle < 20; cycle++) {
      expect(mounted - unmounted).toBe(1);
      await act(async () => editor.load(inlineWidgetDocument()));
      expect(mounted).toBe(cycle + 2);
      expect(unmounted).toBe(cycle + 1);
      expect(mounted - unmounted).toBe(1);
      expect(() => mountReactInlineHost(editor)).toThrow('already has a React inline host');
    }
    await act(async () => editor.destroy());
    await act(async () => host!.dispose());
    expect(unmounted).toBe(21);
  } finally {
    await act(async () => {
      host?.dispose();
      editor.destroy();
    });
    environment.IS_REACT_ACT_ENVIRONMENT = previous;
  }
});
