import { expect, test } from 'bun:test';
import { act, createContext, useContext, useState } from 'react';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { inlineWidgetDocument, reactBlockDocument } from '../../test/support/react-block-document';
import { buildBlockIndex } from '../blocks/block-index';
import { writeBlockPayload } from '../blocks/payload-part';
import { removeBlock } from '../blocks/remove';
import { BlockUi } from './block-ui';
import { mountReactBlockHost, type ReactBlockViewProps } from './host';
import { mountReactInlineHost } from './inline-host';

const environment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean };
const Locale = createContext('initial');
function Widget({ block }: ReactBlockViewProps) {
  const locale = useContext(Locale);
  const [count, setCount] = useState(0);
  return (
    <BlockUi>
      <button onClick={() => setCount((value) => value + 1)}>
        {locale}:{String(block.payload.attributes.name)}:{count}
      </button>
    </BlockUi>
  );
}

for (const inline of [false, true])
  test(`${inline ? 'inline' : 'block'} hosts retain registrations and providers across loads, with fresh document state`, async () => {
    const previous = environment.IS_REACT_ACT_ENVIRONMENT;
    environment.IS_REACT_ACT_ENVIRONMENT = true;
    const container = document.createElement('div');
    const moved = document.createElement('div');
    document.body.append(container, moved);
    const fixture = inline ? inlineWidgetDocument : reactBlockDocument;
    const editor = createDocxEditor({ container, document: fixture() });
    let host:
      ReturnType<typeof mountReactBlockHost> | ReturnType<typeof mountReactInlineHost> | undefined;
    try {
      let release: () => void;
      await act(async () => {
        host = inline
          ? mountReactInlineHost(editor, new Map(), {
              measure: () => ({ width: 80, ascent: 20, descent: 4 }),
            })
          : mountReactBlockHost(editor);
        release = host.register(inline ? 'source_chip' : 'test_counter', Widget);
        host.setProviders((children) => <Locale.Provider value="kept">{children}</Locale.Provider>);
      });
      const first = container.querySelector<HTMLButtonElement>('[data-sciobot-block-host] button')!;
      await act(async () => {
        first.click();
      });
      expect(first.textContent).toBe('kept:Initial:1');
      await act(async () => {
        editor.detach();
        editor.attach(moved);
      });
      expect(moved.querySelector('[data-sciobot-block-host] button')).toBe(first);
      await act(async () => {
        editor.load(fixture());
      });
      const second = moved.querySelector<HTMLButtonElement>('[data-sciobot-block-host] button')!;
      expect(second).not.toBeNull();
      expect(second).not.toBe(first);
      expect(first.isConnected).toBe(false);
      expect(second.textContent).toBe('kept:Initial:0');
      if (inline) expect(moved.querySelector('[data-docx-inline-widget]')).not.toBeNull();
      await act(async () => {
        release!();
      });
      expect(moved.querySelector('[data-sciobot-block-host]')).toBeNull();
      await act(async () => {
        host!.register(inline ? 'source_chip' : 'test_counter', Widget);
        release!();
      });
      expect(moved.querySelector('[data-sciobot-block-host] button')).not.toBeNull();
      await act(async () => {
        host!.dispose();
        editor.load(fixture());
      });
      expect(moved.querySelector('[data-sciobot-block-host]')).toBeNull();
    } finally {
      await act(async () => {
        host?.dispose();
        editor.destroy();
      });
      container.remove();
      moved.remove();
      environment.IS_REACT_ACT_ENVIRONMENT = previous;
    }
  }, 30_000);

for (const receiver of [0, 1] as const)
  test(`a reloaded block host follows shared edits, history, deletion, and reopened content; receiver ${receiver}`, async () => {
    const previous = environment.IS_REACT_ACT_ENVIRONMENT;
    environment.IS_REACT_ACT_ENVIRONMENT = true;
    const pair = await createTwoEditorFixture({
      document: reactBlockDocument(),
      clientIds: [41, 73],
    });
    const hosts: ReturnType<typeof mountReactBlockHost>[] = [];
    try {
      await act(async () => {
        for (const editor of pair.editors)
          hosts.push(mountReactBlockHost(editor, new Map([['test_counter', Widget]])));
      });
      const editor = pair.editors[receiver]!;
      const peer = pair.editors[1 - receiver]!;
      const bytes = new Uint8Array(await editor.save());
      const old = pair.containers[receiver]!.querySelector('[data-sciobot-block-host] button');
      await act(async () => {
        editor.load(bytes);
        await pair.settle();
      });
      expect(
        pair.containers[receiver]!.querySelector('[data-sciobot-block-host] button')
      ).not.toBeNull();
      expect(pair.containers[receiver]!.querySelector('[data-sciobot-block-host] button')).not.toBe(
        old
      );
      pair.pause();
      await act(async () => {
        const block = buildBlockIndex(peer.surface!.session.currentPackage()).byId.get(
          'counter-1'
        )!;
        expect(
          writeBlockPayload(peer, {
            controlId: block.controlId,
            payload: { ...block.payload, attributes: { name: 'Peer' } },
            label: 'Peer',
            fallbackOps: [],
          }).success
        ).toBe(true);
        const surface = editor.surface!;
        const id = surface.session
          .paragraphIds()
          .find((id) => paragraphTextOf(surface.session.part(), id) === 'Tail')!;
        const point = { paragraphId: id, offset: 4 };
        surface.setSelection({ anchor: point, head: point });
        surface.insertPlainText(' kept');
        surface.flushPendingInput();
        pair.reconnect(receiver);
        await pair.settle();
      });
      for (const container of pair.containers) {
        expect(container.querySelector('[data-sciobot-block-host] button')?.textContent).toBe(
          'initial:Peer:0'
        );
        expect(container.textContent).toContain('Tail kept');
      }
      await act(async () => {
        expect(peer.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
      });
      for (const container of pair.containers)
        expect(container.querySelector('[data-sciobot-block-host] button')?.textContent).toBe(
          'initial:Initial:0'
        );
      await act(async () => {
        expect(peer.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
      });
      await act(async () => {
        const block = buildBlockIndex(peer.surface!.session.currentPackage()).byId.get(
          'counter-1'
        )!;
        expect(removeBlock(peer, { controlId: block.controlId, mode: 'delete' }).success).toBe(
          true
        );
        await pair.settle();
      });
      for (const container of pair.containers)
        expect(container.querySelector('[data-sciobot-block-host]')).toBeNull();
      await act(async () => {
        expect(peer.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
      });
      for (const container of pair.containers)
        expect(container.querySelector('[data-sciobot-block-host] button')?.textContent).toBe(
          'initial:Peer:0'
        );
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      await act(async () => {
        hosts.forEach((host) => host.dispose());
        pair.destroy();
      });
      environment.IS_REACT_ACT_ENVIRONMENT = previous;
    }
  }, 30_000);
