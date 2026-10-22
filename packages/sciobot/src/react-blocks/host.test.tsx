import { expect, test } from 'bun:test';
import { act, useEffect, useState } from 'react';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { canonicalOoxmlFingerprint, paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { reactBlockDocument } from '../../test/support/react-block-document';
import { buildBlockIndex } from '../blocks/block-index';
import { writeBlockPayload } from '../blocks/payload-part';
import { BlockUi } from './block-ui';
import { mountReactBlockHost, type ReactBlockViewProps } from './host';

const environment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean };
function probe() {
  const events = { mounted: 0, unmounted: 0 };
  function Counter({ block, paragraphs }: ReactBlockViewProps) {
    const [count, setCount] = useState(0);
    useEffect(() => {
      events.mounted++;
      return () => {
        events.unmounted++;
      };
    }, []);
    return (
      <BlockUi>
        <button type="button" onClick={() => setCount((old) => old + 1)}>
          Count {count}
        </button>
        <span data-name="">{String(block.payload.attributes.name)}</span>
        <output>{paragraphs.map((p) => p.text).join('|')}</output>
      </BlockUi>
    );
  }
  return { events, Counter };
}
function paragraph(editor: DocxEditorInstance, text: string) {
  const session = editor.surface!.session;
  const id = session.paragraphIds().find((id) => paragraphTextOf(session.part(), id) === text);
  if (!id) throw new Error(`Missing paragraph ${text}`);
  return id;
}
function rename(editor: DocxEditorInstance, name: string) {
  const block = buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('counter-1')!;
  return writeBlockPayload(editor, {
    controlId: block.controlId,
    payload: { ...block.payload, attributes: { name } },
    label: name,
    fallbackOps: [],
  });
}
const click = (button: Element) => button.dispatchEvent(new MouseEvent('click', { bubbles: true }));

test('registered roots retain identity and local state across native and payload edits; disposal restores native-only rendering', async () => {
  const previous = environment.IS_REACT_ACT_ENVIRONMENT;
  environment.IS_REACT_ACT_ENVIRONMENT = true;
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: reactBlockDocument() });
  const widget = probe();
  let host: ReturnType<typeof mountReactBlockHost> | undefined;
  try {
    const fingerprint = canonicalOoxmlFingerprint(editor.surface!.session.part());
    await act(async () => {
      host = mountReactBlockHost(editor, new Map([['test_counter', widget.Counter]]));
    });
    expect(widget.events).toEqual({ mounted: 1, unmounted: 0 });
    expect(container.querySelectorAll('[data-sciobot-block-host]')).toHaveLength(1);
    const button = container.querySelector('[data-sciobot-block-host] button')!;
    await act(async () => {
      click(button);
    });
    expect(button.textContent).toBe('Count 1');
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(fingerprint);
    const id = paragraph(editor, 'Editable body');
    await act(async () => {
      editor.surface!.setSelection({
        anchor: { paragraphId: id, offset: 13 },
        head: { paragraphId: id, offset: 13 },
      });
      editor.surface!.insertPlainText('!');
      editor.surface!.flushPendingInput();
    });
    expect(container.querySelector('[data-sciobot-block-host] button') === button).toBe(true);
    expect(button.textContent).toBe('Count 1');
    expect(container.querySelector('[data-sciobot-block-host] output')?.textContent).toBe(
      'Widget fallback|Editable body!'
    );
    await act(async () => {
      expect(rename(editor, 'Updated').success).toBe(true);
    });
    expect(container.querySelector('[data-sciobot-block-host] [data-name]')?.textContent).toBe(
      'Updated'
    );
    expect(widget.events).toEqual({ mounted: 1, unmounted: 0 });
    expect(container.querySelectorAll('.docx-paragraph-fragment')).toHaveLength(5);
    expect(container.textContent).toContain('Unregistered fallback');
    await act(async () => {
      host!.dispose();
      host!.dispose();
      await Promise.resolve();
    });
    expect(widget.events).toEqual({ mounted: 1, unmounted: 1 });
    expect(container.querySelectorAll('[data-sciobot-block-host]')).toHaveLength(0);
    expect(container.textContent).toContain('Editable body!');
    expect(container.textContent).toContain('Unregistered fallback');
    expect(() => host!.register('later', widget.Counter)).toThrow('disposed');
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

test('dynamic registration and unregistration mount only matching payload types', async () => {
  const previous = environment.IS_REACT_ACT_ENVIRONMENT;
  environment.IS_REACT_ACT_ENVIRONMENT = true;
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: reactBlockDocument() });
  const widget = probe();
  let host: ReturnType<typeof mountReactBlockHost> | undefined;
  try {
    await act(async () => {
      host = mountReactBlockHost(editor);
    });
    expect(container.querySelectorAll('[data-sciobot-block-host]')).toHaveLength(0);
    let stop: (() => void) | undefined;
    await act(async () => {
      stop = host!.register('test_counter', widget.Counter);
    });
    expect(widget.events.mounted).toBe(1);
    expect(() => host!.register('test_counter', widget.Counter)).toThrow('already registered');
    expect(() => host!.register('', widget.Counter)).toThrow('Invalid');
    await act(async () => {
      stop!();
      stop!();
      await Promise.resolve();
    });
    expect(widget.events.unmounted).toBe(1);
    expect(container.querySelectorAll('[data-sciobot-block-host]')).toHaveLength(0);
    expect(container.textContent).toContain('Widget fallback');
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

for (const receiver of [0, 1] as const)
  test(`two real sessions retain widget roots through reconnect, peer-preserving history, deletion, and save; receiver ${receiver}`, async () => {
    const previous = environment.IS_REACT_ACT_ENVIRONMENT;
    environment.IS_REACT_ACT_ENVIRONMENT = true;
    const pair = await createTwoEditorFixture({
      document: reactBlockDocument(),
      clientIds: [73, 41],
    });
    const widgets = [probe(), probe()];
    const hosts: ReturnType<typeof mountReactBlockHost>[] = [];
    try {
      await act(async () => {
        pair.editors.forEach((editor, i) =>
          hosts.push(mountReactBlockHost(editor, new Map([['test_counter', widgets[i]!.Counter]])))
        );
      });
      const buttons = pair.containers.map(
        (node) => node.querySelector('[data-sciobot-block-host] button')!
      );
      await act(async () => {
        click(buttons[0]!);
        click(buttons[1]!);
        click(buttons[1]!);
      });
      const [alice, bob] = pair.editors;
      const id = paragraph(bob!, 'Editable body');
      pair.pause();
      await act(async () => {
        expect(rename(alice!, 'Peer shared').success).toBe(true);
        bob!.surface!.setSelection({
          anchor: { paragraphId: id, offset: 13 },
          head: { paragraphId: id, offset: 13 },
        });
        bob!.surface!.insertPlainText(' Bob');
        bob!.surface!.flushPendingInput();
        pair.reconnect(receiver);
        await pair.settle();
      });
      for (const container of pair.containers) {
        expect(container.querySelector('[data-name]')?.textContent).toBe('Peer shared');
        expect(container.querySelector('output')?.textContent).toContain('Editable body Bob');
      }
      expect(buttons.map((button) => button.textContent)).toEqual(['Count 1', 'Count 2']);
      expect(widgets.map((widget) => widget.events.mounted)).toEqual([1, 1]);
      const remote = pair.rooms[0]!.session.remoteSelections().find(
        (selection) => selection.actorId === 'bob'
      )!;
      expect(Boolean(remote)).toBe(true);
      const a = caretAt(alice!.surface!.layout(), {
        paragraphId: remote.head.nodeId,
        offset: remote.head.offset,
      });
      const b = caretAt(bob!.surface!.layout(), {
        paragraphId: remote.head.nodeId,
        offset: remote.head.offset,
      });
      expect(a !== null && b !== null).toBe(true);
      expect({ x: a!.x, y: a!.y }).toEqual({ x: b!.x, y: b!.y });
      await act(async () => {
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
      });
      for (const container of pair.containers) {
        expect(container.querySelector('[data-name]')?.textContent).toBe('Initial');
        expect(container.querySelector('output')?.textContent).toContain('Editable body Bob');
      }
      await act(async () => {
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
      });
      expect(buttons.map((button) => button.textContent)).toEqual(['Count 1', 'Count 2']);
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
      pair.pause();
      await act(async () => {
        bob!.surface!.insertPlainText(' Still editing');
        bob!.surface!.flushPendingInput();
        const block = buildBlockIndex(alice!.surface!.session.currentPackage()).byId.get(
          'counter-1'
        )!;
        expect(
          alice!.surface!.applyAutomationOps(() => [
            { op: 'removeContentControl', controlId: block.controlId, keepContent: false },
          ]).rejected
        ).toBe(false);
        pair.reconnect(receiver);
        await pair.settle();
        await Promise.resolve();
      });
      for (const container of pair.containers)
        expect(container.querySelectorAll('[data-sciobot-block-host]')).toHaveLength(0);
      expect(widgets.map((widget) => widget.events.unmounted)).toEqual([1, 1]);
      const deleted = await pair.savedDigests();
      expect(deleted[0]).toEqual(deleted[1]);
    } finally {
      await act(async () => {
        hosts.forEach((host) => host.dispose());
        await Promise.resolve();
        pair.destroy();
      });
      environment.IS_REACT_ACT_ENVIRONMENT = previous;
    }
  }, 30_000);
