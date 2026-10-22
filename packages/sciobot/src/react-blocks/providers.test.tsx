import { expect, test } from 'bun:test';
import { act, createContext, useContext, useState } from 'react';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { reactBlockDocument } from '../../test/support/react-block-document';
import { buildBlockIndex } from '../blocks/block-index';
import { writeBlockPayload } from '../blocks/payload-part';
import { BlockUi } from './block-ui';
import { mountReactBlockHost, type ReactBlockViewProps } from './host';
import type { ReactBlockProviders } from './providers';

const Locale = createContext('missing-locale');
const Services = createContext({ format: (value: string) => `missing:${value}` });
const environment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean };
function Widget({ block, children }: ReactBlockViewProps) {
  const locale = useContext(Locale),
    services = useContext(Services),
    [count, setCount] = useState(0);
  return (
    <div style={{ display: 'contents' }}>
      <BlockUi>
        <output>
          {locale}|{services.format(String(block.payload.attributes.name))}
        </output>
        <button type="button" onClick={() => setCount((old) => old + 1)}>
          Count {count}
        </button>
      </BlockUi>
      {children}
    </div>
  );
}
const providers =
  (locale: string, prefix: string): ReactBlockProviders =>
  (children) => (
    <Locale.Provider value={locale}>
      <Services.Provider value={{ format: (value) => `${prefix}:${value}` }}>
        {children}
      </Services.Provider>
    </Locale.Provider>
  );

for (const receiver of [0, 1] as const)
  test(`explicit providers update independently without remounting native containers or losing shared edits; receiver ${receiver}`, async () => {
    const previous = environment.IS_REACT_ACT_ENVIRONMENT;
    environment.IS_REACT_ACT_ENVIRONMENT = true;
    const pair = await createTwoEditorFixture({
      document: reactBlockDocument(),
      clientIds: [73, 41],
    });
    const hosts: ReturnType<typeof mountReactBlockHost>[] = [];
    const [alice, bob] = pair.editors;
    try {
      await act(async () => {
        for (const [peer, editor] of pair.editors.entries())
          hosts.push(
            mountReactBlockHost(
              editor,
              new Map([
                ['test_counter', { view: Widget, nativeContent: true }],
                ['future_widget', { view: Widget, nativeContent: true }],
              ]),
              {
                providers: (children, scope) => {
                  expect(scope.editor === editor).toBe(true);
                  expect(['counter-1', 'future-1']).toContain(scope.block.id);
                  return providers(peer === 0 ? 'en' : 'cs', peer === 0 ? 'Alice' : 'Bob')(
                    children,
                    scope
                  );
                },
              }
            )
          );
      });
      const buttons = pair.containers.map((node) => node.querySelector('button')!);
      await act(async () => {
        buttons[0]!.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      });
      expect(buttons.map((button) => button.textContent)).toEqual(['Count 1', 'Count 0']);
      expect(pair.containers.map((node) => node.querySelector('output')?.textContent)).toEqual([
        'en|Alice:Initial',
        'cs|Bob:Initial',
      ]);
      expect(
        pair.containers.map((node) => node.querySelectorAll('[data-sciobot-block-host]').length)
      ).toEqual([2, 2]);
      const before = alice!.surface!.session.packageRevision();
      await act(async () => {
        hosts[0]!.setProviders(providers('fr', 'Alice updated'));
      });
      expect(alice!.surface!.session.packageRevision()).toBe(before);
      expect(pair.containers.map((node) => node.querySelector('output')?.textContent)).toEqual([
        'fr|Alice updated:Initial',
        'cs|Bob:Initial',
      ]);
      expect(pair.containers[0]!.querySelector('button') === buttons[0]).toBe(true);
      expect(
        pair.containers.map((node) =>
          [...node.querySelectorAll('output')].map((output) => output.textContent)
        )
      ).toEqual([
        ['fr|Alice updated:Initial', 'fr|Alice updated:Initial'],
        ['cs|Bob:Initial', 'cs|Bob:Initial'],
      ]);
      expect(buttons[0]!.textContent).toBe('Count 1');
      pair.pause();
      await act(async () => {
        const block = buildBlockIndex(alice!.surface!.session.currentPackage()).byId.get(
          'counter-1'
        )!;
        expect(
          writeBlockPayload(alice!, {
            controlId: block.controlId,
            payload: { ...block.payload, attributes: { name: 'Shared' } },
            label: 'Shared',
            fallbackOps: [],
          }).success
        ).toBe(true);
        const session = bob!.surface!.session;
        const paragraphId = session
          .paragraphIds()
          .find((id) => paragraphTextOf(session.part(), id) === 'Editable body')!;
        bob!.surface!.setSelection({
          anchor: { paragraphId, offset: 13 },
          head: { paragraphId, offset: 13 },
        });
        bob!.surface!.insertPlainText(' Bob');
        bob!.surface!.flushPendingInput();
        pair.reconnect(receiver);
        await pair.settle();
      });
      expect(pair.containers.map((node) => node.querySelector('output')?.textContent)).toEqual([
        'fr|Alice updated:Shared',
        'cs|Bob:Shared',
      ]);
      expect(
        pair.containers.every((node) =>
          node
            .querySelector('[data-sciobot-native-slot]')
            ?.textContent?.includes('Editable body Bob')
        )
      ).toBe(true);
      await act(async () => {
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
      });
      expect(pair.containers.map((node) => node.querySelector('output')?.textContent)).toEqual([
        'fr|Alice updated:Initial',
        'cs|Bob:Initial',
      ]);
      await act(async () => {
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
      });
      expect(pair.containers.map((node, i) => node.querySelector('button') === buttons[i])).toEqual(
        [true, true]
      );
      expect(buttons.map((button) => button.textContent)).toEqual(['Count 1', 'Count 0']);
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      pair.pause();
      await act(async () => {
        const session = bob!.surface!.session;
        const paragraphId = session
          .paragraphIds()
          .find((id) => paragraphTextOf(session.part(), id) === 'Editable body Bob')!;
        bob!.surface!.setSelection({
          anchor: { paragraphId, offset: 17 },
          head: { paragraphId, offset: 17 },
        });
        bob!.surface!.insertPlainText(' Editing');
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
      expect(
        pair.containers.map((node) => node.querySelectorAll('[data-sciobot-block-host]').length)
      ).toEqual([1, 1]);
      expect(pair.containers.map((node) => node.querySelector('output')?.textContent)).toEqual([
        'fr|Alice updated:Initial',
        'cs|Bob:Initial',
      ]);
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
