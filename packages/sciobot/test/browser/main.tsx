// Test-only browser transport evaluation under the EigenPal Pro License.
import { createContext, useContext, useEffect, useState } from 'react';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import {
  BlockUi,
  mountReactBlockHost,
  type ReactBlockViewProps,
  type ReactBlockProviders,
} from '@sciobot/docx/react';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { writeBlockPayload } from '../../src/blocks/payload-part';
import { createTwoEditorFixture } from '../support/collaboration';
import { hostParagraph, reactBlockDocument } from '../support/react-block-document';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';

const fonts = await loadDefaultFonts({ families: ['Calibri'] });
if (fonts.failures.length) throw new Error('Packaged browser fonts failed.');
const prefix = Array.from({ length: 12 }, (_, index) => hostParagraph(`Prelude ${index + 1}`)).join(
  ''
);
const pair = await createTwoEditorFixture({
  document: reactBlockDocument(prefix),
  clientIds: [73, 41],
  editor: { fonts },
});
const events = [
  { mounted: 0, unmounted: 0 },
  { mounted: 0, unmounted: 0 },
  { mounted: 0, unmounted: 0 },
];
const Locale = createContext('missing');
const Services = createContext({ format: (value: string) => `missing:${value}` });
function providers(locale: string, prefix: string): ReactBlockProviders {
  const services = { format: (value: string) => `${prefix}:${value}` };
  return (children) => (
    <Locale.Provider value={locale}>
      <Services.Provider value={services}>{children}</Services.Provider>
    </Locale.Provider>
  );
}
const providerScopes = [
  providers('en', 'Alice'),
  providers('cs', 'Bob'),
  providers('en', 'Reopened'),
];
function widget(peer: number) {
  return function Widget({ block, paragraphs, fragment }: ReactBlockViewProps) {
    const [count, setCount] = useState(0),
      [notes, setNotes] = useState('');
    const locale = useContext(Locale),
      services = useContext(Services);
    useEffect(() => {
      events[peer]!.mounted++;
      return () => {
        events[peer]!.unmounted++;
      };
    }, []);
    return (
      <BlockUi className="widget-ui" data-testid={`widget-${peer}`}>
        <button type="button" onClick={() => setCount((old) => old + 1)}>
          Count {count}
        </button>
        <strong data-name="">{String(block.payload.attributes.name)}</strong>
        <output data-provider="">
          {locale}|{services.format(String(block.payload.attributes.name))}
        </output>
        <label>
          Notes {peer}
          <input value={notes} onChange={(event) => setNotes(event.currentTarget.value)} />
        </label>
        <small data-native-text="">{paragraphs.map((p) => p.text).join('|')}</small>
        <span data-page="">Page {fragment.pageIndex}</span>
      </BlockUi>
    );
  };
}
const views = [widget(0), widget(1), widget(2)];
for (const [index, container] of pair.containers.entries()) {
  const pane = document.createElement('section');
  pane.dataset.pane = String(index);
  pane.className = 'editor-pane';
  const label = document.createElement('h2');
  label.textContent = index === 0 ? 'Alice' : 'Bob';
  document.body.append(label, pane);
  pane.append(container);
}
const hosts = pair.editors.map((editor, peer) =>
  mountReactBlockHost(editor, new Map(), { providers: providerScopes[peer] })
);
const registrations = hosts.map((host, index) => host.register('test_counter', views[index]!));
const surface = (peer = 0) => pair.editors[peer]!.surface!;
const nativeId = (text: string, peer = 0) => {
  const current = surface(peer);
  const id = current.session
    .paragraphIds()
    .find((id) => paragraphTextOf(current.session.part(), id) === text);
  if (!id) throw new Error(`Missing native text: ${text}`);
  return id;
};
const rename = (name: string, peer = 0) => {
  const editor = pair.editors[peer]!;
  const block = buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('counter-1')!;
  return writeBlockPayload(editor, {
    controlId: block.controlId,
    payload: { ...block.payload, attributes: { name } },
    label: name,
    fallbackOps: [],
  }).success;
};
let reopened: ReturnType<typeof createDocxEditor> | undefined;
let reopenedHost: ReturnType<typeof mountReactBlockHost> | undefined;
const api = {
  events,
  rename,
  providers(locale: string, prefix: string, peer = 0) {
    providerScopes[peer] = providers(locale, prefix);
    hosts[peer]!.setProviders(providerScopes[peer]);
  },
  selection(peer = 0) {
    const current = surface(peer),
      selection = current.hostElement().ownerDocument.defaultView!.getSelection();
    return {
      model: current.state().selection,
      text: selection?.anchorNode?.textContent ?? null,
      offset: selection?.anchorOffset ?? null,
    };
  },
  pause: () => pair.pause(),
  async reconnect(first: 0 | 1 = 0) {
    pair.reconnect(first);
    await pair.settle();
  },
  async settle() {
    await pair.settle();
  },
  async append(text: string, peer = 1) {
    const current = surface(peer),
      id = nativeId('Editable body', peer);
    current.setSelection({
      anchor: { paragraphId: id, offset: 13 },
      head: { paragraphId: id, offset: 13 },
    });
    current.insertPlainText(text);
    current.flushPendingInput();
    await pair.settle();
  },
  async undo() {
    const result = pair.editors[0]!.exec({ type: 'undo' });
    await pair.settle();
    return result.ok;
  },
  async redo() {
    const result = pair.editors[0]!.exec({ type: 'redo' });
    await pair.settle();
    return result.ok;
  },
  async push() {
    const current = surface();
    const result = current.applyAutomationOps(() => [
      {
        op: 'setParagraphProperties',
        paragraphId: nativeId('Intro'),
        properties: [{ localName: 'spacing', attributes: { after: '9000' } }],
      },
    ]);
    await pair.settle();
    return !result.rejected;
  },
  async remove() {
    const current = surface();
    const block = buildBlockIndex(current.session.currentPackage()).byId.get('counter-1')!;
    const result = current.applyAutomationOps(() => [
      { op: 'removeContentControl', controlId: block.controlId, keepContent: false },
    ]);
    await pair.settle();
    return !result.rejected;
  },
  unregister(peer = 0) {
    registrations[peer]!();
  },
  register(peer = 0) {
    registrations[peer] = hosts[peer]!.register('test_counter', views[peer]!);
  },
  dispose(peer = 0) {
    hosts[peer]!.dispose();
  },
  async savedDigests() {
    return pair.savedDigests();
  },
  async reopen() {
    const bytes = await pair.editors[0]!.save();
    const container = document.createElement('div');
    container.dataset.reopened = '';
    document.body.append(container);
    reopened = createDocxEditor({ container, document: new Uint8Array(bytes), fonts });
    reopenedHost = mountReactBlockHost(reopened, new Map([['test_counter', views[2]!]]), {
      providers: providerScopes[2],
    });
  },
};
declare global {
  interface Window {
    blockHostProof: typeof api;
  }
}
window.blockHostProof = api;
window.addEventListener(
  'pagehide',
  () => {
    hosts.forEach((host) => host.dispose());
    reopenedHost?.dispose();
    pair.destroy();
    reopened?.destroy();
  },
  { once: true }
);
