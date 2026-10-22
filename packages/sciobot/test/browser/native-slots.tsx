// Test-only collaboration transport under the EigenPal Pro License.
import { useState } from 'react';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { BlockUi, mountReactBlockHost, type ReactBlockViewProps } from '@sciobot/docx/react';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { createTwoEditorFixture } from '../support/collaboration';
import { nativeSlotDocument } from '../support/react-block-document';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';

const fonts = await loadDefaultFonts({ families: ['Calibri'] });
if (fonts.failures.length) throw new Error('Packaged browser fonts failed.');
const pair = await createTwoEditorFixture({
  document: nativeSlotDocument(),
  clientIds: [73, 41],
  editor: { fonts },
});
function Container({ block, children, fragment }: ReactBlockViewProps) {
  const [count, setCount] = useState(0);
  return (
    <div data-container={block.id} style={{ display: 'contents' }}>
      {block.id === 'outer' ? (
        <BlockUi
          style={{
            position: 'absolute',
            left: (fragment.box.x + fragment.box.width - 90) * fragment.scale,
            top: fragment.box.y * fragment.scale,
            width: 85,
            zIndex: 1,
          }}
        >
          <button type="button" onClick={() => setCount((old) => old + 1)}>
            Count {count}
          </button>
        </BlockUi>
      ) : null}
      {children}
    </div>
  );
}
for (const [index, container] of pair.containers.entries()) {
  const pane = document.createElement('section');
  pane.dataset.pane = String(index);
  pane.className = 'editor-pane';
  document.body.append(pane);
  pane.append(container);
}
const hosts = pair.editors.map((editor) => mountReactBlockHost(editor));
const stops = hosts.map((host) =>
  host.register('native_container', Container, { nativeContent: true })
);
const surface = (peer = 0) => pair.editors[peer]!.surface!;
const idOf = (text: string, peer = 0) => {
  const current = surface(peer);
  const id = current.session
    .paragraphIds()
    .find((id) => paragraphTextOf(current.session.part(), id) === text);
  if (!id) throw new Error(`Missing paragraph: ${text}`);
  return id;
};
const nativeIds = Object.fromEntries(
  ['Inner text', 'Cell text', 'Other cell', 'Row cell text'].map((text) => [text, idOf(text)])
);
let reopened: ReturnType<typeof createDocxEditor> | undefined;
let reopenedHost: ReturnType<typeof mountReactBlockHost> | undefined;
const api = {
  async append(text: string, peer = 1, paragraph = 'Other cell') {
    const current = surface(peer),
      paragraphId = nativeIds[paragraph]!;
    const offset = paragraphTextOf(current.session.part(), paragraphId)!.length;
    current.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
    current.insertPlainText(text);
    current.flushPendingInput();
    await pair.settle();
  },
  selection(peer = 0) {
    const current = surface(peer),
      dom = current.hostElement().ownerDocument.defaultView!.getSelection()!;
    const node = dom.anchorNode;
    return {
      model: current.state().selection,
      dom: {
        paragraphId:
          node?.parentElement?.closest<HTMLElement>('.docx-paragraph-fragment')?.dataset
            .paragraphId ?? null,
        text: node?.textContent ?? null,
        offset: dom.anchorOffset,
      },
    };
  },
  remoteGeometry() {
    const remote = pair.rooms[0]!.session.remoteSelections().find(
      (selection) => selection.actorId === 'bob'
    );
    if (!remote) return null;
    return pair.editors.map((editor) => {
      const caret = caretAt(editor.surface!.layout(), {
        paragraphId: remote.head.nodeId,
        offset: remote.head.offset,
      });
      return caret ? { x: caret.x, y: caret.y } : null;
    });
  },
  refresh(peer = 0) {
    hosts[peer]!.refresh();
  },
  unregister(peer = 0) {
    stops[peer]!();
  },
  register(peer = 0) {
    stops[peer] = hosts[peer]!.register('native_container', Container, { nativeContent: true });
  },
  dispose(peer = 0) {
    hosts[peer]!.dispose();
  },
  pause() {
    pair.pause();
  },
  async reconnect(first: 0 | 1 = 0) {
    pair.reconnect(first);
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
  async remove() {
    const current = surface(),
      outer = buildBlockIndex(current.session.currentPackage()).byId.get('outer')!;
    const result = current.applyAutomationOps(() => [
      { op: 'removeContentControl', controlId: outer.controlId, keepContent: false },
    ]);
    await pair.settle();
    return !result.rejected;
  },
  savedDigests: () => pair.savedDigests(),
  async reopen() {
    const bytes = await pair.editors[0]!.save(),
      container = document.createElement('div');
    container.dataset.reopened = '';
    document.body.append(container);
    reopened = createDocxEditor({ container, document: new Uint8Array(bytes), fonts });
    reopenedHost = mountReactBlockHost(
      reopened,
      new Map([['native_container', { view: Container, nativeContent: true }]])
    );
  },
};
declare global {
  interface Window {
    nativeSlotProof: typeof api;
  }
}
window.nativeSlotProof = api;
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
