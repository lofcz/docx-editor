// Test-only collaboration transport under the EigenPal Pro License.
import { useEffect, useState } from 'react';
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import {
  BlockUi,
  mountReactBlockHost,
  type ReactBlockViewProps,
  type ReactBlockView,
  type ReactBlockRegistration,
} from '@sciobot/docx/react';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { createTwoEditorFixture } from '../support/collaboration';
import { blockSelectionDocument } from '../support/react-block-document';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';

const fonts = await loadDefaultFonts({ families: ['Calibri'] });
if (fonts.failures.length) throw new Error('Packaged browser fonts failed.');
const pair = await createTwoEditorFixture({
  document: blockSelectionDocument(),
  clientIds: [73, 41],
  editor: { fonts },
});
const events = [
  { mounted: 0, unmounted: 0 },
  { mounted: 0, unmounted: 0 },
];
function Container({ block, fragment, children }: ReactBlockViewProps) {
  return (
    <div
      data-container={block.id}
      data-native-page={fragment.pageIndex}
      style={{ display: 'contents' }}
    >
      {children}
    </div>
  );
}
function formView(peer: number) {
  return function Form({ fragment }: ReactBlockViewProps) {
    const [notes, setNotes] = useState('');
    const [count, setCount] = useState(0);
    useEffect(() => {
      events[peer]!.mounted++;
      return () => {
        events[peer]!.unmounted++;
      };
    }, []);
    return (
      <BlockUi data-form-page={fragment.pageIndex} style={{ width: 180, minHeight: 80 }}>
        <button type="button" onClick={() => setCount((value) => value + 1)}>
          Count {count}
        </button>
        <input
          aria-label="Local notes"
          value={notes}
          onChange={(event) => setNotes(event.currentTarget.value)}
        />
      </BlockUi>
    );
  };
}
for (const [index, container] of pair.containers.entries()) {
  const pane = document.createElement('section');
  pane.dataset.pane = String(index);
  pane.className = 'editor-pane';
  document.body.append(pane);
  pane.append(container);
}
const hosts = pair.editors.map((editor, peer) =>
  mountReactBlockHost(
    editor,
    new Map<string, ReactBlockView | ReactBlockRegistration>([
      ['native_container', { view: Container, nativeContent: true }],
      ['test_form', { view: formView(peer), reserveHeight: true }],
    ])
  )
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
const api = {
  events,
  async push(peer = 0) {
    const current = surface(peer);
    const result = current.applyAutomationOps(() => [
      {
        op: 'setParagraphProperties',
        paragraphId: idOf('Intro', peer),
        properties: [{ localName: 'spacing', attributes: { after: '9000' } }],
      },
    ]);
    await pair.settle();
    return !result.rejected;
  },
  settle: () => pair.settle(),
  async removeBlock(id: string, peer = 0) {
    const current = surface(peer),
      block = buildBlockIndex(current.session.currentPackage()).byId.get(id)!;
    const result = current.applyAutomationOps(() => [
      { op: 'removeContentControl', controlId: block.controlId, keepContent: false },
    ]);
    await pair.settle();
    return !result.rejected;
  },
  remoteGeometry() {
    const remote = pair.rooms[0]!.session.remoteSelections().find(
      (selection) => selection.actorId === 'bob'
    );
    if (!remote) return null;
    return pair.editors.map((editor) =>
      [remote.anchor, remote.head].map((position) => {
        const caret = caretAt(editor.surface!.layout(), {
          paragraphId: position.nodeId,
          offset: position.offset,
        });
        return caret && { x: caret.x, y: caret.y, page: caret.pageIndex };
      })
    );
  },
  async select(anchor: [string, number], head: [string, number], peer = 1) {
    surface(peer).setSelection({
      anchor: { paragraphId: idOf(anchor[0], peer), offset: anchor[1] },
      head: { paragraphId: idOf(head[0], peer), offset: head[1] },
    });
    await pair.settle();
  },
  /** Edit without moving the shared page selection, as a separate browser would. */
  async insert(text: string, paragraph: string, peer = 1) {
    const current = surface(peer),
      paragraphId = idOf(paragraph, peer);
    const offset = paragraphTextOf(current.session.part(), paragraphId)!.length;
    const result = current.applyAutomationOps(() => [
      { op: 'insertText', paragraphId, offset, text },
    ]);
    await pair.settle();
    return !result.rejected;
  },
  pause: () => pair.pause(),
  async reconnect(first: 0 | 1 = 0) {
    pair.reconnect(first);
    await pair.settle();
  },
  async exec(type: 'undo' | 'redo', peer = 0) {
    const result = pair.editors[peer]!.exec({ type });
    await pair.settle();
    return result.ok;
  },
  savedDigests: () => pair.savedDigests(),
};
declare global {
  interface Window {
    blockFragments: typeof api;
  }
}
window.blockFragments = api;
window.addEventListener(
  'pagehide',
  () => {
    hosts.forEach((host) => host.dispose());
    pair.destroy();
  },
  { once: true }
);
