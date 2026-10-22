// Test-only collaboration transport under the EigenPal Pro License.
import { useState } from 'react';
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import {
  BlockUi,
  mountReactBlockHost,
  readBlockSelection,
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
function Container({ block, children }: ReactBlockViewProps) {
  return (
    <div data-container={block.id} style={{ display: 'contents' }}>
      {children}
    </div>
  );
}
function Form({ fragment }: ReactBlockViewProps) {
  const [notes, setNotes] = useState('Widget words');
  return (
    <BlockUi
      className="widget-ui"
      style={{ left: fragment.box.width * fragment.scale - 180, right: 'auto' }}
    >
      <textarea
        aria-label="Widget notes"
        value={notes}
        onChange={(event) => setNotes(event.currentTarget.value)}
      />
    </BlockUi>
  );
}
for (const [index, container] of pair.containers.entries()) {
  const pane = document.createElement('section');
  pane.dataset.pane = String(index);
  pane.className = 'editor-pane';
  document.body.append(pane);
  pane.append(container);
}
const hosts = pair.editors.map((editor) =>
  mountReactBlockHost(
    editor,
    new Map<string, ReactBlockView | ReactBlockRegistration>([
      ['native_container', { view: Container, nativeContent: true }],
      ['test_form', Form],
    ])
  )
);
const copies: string[] = [];
document.addEventListener('copy', (event) => {
  copies.push(event.clipboardData?.getData('text/plain') ?? '');
});
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
  copies,
  idOf,
  model(peer = 0) {
    surface(peer).flushPendingInput();
    return surface(peer).state().selection;
  },
  blocks(peer = 0) {
    const result = readBlockSelection(pair.editors[peer]!);
    return result && { ...result, blocks: result.blocks.map((entry) => ({ ...entry })) };
  },
  pageOf(paragraphId: string, offset: number, peer = 0) {
    return caretAt(surface(peer).layout(), { paragraphId, offset })?.pageIndex ?? null;
  },
  /** Canonical oracle for the selected range: paragraph text joined by line feeds. */
  rangeText(peer = 0) {
    const current = surface(peer);
    current.flushPendingInput();
    const { anchor, head } = current.state().selection;
    const ids = current.session.paragraphIds();
    const a = ids.indexOf(anchor.paragraphId),
      h = ids.indexOf(head.paragraphId);
    const forward = a < h || (a === h && anchor.offset <= head.offset);
    const [start, end] = forward ? [anchor, head] : [head, anchor];
    const from = ids.indexOf(start.paragraphId),
      to = ids.indexOf(end.paragraphId);
    const part = current.session.part();
    return ids
      .slice(from, to + 1)
      .map((id, index) => {
        const text = paragraphTextOf(part, id) ?? '';
        return text.slice(
          index === 0 ? start.offset : 0,
          from + index === to ? end.offset : undefined
        );
      })
      .join('\n');
  },
  textAt(position: { paragraphId: string }, peer = 0) {
    return paragraphTextOf(surface(peer).session.part(), position.paragraphId);
  },
  revision(peer = 0) {
    surface(peer).flushPendingInput();
    return surface(peer).session.packageRevision();
  },
  async removeBlock(id: string, peer = 0) {
    const current = surface(peer),
      block = buildBlockIndex(current.session.currentPackage()).byId.get(id)!;
    const result = current.applyAutomationOps(() => [
      { op: 'removeContentControl', controlId: block.controlId, keepContent: false },
    ]);
    await pair.settle();
    return !result.rejected;
  },
  text(paragraph: string, peer = 0) {
    return paragraphTextOf(surface(peer).session.part(), idOf(paragraph, peer));
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
  async append(text: string, paragraph: string, peer = 1) {
    const current = surface(peer),
      paragraphId = idOf(paragraph, peer);
    const offset = paragraphTextOf(current.session.part(), paragraphId)!.length;
    current.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
    current.insertPlainText(text);
    current.flushPendingInput();
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
    blockSelection: typeof api;
  }
}
window.blockSelection = api;
window.addEventListener(
  'pagehide',
  () => {
    hosts.forEach((host) => host.dispose());
    pair.destroy();
  },
  { once: true }
);
