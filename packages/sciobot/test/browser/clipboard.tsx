// Test-only collaboration transport under the EigenPal Pro License.
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import {
  mountBlockClipboard,
  mountReactBlockHost,
  type ReactBlockViewProps,
} from '@sciobot/docx/react';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { createTwoEditorFixture } from '../support/collaboration';
import { blockClipboardDocument } from '../support/react-block-document';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';

const fonts = await loadDefaultFonts({ families: ['Calibri'] });
if (fonts.failures.length) throw new Error('Packaged browser fonts failed.');
const pair = await createTwoEditorFixture({
  document: blockClipboardDocument(),
  clientIds: [73, 41],
  editor: { fonts },
});
function Container({ block, children }: ReactBlockViewProps) {
  return (
    <div
      data-block={block.id}
      data-type={block.payload.type}
      data-attributes={JSON.stringify(block.payload.attributes)}
      style={{ display: 'contents' }}
    >
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
const registry = new Map(
  ['phase', 'activity', 'note'].map((type) => [type, { view: Container, nativeContent: true }])
);
const hosts = pair.editors.map((editor) => mountReactBlockHost(editor, registry));
const clipboards = pair.editors.map((editor) => mountBlockClipboard(editor));
const transfers: Record<string, string>[] = [];
document.addEventListener('copy', (event) => {
  const data: Record<string, string> = {};
  for (const type of event.clipboardData?.types ?? [])
    data[type] = event.clipboardData!.getData(type);
  transfers.push(data);
});
document.addEventListener('cut', (event) => {
  const data: Record<string, string> = {};
  for (const type of event.clipboardData?.types ?? [])
    data[type] = event.clipboardData!.getData(type);
  transfers.push(data);
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
  transfers,
  idOf,
  /** Blocks in document order with parent, payload, and child paragraph texts. */
  blocks(peer = 0) {
    const current = surface(peer);
    current.flushPendingInput();
    const index = buildBlockIndex(current.session.currentPackage());
    return {
      diagnostics: index.diagnostics.map((entry) => entry.code + ':' + (entry.id ?? '')),
      blocks: [...index.byId.values()].map((block) => ({
        id: block.id,
        parentId: block.parentId,
        type: block.payload.type,
        attributes: block.payload.attributes,
      })),
    };
  },
  texts(peer = 0) {
    const current = surface(peer);
    current.flushPendingInput();
    return current.session
      .paragraphIds()
      .map((id) => paragraphTextOf(current.session.part(), id) ?? '');
  },
  revision(peer = 0) {
    surface(peer).flushPendingInput();
    return surface(peer).session.packageRevision();
  },
  async select(anchor: [string, number], head: [string, number], peer = 0) {
    surface(peer).setSelection({
      anchor: { paragraphId: idOf(anchor[0], peer), offset: anchor[1] },
      head: { paragraphId: idOf(head[0], peer), offset: head[1] },
    });
    await pair.settle();
  },
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
  async removeBlock(id: string, peer = 0) {
    const current = surface(peer),
      block = buildBlockIndex(current.session.currentPackage()).byId.get(id)!;
    const result = current.applyAutomationOps(() => [
      { op: 'removeContentControl', controlId: block.controlId, keepContent: false },
    ]);
    await pair.settle();
    return !result.rejected;
  },
  pause: () => pair.pause(),
  async reconnect(first: 0 | 1 = 0) {
    pair.reconnect(first);
    await pair.settle();
  },
  async settle() {
    await pair.settle();
  },
  async exec(type: 'undo' | 'redo', peer = 0) {
    const result = pair.editors[peer]!.exec({ type });
    await pair.settle();
    return result.ok;
  },
  savedDigests: () => pair.savedDigests(),
  async reopenedBlocks(peer = 0) {
    const { readOoxmlPackage } = await import('@docx-editor.dev/core/store');
    const read = readOoxmlPackage(new Uint8Array(await pair.editors[peer]!.save()));
    if (!read.ok) throw new Error(read.reason);
    const index = buildBlockIndex(read.package);
    return {
      diagnostics: index.diagnostics.map((entry) => entry.code),
      ids: [...index.byId.keys()],
    };
  },
};
declare global {
  interface Window {
    blockClipboard: typeof api;
  }
}
window.blockClipboard = api;
window.addEventListener(
  'pagehide',
  () => {
    clipboards.forEach((clipboard) => clipboard.dispose());
    hosts.forEach((host) => host.dispose());
    pair.destroy();
  },
  { once: true }
);
