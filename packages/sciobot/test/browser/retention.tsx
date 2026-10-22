// Test-only transport uses the EigenPal Pro License.
import { useEffect, useState } from 'react';
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  paragraphTextOf,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { BlockUi, mountInlineInput, mountReactBlockHost } from '@sciobot/docx/react';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { removeBlock } from '../../src/blocks/remove';
import { createTwoEditorFixture } from '../support/collaboration';
import { reactBlockDocument, hostParagraph } from '../support/react-block-document';
import { bodyPart } from '../support/document';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';

// Count only live observations. Disconnected observers never keep a test-side strong reference.
const observations = new Map<ResizeObserver, Set<Element>>();
const NativeObserver = window.ResizeObserver;
window.ResizeObserver = class extends NativeObserver {
  observe(target: Element, options?: ResizeObserverOptions) {
    const targets = observations.get(this) ?? new Set<Element>();
    targets.add(target);
    observations.set(this, targets);
    super.observe(target, options);
  }
  unobserve(target: Element) {
    const targets = observations.get(this);
    targets?.delete(target);
    if (!targets?.size) observations.delete(this);
    super.unobserve(target);
  }
  disconnect() {
    observations.delete(this);
    super.disconnect();
  }
};
const listeners = new Map<EventTarget, Set<string>>();
const add = EventTarget.prototype.addEventListener,
  remove = EventTarget.prototype.removeEventListener;
EventTarget.prototype.addEventListener = function (type, callback, options) {
  if (typeof callback === 'function' && callback.name === 'stopWidgetEvent') {
    const events = listeners.get(this) ?? new Set<string>();
    events.add(type);
    listeners.set(this, events);
  }
  if (callback) return add.call(this, type, callback, options);
};
EventTarget.prototype.removeEventListener = function (type, callback, options) {
  if (typeof callback === 'function' && callback.name === 'stopWidgetEvent') {
    const events = listeners.get(this);
    events?.delete(type);
    if (!events?.size) listeners.delete(this);
  }
  if (callback) return remove.call(this, type, callback, options);
};
const fonts = await loadDefaultFonts({ families: ['Calibri'] });
if (fonts.failures.length) throw new Error('Packaged fonts failed.');
const opened = readOoxmlPackage(reactBlockDocument());
if (!opened.ok) throw new Error(opened.reason);
const original = opened.package.parts.get(opened.package.mainDocumentPart)!;
const body = original.root.children.find((node) => node.kind === 'body')!;
if (body.kind === 'textValue') throw new Error('Missing body');
const read = bodyPart(
  '<w:p/>' +
    Array.from(
      { length: 18 },
      (_, i) => `<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Page ${i}</w:t></w:r></w:p>`
    ).join('') +
    hostParagraph('End')
);
const tail = read.root.children.find((node) => node.kind === 'body')!;
if (tail.kind === 'textValue') throw new Error('Missing tail');
// Parse the complete fixture to allocate distinct canonical IDs.
import { serializeOoxmlPart, readOoxmlPart } from '@docx-editor.dev/core/store';
const longXml = serializeOoxmlPart(original).replace(
  '</w:body>',
  serializeOoxmlPart(read).match(/<w:body>(.*)<\/w:body>/s)![1] + '</w:body>'
);
const long = readOoxmlPart(longXml, { name: original.name, contentType: original.contentType });
if (!long.ok) throw new Error(long.reason);
const bytes = writeOoxmlPackage(withPart(opened.package, long.part));
const events = { mounted: 0, unmounted: 0 };
function Widget() {
  const [count, setCount] = useState(0);
  useEffect(() => {
    events.mounted++;
    return () => {
      events.unmounted++;
    };
  }, []);
  return (
    <BlockUi style={{ background: 'white', width: 160, height: 55 }}>
      <button onClick={() => setCount((old) => old + 1)}>Count {count}</button>
      <input aria-label="Retained note" defaultValue="" />
    </BlockUi>
  );
}
let pair: Awaited<ReturnType<typeof createTwoEditorFixture>>;
const panes: HTMLElement[] = [];
async function settle() {
  await pair.settle();
  await new Promise<void>((resolve) =>
    requestAnimationFrame(() => requestAnimationFrame(() => resolve()))
  );
}
async function start() {
  pair = await createTwoEditorFixture({ document: bytes, editor: { fonts }, clientIds: [73, 41] });
  for (const [peer, editor] of pair.editors.entries()) {
    const pane = document.createElement('section');
    pane.className = 'editor-pane docx-editor__scroll-container';
    pane.dataset.pane = String(peer);
    pane.append(pair.containers[peer]!);
    document.body.append(pane);
    panes.push(pane);
    mountReactBlockHost(editor, new Map([['test_counter', { view: Widget, reserveHeight: true }]]));
    mountInlineInput(editor);
  }
  await settle();
}
const stats = () => ({
  roots: events.mounted - events.unmounted,
  observers: observations.size,
  observations: [...observations.values()].reduce((sum, set) => sum + set.size, 0),
  listenerRoots: listeners.size,
});
async function stop() {
  pair.destroy();
  panes.splice(0).forEach((pane) => pane.remove());
  await new Promise<void>((resolve) =>
    requestAnimationFrame(() => requestAnimationFrame(() => resolve()))
  );
  return stats();
}
await start();
const api = {
  events,
  stats,
  settle,
  parked: (peer = 0) => !panes[peer]!.querySelector('[data-sciobot-block-host]'),
  async scroll(bottom: boolean, peer = 0) {
    const surface = pair.editors[peer]!.surface!;
    const ids = surface.session.paragraphIds();
    const position = { paragraphId: bottom ? ids.at(-1)! : ids[0]!, offset: 0 };
    surface.setSelection({ anchor: position, head: position });
    panes[peer]!.scrollTo(0, bottom ? panes[peer]!.scrollHeight : 0);
    await settle();
  },
  async cycle() {
    const empty = await stop();
    await start();
    return empty;
  },
  stop,
  pause: () => pair.pause(),
  async reconnect(receiver: 0 | 1) {
    pair.reconnect(receiver);
    await settle();
  },
  async insert(text: string) {
    const current = pair.editors[1]!.surface!,
      paragraphId = current.session
        .paragraphIds()
        .find((id) => paragraphTextOf(current.session.part(), id)?.startsWith('Editable body'))!;
    const point = {
      paragraphId,
      offset: paragraphTextOf(current.session.part(), paragraphId)!.length,
    };
    current.setSelection({ anchor: point, head: point });
    current.insertPlainText(text);
    current.flushPendingInput();
    await settle();
  },
  text: (peer = 0) => pair.editors[peer]!.surface!.session.bodyText(),
  async remove() {
    const editor = pair.editors[0]!,
      controlId = buildBlockIndex(editor.surface!.session.currentPackage()).byId.get(
        'counter-1'
      )!.controlId;
    const result = removeBlock(editor, { controlId, mode: 'delete' });
    await settle();
    return result.success;
  },
  async undo() {
    const result = pair.editors[0]!.exec({ type: 'undo' }).ok;
    await settle();
    return result;
  },
  async redo() {
    const result = pair.editors[0]!.exec({ type: 'redo' }).ok;
    await settle();
    return result;
  },
  saved: () => pair.savedDigests(),
  remote() {
    const editor = pair.editors[0]!,
      selection = pair.rooms[0]!.session.remoteSelections().find(
        (selection) => selection.actorId === 'bob'
      );
    return selection
      ? caretAt(editor.surface!.layout(), {
          paragraphId: selection.head.nodeId,
          offset: selection.head.offset,
        })
      : null;
  },
};
declare global {
  interface Window {
    blockRetention: typeof api;
  }
}
window.blockRetention = api;
window.addEventListener('pagehide', () => pair.destroy(), { once: true });
