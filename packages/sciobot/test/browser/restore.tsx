// Test-only native transport under the EigenPal Pro License.
import { useSyncExternalStore } from 'react';
import { createRoot } from 'react-dom/client';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { mountReactBlockHost, type ReactBlockViewProps } from '@sciobot/docx/react';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { readEditorRevision } from '../../src/projection/revision-cache';
import { acceptReviewGroup } from '../../src/review/accept-group';
import { rejectReviewGroup } from '../../src/review/reject-group';
import { acceptAllReviewGroups } from '../../src/review/accept-all';
import { rejectAllReviewGroups } from '../../src/review/reject-all';
import { restoreReviewState, type RestoreReviewResult } from '../../src/review/restore';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../../src/review/native-module';
import { createTwoEditorFixture } from '../support/collaboration';
import { groupFixture } from '../support/review-groups';
import { bulkSource, proposeMixedBulk } from '../support/bulk-decisions';
import { propertyContext } from '../support/property-execution';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';

const STORAGE = 'sciobot-review-docx-v1';
const query = new URLSearchParams(globalThis.location.search);
const mode = query.get('storage') === 'room' ? 'room' : 'docx';
const reverse = query.get('reverse') === '1';
const context = { documentId: propertyContext.documentId, epoch: crypto.randomUUID() };
function decode(text: string) {
  return Uint8Array.from(atob(text), (character) => character.charCodeAt(0));
}
function encode(bytes: Uint8Array) {
  let text = '';
  for (const byte of bytes) text += String.fromCharCode(byte);
  return btoa(text);
}
const raw = globalThis.localStorage.getItem(STORAGE);
const saved = raw
  ? (JSON.parse(raw) as { version: number; document: string; snapshot?: string })
  : null;
if (saved && (saved.version !== 1 || typeof saved.document !== 'string'))
  throw new Error('Invalid saved DOCX fixture.');
const snapshot = saved?.snapshot ? decode(saved.snapshot) : undefined;
// Reopened native rooms need new client IDs. Their retained update already owns old IDs.
const previous = snapshot
  ? [...Y.decodeStateVector(Y.encodeStateVectorFromUpdate(snapshot)).keys()]
  : [];
const start = previous.length ? Math.max(...previous) + 100 : 41;
const clients: readonly [number, number] = reverse ? [start + 32, start] : [start, start + 32];
const fonts = await loadDefaultFonts({ families: ['Calibri'] });
if (fonts.failures.length) throw new Error('Packaged browser fonts failed.');
const pair = await createTwoEditorFixture({
  document: saved ? decode(saved.document) : writeOoxmlPackage(groupFixture()),
  roomSnapshot: snapshot,
  documentId: context.documentId,
  clientIds: clients,
  ownedAttributes: true,
  editor: { fonts, modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
});
const read = (editor: DocxEditorInstance) =>
  restoreReviewState(bulkSource(editor), {
    ...context,
    revision: editor.surface!.session.packageRevision(),
  });
function reviewStore(editor: DocxEditorInstance) {
  let previousPackage: ReturnType<typeof bulkSource> | undefined;
  let current: RestoreReviewResult;
  return {
    getSnapshot() {
      const pkg = bulkSource(editor);
      if (pkg !== previousPackage) {
        current = read(editor);
        previousPackage = pkg;
      }
      return current;
    },
    subscribe(notify: () => void) {
      return editor.on('change', notify);
    },
  };
}
const stores = pair.editors.map(reviewStore);
async function decide(action: 'accept' | 'reject', groupId: string | null, peer = 0) {
  const editor = pair.editors[peer]!,
    token = readEditorRevision(editor, context.documentId, context.epoch)!;
  const result =
    groupId === null
      ? (action === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
          editor,
          token,
          context
        )
      : (action === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
          editor,
          groupId,
          token,
          context
        );
  if (!result.success) throw new Error(result.error);
  await pair.settle();
}
const keepCaret = (event: React.MouseEvent) => event.preventDefault();
function ReviewPanel({ peer }: { peer: number }) {
  const store = stores[peer]!,
    result = useSyncExternalStore(store.subscribe, store.getSnapshot);
  if (!result.success) return <p role="alert">{result.error}</p>;
  const state = result.state;
  return (
    <section data-review-panel={peer}>
      <output data-group-count="">{state.groups.length}</output>
      <output data-property-count="">{state.pendingProperties.length}</output>
      {state.hasPendingReview ? (
        <div data-review-controls="">
          <button
            type="button"
            onMouseDown={keepCaret}
            disabled={state.reviewableGroupIds.length !== state.groups.length}
            onClick={() => void decide('accept', null, peer)}
          >
            Accept all
          </button>
          <button
            type="button"
            onMouseDown={keepCaret}
            disabled={state.reviewableGroupIds.length !== state.groups.length}
            onClick={() => void decide('reject', null, peer)}
          >
            Reject all
          </button>
          {state.groups.map((group, index) => (
            <div key={group.id} data-review-group={group.id}>
              <span>Change {index + 1}</span>
              <button
                type="button"
                onMouseDown={keepCaret}
                disabled={group.readOnly}
                onClick={() => void decide('accept', group.id, peer)}
              >
                Accept
              </button>
              <button
                type="button"
                onMouseDown={keepCaret}
                disabled={group.readOnly}
                onClick={() => void decide('reject', group.id, peer)}
              >
                Reject
              </button>
            </div>
          ))}
        </div>
      ) : null}
    </section>
  );
}
function Container({ block, children }: ReactBlockViewProps) {
  return (
    <div data-block={block.id} style={{ display: 'contents' }}>
      {children}
    </div>
  );
}
const roots: ReturnType<typeof createRoot>[] = [];
const hosts: ReturnType<typeof mountReactBlockHost>[] = [];
for (const [peer, container] of pair.containers.entries()) {
  const panel = document.createElement('div'),
    pane = document.createElement('section');
  pane.dataset.pane = String(peer);
  pane.className = 'editor-pane';
  pane.append(container);
  document.body.append(panel, pane);
  const root = createRoot(panel);
  roots.push(root);
  root.render(<ReviewPanel peer={peer} />);
  const host = mountReactBlockHost(pair.editors[peer]!);
  for (const type of ['lesson_phase', 'lesson_activity'])
    host.register(type, Container, { nativeContent: true });
  hosts.push(host);
}
const status = document.createElement('output');
status.dataset.saved = '';
status.textContent = 'Not saved';
const api = {
  ready: true,
  epoch: context.epoch,
  mode,
  state(peer = 0) {
    return read(pair.editors[peer]!);
  },
  async create() {
    proposeMixedBulk(pair.editors[0]!);
    await pair.settle();
  },
  async save() {
    pair.editors.forEach((editor) => editor.surface!.flushPendingInput());
    await pair.settle();
    const document = encode(new Uint8Array(await pair.editors[0]!.save()));
    globalThis.localStorage.setItem(
      STORAGE,
      JSON.stringify({
        version: 1,
        document,
        ...(mode === 'room' ? { snapshot: encode(Y.encodeStateAsUpdate(pair.documents[0]!)) } : {}),
      })
    );
    status.textContent = 'Saved';
  },
  storageKeys() {
    return Object.keys(JSON.parse(globalThis.localStorage.getItem(STORAGE)!)).sort();
  },
  decide,
  async undo(peer = 0) {
    const result = pair.editors[peer]!.exec({ type: 'undo' });
    await pair.settle();
    return result.ok;
  },
  async redo(peer = 0) {
    const result = pair.editors[peer]!.exec({ type: 'redo' });
    await pair.settle();
    return result.ok;
  },
  selection(peer = 0) {
    const surface = pair.editors[peer]!.surface!,
      selection = surface.hostElement().ownerDocument.defaultView!.getSelection();
    return {
      model: surface.state().selection,
      domParagraphId:
        selection?.anchorNode?.parentElement?.closest<HTMLElement>('.docx-paragraph-fragment')
          ?.dataset.paragraphId ?? null,
    };
  },
  canUndo(peer = 0) {
    return pair.editors[peer]!.surface!.session.canUndo();
  },
  texts(peer = 0) {
    const surface = pair.editors[peer]!.surface!;
    return surface.session.paragraphIds().map((id) => paragraphTextOf(surface.session.part(), id));
  },
  word(peer = 0) {
    const read = readOoxmlPackage(writeOoxmlPackage(bulkSource(pair.editors[peer]!)));
    if (!read.ok) throw new Error(read.reason);
    return canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!);
  },
  pause: () => pair.pause(),
  async peerAppend(text: string) {
    const surface = pair.editors[1]!.surface!,
      paragraphId = surface.session
        .paragraphIds()
        .find((id) => paragraphTextOf(surface.session.part(), id)?.includes('Keep tail'))!,
      offset = paragraphTextOf(surface.session.part(), paragraphId)!.length;
    surface.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
    surface.insertPlainText(text);
    surface.flushPendingInput();
  },
  async reconnect(receiver: 0 | 1) {
    pair.reconnect(receiver);
    await pair.settle();
  },
  remoteGeometry() {
    const remote = pair.rooms[0]!.session.remoteSelections().find(
      (selection) => selection.actorId === 'bob'
    );
    return remote
      ? pair.editors.map((editor) => {
          const caret = caretAt(editor.surface!.layout(), {
            paragraphId: remote.head.nodeId,
            offset: remote.head.offset,
          });
          return caret ? { x: caret.x, y: caret.y } : null;
        })
      : null;
  },
  async removeActivity() {
    const editor = pair.editors[0]!,
      owner = buildBlockIndex(bulkSource(editor)).byId.get('activity')!;
    const result = editor.surface!.applyAutomationOps(() => [
      { op: 'removeContentControl', controlId: owner.controlId, keepContent: false },
    ]);
    await pair.settle();
    return result.committed;
  },
  savedDigests: () => pair.savedDigests(),
};
declare global {
  interface Window {
    reviewRestore: typeof api;
  }
}
window.reviewRestore = api;
const controls = document.createElement('div');
controls.style.cssText = 'position:fixed;top:0;right:0;z-index:100;background:white';
for (const [label, action] of [
  ['Create review', () => api.create()],
  ['Save DOCX', () => api.save()],
  ['Undo', () => api.undo()],
  ['Redo', () => api.redo()],
] as const) {
  const button = document.createElement('button');
  button.type = 'button';
  button.textContent = label;
  button.addEventListener('mousedown', (event) => event.preventDefault());
  button.addEventListener('click', () => {
    void action();
  });
  controls.append(button);
}
controls.append(status);
document.body.append(controls);
window.addEventListener(
  'pagehide',
  () => {
    roots.forEach((root) => root.unmount());
    hosts.forEach((host) => host.dispose());
    pair.destroy();
  },
  { once: true }
);
