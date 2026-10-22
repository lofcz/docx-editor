// Test-only collaboration transport under the EigenPal Pro License.
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, readOoxmlPackage, storyParagraphs } from '@docx-editor.dev/core/store';
import { BlockUi, mountReactBlockHost, type ReactBlockViewProps } from '@sciobot/docx/react';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { writeBlockPayload } from '../../src/blocks/payload-part';
import { createTwoEditorFixture } from '../support/collaboration';
import { blockMeasurementDocument } from '../support/react-block-document';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';

const fonts = await loadDefaultFonts({ families: ['Calibri'] });
if (fonts.failures.length) throw new Error('Packaged browser fonts failed.');
const overflow = new URLSearchParams(location.search).get('overflow');
const transcript =
  overflow === 'long'
    ? Array.from({ length: 90 }, (_, i) => `Speaker ${i + 1}: Listening transcript turn ${i + 1}.`)
    : undefined;
const pair = await createTwoEditorFixture({
  document: blockMeasurementDocument(transcript),
  clientIds: [73, 41],
  editor: { fonts },
});
const ROW = 36;
const writeRows = (editor: DocxEditorInstance, rows: number) => {
  const block = buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('tall');
  if (!block) return false;
  return writeBlockPayload(editor, {
    controlId: block.controlId,
    payload: { ...block.payload, attributes: { rows } },
    label: 'Tall',
    fallbackOps: [],
  }).success;
};
function tallView(editor: DocxEditorInstance) {
  return function Tall({ block, paragraphs }: ReactBlockViewProps) {
    const rows = Number(block.payload.attributes.rows ?? 0);
    return (
      <BlockUi className="tall-ui" style={{ background: '#eef', border: '1px solid #557' }}>
        <button type="button" onClick={() => writeRows(editor, rows + 1)}>
          Add row
        </button>
        {Array.from({ length: rows }, (_, index) => (
          <div key={index} className="tall-row" style={{ height: ROW }}>
            Row {index + 1}
            {overflow && (
              <input
                aria-label={`Turn ${index + 1}`}
                defaultValue={paragraphs[index]?.text ?? ''}
                onBlur={(event) => {
                  const paragraph = paragraphs[index];
                  if (!paragraph || !event.currentTarget.value.startsWith(paragraph.text)) return;
                  const text = event.currentTarget.value.slice(paragraph.text.length);
                  if (text)
                    writeBlockPayload(editor, {
                      controlId: block.controlId,
                      payload: block.payload,
                      label: 'Listening transcript',
                      fallbackOps: [
                        {
                          op: 'insertText',
                          paragraphId: paragraph.id,
                          offset: paragraph.text.length,
                          text,
                        },
                      ],
                    });
                }}
              />
            )}
          </div>
        ))}
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
const hosts = pair.editors.map((editor) =>
  mountReactBlockHost(
    editor,
    new Map([['test_tall', { view: tallView(editor), reserveHeight: true }]])
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
/** Tail's laid-out line top per published paint, per session. */
const tailHistory: number[][] = [[], []];
pair.editors.forEach((editor, peer) =>
  editor.surface!.observePaint({
    afterPaint(frame) {
      const id = editor
        .surface!.session.paragraphIds()
        .find((id) => (paragraphTextOf(frame.part, id) ?? '').startsWith('Tail'));
      const caret = id && caretAt(frame.layout, { paragraphId: id, offset: 0 });
      if (caret) tailHistory[peer]!.push(caret.y);
    },
  })
);
const api = {
  idOf,
  tailHistory: (peer = 0) => [...tailHistory[peer]!],
  resetHistory: () => tailHistory.forEach((history) => (history.length = 0)),
  lineY(text: string, peer = 0) {
    surface(peer).flushPendingInput();
    return caretAt(surface(peer).layout(), { paragraphId: idOf(text, peer), offset: 0 })!.y;
  },
  text: (paragraph: string, peer = 0) =>
    paragraphTextOf(surface(peer).session.part(), idOf(paragraph, peer)),
  rows(peer = 0) {
    const block = buildBlockIndex(surface(peer).session.currentPackage()).byId.get('tall');
    return block ? Number(block.payload.attributes.rows ?? 0) : null;
  },
  canUndo: (peer = 0) => surface(peer).state().canUndo,
  revision(peer = 0) {
    surface(peer).flushPendingInput();
    return surface(peer).session.packageRevision();
  },
  /** The widget's measured bottom and Tail's painted top, in CSS pixels. */
  geometry(peer = 0) {
    const pane = document.querySelector(`[data-pane="${peer}"]`)!;
    const widget = pane.querySelector<HTMLElement>('[data-sciobot-block-measure]');
    const tail = pane.querySelector<HTMLElement>(
      `.docx-paragraph-fragment[data-paragraph-id="${idOf(api.textOfTail(peer), peer)}"]`
    );
    return {
      widgetBottom: widget ? widget.getBoundingClientRect().bottom : null,
      tailTop: tail ? tail.getBoundingClientRect().top : null,
    };
  },
  textOfTail(peer = 0) {
    const current = surface(peer);
    return (
      current.session
        .paragraphIds()
        .map((id) => paragraphTextOf(current.session.part(), id) ?? '')
        .find((text) => text.startsWith('Tail')) ?? 'Tail'
    );
  },
  remoteY() {
    const remote = pair.rooms[0]!.session.remoteSelections().find(
      (selection) => selection.actorId === 'bob'
    );
    if (!remote) return null;
    return pair.editors.map(
      (editor) =>
        caretAt(editor.surface!.layout(), {
          paragraphId: remote.head.nodeId,
          offset: remote.head.offset,
        })?.y ?? null
    );
  },
  async select(paragraph: string, offset: number, peer = 1) {
    const paragraphId = idOf(paragraph, peer);
    surface(peer).setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
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
  async removeBlock(peer = 1) {
    const current = surface(peer),
      block = buildBlockIndex(current.session.currentPackage()).byId.get('tall')!;
    const result = current.applyAutomationOps(() => [
      { op: 'removeContentControl', controlId: block.controlId, keepContent: false },
    ]);
    await pair.settle();
    return !result.rejected;
  },
  setRows: (rows: number, peer = 0) => writeRows(pair.editors[peer]!, rows),
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
  settle: () => pair.settle(),
  statuses: () => pair.rooms.map((room) => room.session.status()),
  savedDigests: () => pair.savedDigests(),
  async savedTexts(peer = 0) {
    const read = readOoxmlPackage(new Uint8Array(await pair.editors[peer]!.save()));
    if (!read.ok) throw new Error(read.reason);
    const part = read.package.parts.get('/word/document.xml')!;
    const body = part.root.children.find((node) => node.kind === 'body');
    if (!body) throw new Error('Saved Word body is missing.');
    return [...storyParagraphs(body)].map((p) => paragraphTextOf(part, p.id));
  },
  viewport(peer = 0) {
    const pane = pair.containers[peer]!;
    const root = pane.querySelector<HTMLElement>('[data-sciobot-block-measure]')!;
    const page = root.closest('.docx-page-content')!;
    return {
      height: root.clientHeight,
      scrollHeight: root.scrollHeight,
      bottom: root.getBoundingClientRect().bottom,
      pageBottom: page.getBoundingClientRect().bottom,
      roots: pane.querySelectorAll('[data-sciobot-block-measure]').length,
      nativePages: surface(peer).layout().pages.length,
    };
  },
};
declare global {
  interface Window {
    blockMeasurement: typeof api;
  }
}
window.blockMeasurement = api;
window.addEventListener(
  'pagehide',
  () => {
    hosts.forEach((host) => host.dispose());
    pair.destroy();
  },
  { once: true }
);
