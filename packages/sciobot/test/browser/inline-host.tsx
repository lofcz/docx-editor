// Test-only collaboration transport uses the EigenPal Pro License.
import { createContext, useContext, useEffect, useState } from 'react';
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { caretAt, hitTestSemantic, linesOf } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, readOoxmlPackage, semanticDigest } from '@docx-editor.dev/core/store';
import { BlockUi, mountReactInlineHost, type ReactBlockViewProps } from '@sciobot/docx/react';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { writeBlockPayload } from '../../src/blocks/payload-part';
import { createTwoEditorFixture } from '../support/collaboration';
import { inlineWidgetDocument } from '../support/react-block-document';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';

const fonts = await loadDefaultFonts({ families: ['Calibri'] });
if (fonts.failures.length) throw new Error('Packaged fonts failed.');
const reverse = new URL(location.href).searchParams.has('reverse');
const pair = await createTwoEditorFixture({
  document: inlineWidgetDocument(),
  clientIds: reverse ? [73, 41] : [41, 73],
  editor: { fonts },
});
const Label = createContext('source');
const events = [
  { mounted: 0, unmounted: 0 },
  { mounted: 0, unmounted: 0 },
];
const views = pair.editors.map(
  (_, peer) =>
    function Chip({ block }: ReactBlockViewProps) {
      const label = useContext(Label),
        [count, setCount] = useState(0);
      useEffect(() => {
        events[peer]!.mounted++;
        return () => {
          events[peer]!.unmounted++;
        };
      }, []);
      return (
        <BlockUi
          style={{ display: 'inline-flex', alignItems: 'baseline', font: '14px/20px Arial' }}
        >
          <button type="button" onClick={() => setCount((value) => value + 1)}>
            {label} {count}
          </button>
          <input aria-label="Local note" defaultValue="" style={{ width: 48 }} />
          <span data-payload-name="">{String(block.payload.attributes.name)}</span>
        </BlockUi>
      );
    }
);
for (const [peer, container] of pair.containers.entries()) {
  const pane = document.createElement('section');
  pane.className = 'editor-pane';
  pane.dataset.pane = String(peer);
  pane.append(container);
  document.body.append(pane);
}
const hosts = pair.editors.map((editor, peer) =>
  mountReactInlineHost(editor, new Map([['source_chip', views[peer]!]]), {
    providers: (children) => (
      <Label.Provider value={peer === 0 ? 'source' : 'peer source'}>{children}</Label.Provider>
    ),
  })
);
const surface = (peer = 0) => pair.editors[peer]!.surface!;
const id = (peer = 0) => surface(peer).session.paragraphIds()[0]!;
const spanOf = (peer = 0) =>
  linesOf(surface(peer).layout())
    .flatMap((line) => line.spans)
    .find((span) => span.inlineWidget);
let reopened: ReturnType<typeof createDocxEditor> | undefined;
let reopenedHost: ReturnType<typeof mountReactInlineHost> | undefined;
const api = {
  events,
  text: (peer = 0) => paragraphTextOf(surface(peer).session.part(), id(peer)),
  revision: (peer = 0) => surface(peer).session.packageRevision(),
  selection: (peer = 0) => surface(peer).state().selection,
  async select(offset: number, peer = 0) {
    const position = { paragraphId: id(peer), offset };
    surface(peer).setSelection({ anchor: position, head: position });
    await pair.settle();
  },
  providers(label: string) {
    hosts[0]!.setProviders((children) => <Label.Provider value={label}>{children}</Label.Provider>);
  },
  zoom: (zoom: number) => pair.editors[0]!.setZoom(zoom).ok,
  refresh: () => hosts[0]!.refresh(),
  async rename(name: string, peer = 0) {
    const editor = pair.editors[peer]!,
      block = buildBlockIndex(surface(peer).session.currentPackage()).byId.get('source-1')!;
    const result = writeBlockPayload(editor, {
      controlId: block.controlId,
      payload: { ...block.payload, attributes: { name } },
      label: 'Source',
      fallbackOps: [],
    });
    await pair.settle();
    return result.success;
  },
  async insert(text: string, peer = 1) {
    const current = surface(peer),
      paragraphId = id(peer),
      offset = paragraphTextOf(current.session.part(), paragraphId)!.length;
    const result = current.applyAutomationOps(() => [
      { op: 'insertText', paragraphId, offset, text },
    ]);
    await pair.settle();
    return !result.rejected;
  },
  async remove(peer = 0) {
    const current = surface(peer),
      block = buildBlockIndex(current.session.currentPackage()).byId.get('source-1')!;
    const result = current.applyAutomationOps(() => [
      { op: 'removeContentControl', controlId: block.controlId, keepContent: false },
    ]);
    await pair.settle();
    return !result.rejected;
  },
  pause: () => pair.pause(),
  async reconnect(receiver: 0 | 1) {
    pair.reconnect(receiver);
    await pair.settle();
  },
  async exec(type: 'undo' | 'redo') {
    const result = pair.editors[0]!.exec({ type });
    await pair.settle();
    return result.ok;
  },
  settle: () => pair.settle(),
  savedDigests: () => pair.savedDigests(),
  remoteGeometry() {
    const remote = pair.rooms[0]!.session.remoteSelections().find(
      (selection) => selection.actorId === 'bob'
    );
    if (!remote) return null;
    const caret = caretAt(
      surface().layout(),
      {
        paragraphId: remote.head.nodeId,
        offset: remote.head.offset,
      },
      surface().textMeasurer?.()
    );
    const element = pair.containers[0]!.querySelector<HTMLElement>('.docx-remote-caret');
    const content = pair.containers[0]!.querySelector<HTMLElement>('.docx-page-content');
    if (!caret || !element || !content) return null;
    const rect = element.getBoundingClientRect(),
      origin = content.getBoundingClientRect(),
      scale = pair.editors[0]!.getRenderScale();
    return {
      x: rect.x - origin.x,
      y: rect.y - origin.y,
      height: rect.height,
      expected: { x: caret.x * scale, y: caret.y * scale, height: caret.height * scale },
    };
  },
  geometry(peer = 0) {
    const current = surface(peer),
      span = spanOf(peer);
    if (!span) return null;
    const line = linesOf(current.layout()).find((line) => line.spans.includes(span))!;
    const root = pair.containers[peer]!.querySelector<HTMLElement>('[data-sciobot-block-host]')!;
    const content = root.parentElement!.getBoundingClientRect(),
      rect = root.getBoundingClientRect();
    const baseline = root
      .querySelector<HTMLElement>('[data-sciobot-inline-baseline]')!
      .getBoundingClientRect();
    const scale = pair.editors[peer]!.getRenderScale(),
      widget = span.inlineWidget!;
    return {
      width: rect.width,
      height: rect.height,
      x: rect.x - content.x,
      y: rect.y - content.y,
      baseline: baseline.top - content.top,
      mouse: {
        // Chrome hit testing rounds subpixel points within one pixel onto widget controls.
        left: { x: rect.left - 1.5, y: baseline.top - 2 * scale },
        right: { x: rect.right + 1.5, y: baseline.top - 2 * scale },
      },
      expected: {
        width: widget.width * scale,
        height: (widget.ascent + widget.descent) * scale,
        x: span.box.x * scale,
        y: (line.box.y + line.baseline - widget.ascent) * scale,
        baseline: (line.box.y + line.baseline) * scale,
      },
      hits: [span.box.x + 0.01, span.box.x + widget.width - 0.01].map(
        (x) =>
          hitTestSemantic(current.layout(), { x, y: line.box.y + line.baseline })?.position.offset
      ),
      range: span.range,
    };
  },
  async reopen() {
    const bytes = new Uint8Array(await pair.editors[0]!.save());
    const read = readOoxmlPackage(bytes);
    if (!read.ok) throw new Error(read.reason);
    const container = document.createElement('div');
    document.body.append(container);
    reopenedHost?.dispose();
    reopened?.destroy();
    reopened = createDocxEditor({ container, document: bytes, fonts });
    reopenedHost = mountReactInlineHost(reopened, new Map([['source_chip', views[0]!]]));
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    return {
      saved: semanticDigest(read.package.parts.values()),
      reopened: semanticDigest(reopened.surface!.session.currentPackage().parts.values()),
      measured: linesOf(reopened.surface!.layout()).some((line) =>
        line.spans.some((span) => span.inlineWidget)
      ),
      text: paragraphTextOf(
        reopened.surface!.session.part(),
        reopened.surface!.session.paragraphIds()[0]!
      ),
    };
  },
  dispose: () => hosts[0]!.dispose(),
};
declare global {
  interface Window {
    inlineHost: typeof api;
  }
}
window.inlineHost = api;
window.addEventListener(
  'pagehide',
  () => {
    hosts.forEach((host) => host.dispose());
    reopenedHost?.dispose();
    reopened?.destroy();
    pair.destroy();
  },
  { once: true }
);
