// Browser-only transport uses the EigenPal Pro License.
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../support/collaboration';
import { listDocument, listParagraph } from '../support/list-document';
import { createListToolbar, manualListState, manualListMarker } from '../../src/chrome/lists';
import { installListEditing } from '../../src/chrome/list-input';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';
const context = { documentId: 'manual-lists', epoch: 'mount' };
const fonts = await loadDefaultFonts({ families: ['Calibri'] });
if (fonts.failures.length) throw new Error('Packaged fonts failed.');
const pair = await createTwoEditorFixture({
  document: listDocument(
    listParagraph('Alpha') +
      listParagraph('Beta') +
      listParagraph('Imported', '3') +
      listParagraph('Tail')
  ),
  documentId: context.documentId,
  ownedAttributes: true,
  editor: { fonts },
});
while (!pair.editors.every((e) => !!e.surface))
  await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
const controls = pair.editors.map((editor, index) => {
  const pane = document.createElement('section');
  pane.dataset.pane = String(index);
  const toolbar = createListToolbar(editor, context),
    stop = installListEditing(editor, context);
  pane.append(toolbar.element, pair.containers[index]!);
  document.body.append(pane);
  return () => {
    toolbar.dispose();
    stop();
  };
});
const api = {
  select(index: number, start = 0, end = start) {
    const s = pair.editors[0]!.surface!,
      id = s.session.paragraphIds()[index]!;
    s.setSelection({
      anchor: { paragraphId: id, offset: start },
      head: { paragraphId: id, offset: end },
    });
    s.focus();
  },
  state(peer = 0) {
    const e = pair.editors[peer]!,
      s = e.surface!;
    s.flushPendingInput();
    return {
      selection: s.state().selection,
      paragraphs: s.session.paragraphIds().map((id) => ({
        id,
        text: paragraphTextOf(s.session.part(), id),
        membership: manualListState(e, context, id),
        marker: manualListMarker(e, id),
      })),
      fingerprint: canonicalOoxmlFingerprint(s.session.part()),
    };
  },
  async settle() {
    await pair.settle();
  },
  async undo() {
    const r = pair.editors[0]!.exec({ type: 'undo' });
    await pair.settle();
    return r.ok;
  },
  async redo() {
    const r = pair.editors[0]!.exec({ type: 'redo' });
    await pair.settle();
    return r.ok;
  },
  async saved() {
    const values = [];
    for (const e of pair.editors) {
      const read = readOoxmlPackage(new Uint8Array(await e.save()));
      if (!read.ok) throw new Error(read.reason);
      values.push({
        fingerprint: canonicalOoxmlFingerprint(
          read.package.parts.get(read.package.mainDocumentPart)!
        ),
        digest: semanticDigest(read.package.parts.values()),
      });
    }
    return values;
  },
  destroy() {
    controls.forEach((stop) => stop());
    pair.destroy();
  },
};
Object.assign(window, { listControls: api });
declare global {
  interface Window {
    listControls: typeof api;
  }
}
