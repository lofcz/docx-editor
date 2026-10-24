// Browser-only verification. The transport uses the EigenPal Pro License.
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import { paragraphTextOf, readOoxmlPackage, serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../support/collaboration';
import { docxFixture } from '../support/document';
import { prepareNativeDocument } from '../../src/native/prepare-document';
import { BASIC_MARKS, basicMarkState } from '../../src/chrome/marks';
import { createMarksToolbar, installMarkShortcuts } from '../../src/chrome/marks-toolbar';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';

const fonts = await loadDefaultFonts({ families: ['Calibri', 'Consolas'] });
if (fonts.failures.length) throw new Error('Packaged fonts failed.');
const prepared = prepareNativeDocument(docxFixture('<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p>'));
if (!prepared.success) throw new Error(prepared.error);
const pair = await createTwoEditorFixture({ document: prepared.document, editor: { fonts } });
const disposals = pair.editors.map((editor, peer) => {
  const pane = document.createElement('section');
  pane.dataset.pane = String(peer);
  const toolbar = createMarksToolbar(editor);
  pane.append(toolbar.element, pair.containers[peer]!);
  document.body.append(pane);
  const stop = installMarkShortcuts(editor);
  return () => {
    stop();
    toolbar.dispose();
  };
});
const api = {
  select(start: number, end = start) {
    const surface = pair.editors[0]!.surface!,
      id = surface.session.paragraphIds()[0]!;
    surface.setSelection({
      anchor: { paragraphId: id, offset: start },
      head: { paragraphId: id, offset: end },
    });
    surface.focus();
  },
  state(peer = 0) {
    const editor = pair.editors[peer]!,
      s = editor.surface!;
    s.flushPendingInput();
    return {
      text: paragraphTextOf(s.session.part(), s.session.paragraphIds()[0]!),
      selection: s.state().selection,
      marks: Object.fromEntries(BASIC_MARKS.map((mark) => [mark, basicMarkState(editor, mark)])),
      xml: serializeOoxmlPart(s.session.part()),
    };
  },
  async undo() {
    const result = pair.editors[0]!.exec({ type: 'undo' });
    await pair.settle();
    return result.ok;
  },
  async settle() {
    await pair.settle();
  },
  async saved() {
    const results = [];
    for (const editor of pair.editors) {
      const read = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!read.ok) throw new Error(read.reason);
      results.push(serializeOoxmlPart(read.package.parts.get(read.package.mainDocumentPart)!));
    }
    return results;
  },
  destroy() {
    disposals.forEach((stop) => stop());
    pair.destroy();
  },
};
Object.assign(window, { marksControls: api });
declare global {
  interface Window {
    marksControls: typeof api;
  }
}
