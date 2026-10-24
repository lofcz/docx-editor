// Browser-only verification. The transport uses the EigenPal Pro License.
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import {
  directParagraphProperties,
  paragraphTextOf,
  readOoxmlPackage,
  canonicalOoxmlFingerprint,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../support/collaboration';
import { docxFixture } from '../support/document';
import { createColorToolbar, textColorState } from '../../src/chrome/colors';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';
const fonts = await loadDefaultFonts({ families: ['Calibri'] });
if (fonts.failures.length) throw new Error('Packaged fonts failed.');
const pair = await createTwoEditorFixture({
  editor: { fonts },
  document: docxFixture(
    '<w:p><w:pPr><w:spacing w:line="251" w:lineRule="exact" w:before="123"/></w:pPr><w:r><w:t>Imported</w:t></w:r></w:p><w:p><w:r><w:rPr><w:color w:val="CC0000"/><w:shd w:val="clear" w:fill="FFFF00"/></w:rPr><w:t>Alpha beta</w:t></w:r></w:p>'
  ),
});
const controls = pair.editors.map((editor, peer) => {
  const pane = document.createElement('section');
  pane.dataset.pane = String(peer);
  const toolbar = createColorToolbar(editor, ['#CC0000', '#FFFF00', '#17A2B8']);
  pane.append(toolbar.element, pair.containers[peer]!);
  document.body.append(pane);
  return toolbar;
});
const api = {
  select(index: number, start: number, end = start) {
    const s = pair.editors[0]!.surface!,
      id = s.session.paragraphIds()[index]!;
    s.setSelection({
      anchor: { paragraphId: id, offset: start },
      head: { paragraphId: id, offset: end },
    });
    s.focus();
  },
  state(peer = 0) {
    const editor = pair.editors[peer]!,
      s = editor.surface!;
    s.flushPendingInput();
    return {
      paragraphs: s.session.paragraphIds().map((id) => ({
        id,
        text: paragraphTextOf(s.session.part(), id),
        properties: directParagraphProperties(s.session.part(), id),
      })),
      selection: s.state().selection,
      foreground: textColorState(editor, 'foreground'),
      background: textColorState(editor, 'background'),
      fingerprint: canonicalOoxmlFingerprint(s.session.part()),
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
    const values = [];
    for (const editor of pair.editors) {
      const read = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!read.ok) throw new Error(read.reason);
      values.push(
        canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!)
      );
    }
    return values;
  },
  destroy() {
    controls.forEach((c) => c.dispose());
    pair.destroy();
  },
};
Object.assign(window, { colorControls: api });
declare global {
  interface Window {
    colorControls: typeof api;
  }
}
