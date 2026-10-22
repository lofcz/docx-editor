// Browser-only verification. The transport uses the EigenPal Pro License.
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import {
  directParagraphProperties,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../support/document';
import { createTwoEditorFixture } from '../support/collaboration';
import { withBlockTypeStyles } from '../../src/chrome/block-styles';
import { createBlockToolbar } from '../../src/chrome/block-toolbar';
import { installBlockInputRules } from '../../src/chrome/block-input';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';

const fonts = await loadDefaultFonts({ families: ['Calibri'] });
if (fonts.failures.length) throw new Error('Packaged fonts failed.');
const loaded = readOoxmlPackage(
  docxFixture(
    '<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p><w:p/><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
  )
);
if (!loaded.ok) throw new Error(loaded.reason);
const pair = await createTwoEditorFixture({
  document: writeOoxmlPackage(withBlockTypeStyles(loaded.package)!),
  editor: { fonts },
});
const disposals = pair.editors.map((editor, peer) => {
  const pane = document.createElement('section');
  pane.dataset.pane = String(peer);
  pane.append(
    createBlockToolbar(editor, { documentId: 'blocks', epoch: 'mount' }),
    pair.containers[peer]!
  );
  document.body.append(pane);
  return installBlockInputRules(editor, { documentId: 'blocks', epoch: 'mount' });
});
const api = {
  select(index: number, start: number, end = start) {
    const surface = pair.editors[0]!.surface!,
      id = surface.session.paragraphIds()[index]!;
    surface.setSelection({
      anchor: { paragraphId: id, offset: start },
      head: { paragraphId: id, offset: end },
    });
    surface.focus();
  },
  state(peer = 0) {
    const surface = pair.editors[peer]!.surface!;
    surface.flushPendingInput();
    return {
      paragraphs: surface.session.paragraphIds().map((id) => ({
        id,
        text: paragraphTextOf(surface.session.part(), id),
        style:
          directParagraphProperties(surface.session.part(), id).find(
            (p) => p.localName === 'pStyle'
          )?.attributes?.val ?? null,
      })),
      selection: surface.state().selection,
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
    return pair.savedDigests();
  },
  destroy() {
    disposals.forEach((stop) => stop());
    pair.destroy();
  },
};
Object.assign(window, { blockControls: api });
declare global {
  interface Window {
    blockControls: typeof api;
  }
}
