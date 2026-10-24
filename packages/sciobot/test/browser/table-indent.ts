// Browser-only transport uses the EigenPal Pro License.
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import { canonicalOoxmlFingerprint, paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../support/collaboration';
import { tableIndentDocument, indentPart } from '../support/table-indent';
import { savedTableFingerprint } from '../support/manual-table-controls';
import { createTableIndentToolbar, manualTableIndent } from '../../src/chrome/table-indent';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';
const fonts = await loadDefaultFonts({ families: ['Calibri'] });
if (fonts.failures.length) throw new Error('Packaged fonts failed.');
const merged = new URL(globalThis.location.href).searchParams.has('merged');
const pair = await createTwoEditorFixture({
  document: tableIndentDocument(0, merged),
  editor: { fonts },
});
while (!pair.editors.every((e) => !!e.surface))
  await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
const controls = pair.editors.map((editor, index) => {
  const pane = document.createElement('section');
  pane.dataset.pane = String(index);
  const toolbar = createTableIndentToolbar(editor);
  pane.append(toolbar.element, pair.containers[index]!);
  document.body.append(pane);
  return toolbar;
});
const api = {
  select(peer = 0) {
    const s = pair.editors[peer]!.surface!,
      point = { paragraphId: s.session.paragraphIds().at(-2)!, offset: 2 };
    s.setSelection({ anchor: point, head: point });
    s.focus();
  },
  state(peer = 0) {
    const editor = pair.editors[peer]!,
      s = editor.surface!;
    s.flushPendingInput();
    return {
      indent: manualTableIndent(editor),
      table: s
        .layout()
        .pages.flatMap((p) => p.fragments)
        .find((f) => f.kind === 'table')!.box,
      selection: s.state().selection,
      paragraphs: s.session
        .paragraphIds()
        .map((id) => ({ id, text: paragraphTextOf(s.session.part(), id) })),
      fingerprint: canonicalOoxmlFingerprint(s.session.part()),
    };
  },
  expected(twips: number) {
    return canonicalOoxmlFingerprint(indentPart(twips, merged));
  },
  async settle() {
    await pair.settle();
  },
  async undo() {
    const result = pair.editors[0]!.exec({ type: 'undo' });
    await pair.settle();
    return result.ok;
  },
  async redo() {
    const result = pair.editors[0]!.exec({ type: 'redo' });
    await pair.settle();
    return result.ok;
  },
  async saved() {
    return Promise.all(pair.editors.map((e) => savedTableFingerprint(e)));
  },
  destroy() {
    controls.forEach((c) => c.dispose());
    pair.destroy();
  },
};
Object.assign(window, { tableIndentControls: api });
declare global {
  interface Window {
    tableIndentControls: typeof api;
  }
}
