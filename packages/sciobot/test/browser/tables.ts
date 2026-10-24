// Browser-only verification. The transport uses the EigenPal Pro License.
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { paragraphsInCells } from '@docx-editor.dev/core/layout';
import { createTwoEditorFixture } from '../support/collaboration';
import { docxFixture } from '../support/document';
import { manualTableDocument, recoverFirstTableCellSource } from '../support/manual-table';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../../src/review/native-module';
import { createTableRowToolbar } from '../../src/chrome/table-rows';
import {
  createTableInsertionGrid,
  installTableNavigation,
  nativeTableAnchor,
  nativeTableCellIds,
} from '../../src/chrome/tables';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';
const fonts = await loadDefaultFonts({ families: ['Calibri'] });
if (fonts.failures.length) throw new Error('Packaged fonts failed.');
const parameters = new URL(globalThis.location.href).searchParams;
const rowMode = parameters.has('rows');
const pair = await createTwoEditorFixture({
  editor: { fonts, modules: rowMode ? [SCIOBOT_NATIVE_REVIEW_MODULE] : [] },
  document: rowMode
    ? manualTableDocument()
    : docxFixture(
        '<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
      ),
});
const recovered = parameters.has('recovered') ? await recoverFirstTableCellSource(pair) : null;
const controls = pair.editors.map((editor, peer) => {
  const pane = document.createElement('section');
  pane.dataset.pane = String(peer);
  const grid = createTableInsertionGrid(editor),
    stop = installTableNavigation(editor);
  const rows = rowMode ? createTableRowToolbar(editor) : null;
  pane.append(grid.element, pair.containers[peer]!);
  if (rows) pane.prepend(rows.element);
  document.body.append(pane);
  return () => {
    grid.dispose();
    rows?.dispose();
    stop();
  };
});
const api = {
  selectRecovered() {
    if (!recovered) throw new Error('No recovered source fixture.');
    const s = pair.editors[0]!.surface!,
      point = { paragraphId: recovered, offset: 2 };
    s.setSelection({ anchor: point, head: point });
    s.focus();
  },
  select(index: number, offset = 0) {
    const s = pair.editors[0]!.surface!,
      point = { paragraphId: s.session.paragraphIds()[index]!, offset };
    s.setSelection({ anchor: point, head: point });
    s.focus();
  },
  selectCell(index: number) {
    const editor = pair.editors[0]!,
      s = editor.surface!,
      anchor = nativeTableAnchor(editor)!;
    const cell = nativeTableCellIds(editor, anchor.tableId)[index]!;
    const point = { paragraphId: paragraphsInCells(s.layout(), [cell])[0]!, offset: 0 };
    s.setSelection({ anchor: point, head: point });
    s.focus();
  },
  state(peer = 0) {
    const editor = pair.editors[peer]!,
      s = editor.surface!;
    s.flushPendingInput();
    const tables = [
      ...new Set(
        s
          .layout()
          .pages.flatMap((p) => p.fragments.filter((f) => f.kind === 'table').map((f) => f.tableId))
      ),
    ];
    return {
      tables: tables.map((id) => ({ id, cells: nativeTableCellIds(editor, id) })),
      anchor: nativeTableAnchor(editor),
      selection: s.state().selection,
      paragraphs: s.session
        .paragraphIds()
        .map((id) => ({ id, text: paragraphTextOf(s.session.part(), id) })),
      fingerprint: canonicalOoxmlFingerprint(s.session.part()),
    };
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
    controls.forEach((stop) => stop());
    pair.destroy();
  },
};
Object.assign(window, { tableControls: api });
declare global {
  interface Window {
    tableControls: typeof api;
  }
}
