// Browser-only verification. The transport uses the EigenPal Pro License.
import { packagedFonts } from '@docx-editor.dev/fonts';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { paragraphFragmentsOf } from '@docx-editor.dev/core/layout';
import { createTwoEditorFixture } from '../support/collaboration';
import { fontLinkDocument, selectFontLink } from '../support/font-link-document';
import { createFontToolbar, manualFontState } from '../../src/chrome/fonts';
import { createLinkToolbar } from '../../src/chrome/links';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';
const pair = await createTwoEditorFixture({
  document: fontLinkDocument(),
  editor: { fonts: packagedFonts() },
});
while (!pair.editors.every((e) => !!e.surface))
  await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
const controls = pair.editors.map((editor, index) => {
  const pane = document.createElement('section');
  pane.dataset.pane = String(index);
  const fonts = createFontToolbar(editor),
    links = createLinkToolbar(editor);
  pane.append(fonts.element, links.element, pair.containers[index]!);
  document.body.append(pane);
  return () => {
    fonts.dispose();
    links.dispose();
  };
});
const api = {
  select(start: number, end = start) {
    selectFontLink(pair.editors[0]!, start, end);
    pair.editors[0]!.surface!.focus();
  },
  state(peer = 0) {
    const editor = pair.editors[peer]!,
      s = editor.surface!;
    s.flushPendingInput();
    const id = s.session.paragraphIds()[0]!;
    return {
      id,
      text: paragraphTextOf(s.session.part(), id),
      selection: s.state().selection,
      font: manualFontState(editor),
      measurement: editor.fontMeasurement(),
      links: s.hyperlinks.linksInCaretParagraph(),
      spans: s
        .layout()
        .pages.flatMap((page) => paragraphFragmentsOf(page))
        .filter((p) => p.paragraphId === id)
        .flatMap((p) =>
          p.lines.flatMap((l) =>
            l.spans.map((span) => ({
              family: span.style.fontFamily,
              points: span.style.fontSizePt,
              start: span.range.start,
              end: span.range.end,
            }))
          )
        ),
      fingerprint: canonicalOoxmlFingerprint(s.session.part()),
    };
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
    const result = [];
    for (const editor of pair.editors) {
      const read = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!read.ok) throw new Error(read.reason);
      result.push(
        canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!)
      );
    }
    return result;
  },
  destroy() {
    controls.forEach((dispose) => dispose());
    pair.destroy();
  },
};
Object.assign(window, { fontLinkControls: api });
declare global {
  interface Window {
    fontLinkControls: typeof api;
  }
}
