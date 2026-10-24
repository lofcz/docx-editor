// Browser-only transport uses the EigenPal Pro License.
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import {
  canonicalOoxmlFingerprint,
  bodyStoryRoot,
  storyParagraphs,
  directParagraphProperties,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../support/collaboration';
import { listDocument, listParagraph } from '../support/list-document';
import { createListToolbar, manualListState, manualListMarker } from '../../src/chrome/lists';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../../src/review/native-module';
import { readReviewGroups } from '../../src/review/groups';
import { acceptReviewGroup } from '../../src/review/accept-group';
import { rejectReviewGroup } from '../../src/review/reject-group';
import { readEditorRevision } from '../../src/projection/revision-cache';
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
  editor: { fonts, modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
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
  selectRange(first: number, start: number, last: number, end: number) {
    const s = pair.editors[0]!.surface!,
      ids = s.session.paragraphIds();
    s.setSelection({
      anchor: { paragraphId: ids[first]!, offset: start },
      head: { paragraphId: ids[last]!, offset: end },
    });
    s.focus();
  },
  suggest() {
    return pair.editors[0]!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok;
  },
  groups() {
    const result = readReviewGroups(pair.editors[0]!.surface!.session.currentPackage());
    if (!result.success) throw new Error(result.error);
    return result.groups;
  },
  accept(id: string) {
    const editor = pair.editors[0]!;
    return acceptReviewGroup(
      editor,
      id,
      readEditorRevision(editor, context.documentId, context.epoch)!,
      context
    ).success;
  },
  reject(id: string) {
    const editor = pair.editors[0]!;
    return rejectReviewGroup(
      editor,
      id,
      readEditorRevision(editor, context.documentId, context.epoch)!,
      context
    ).success;
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
        properties: directParagraphProperties(s.session.part(), id),
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
      const main = read.package.parts.get(read.package.mainDocumentPart)!;
      values.push({
        fingerprint: canonicalOoxmlFingerprint(
          read.package.parts.get(read.package.mainDocumentPart)!
        ),
        digest: semanticDigest(read.package.parts.values()),
        paragraphs: storyParagraphs(bodyStoryRoot(main)!).map((p) => paragraphTextOf(main, p.id)),
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
