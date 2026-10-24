// Browser-only transport uses the EigenPal Pro License.
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../support/collaboration';
import {
  manualContext,
  manualTrackingDocument,
  manualWave,
  manualRevisionFingerprints,
  manualParagraphFingerprints,
  manualToken,
  selectManual,
  type ManualWave,
} from '../support/manual-tracking';
import {
  createManualSuggestionToggle,
  manualSuggestionState,
} from '../../src/review/manual-tracking';
import { createMarksToolbar } from '../../src/chrome/marks-toolbar';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../../src/review/native-module';
import { acceptReviewGroup } from '../../src/review/accept-group';
import { rejectReviewGroup } from '../../src/review/reject-group';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';
const fonts = await loadDefaultFonts({ families: ['Calibri'] });
if (fonts.failures.length) throw new Error('Packaged fonts failed.');
const pair = await createTwoEditorFixture({
  document: manualTrackingDocument(),
  documentId: manualContext.documentId,
  ownedAttributes: true,
  editor: { fonts, modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
});
while (!pair.editors.every((e) => !!e.surface))
  await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
let previous: ManualWave | undefined,
  waveCount = 0,
  error = '';
const controls = pair.editors.map((editor, index) => {
  const pane = document.createElement('section');
  pane.dataset.pane = String(index);
  const mode = createManualSuggestionToggle(editor),
    marks = createMarksToolbar(editor);
  pane.append(mode.element, marks.element, pair.containers[index]!);
  document.body.append(pane);
  return () => {
    mode.dispose();
    marks.dispose();
  };
});
const actions = document.createElement('div');
actions.setAttribute('role', 'toolbar');
actions.setAttribute('aria-label', 'AI proposal actions');
const button = (label: string, run: () => Promise<void>) => {
  const b = document.createElement('button');
  b.type = 'button';
  b.textContent = label;
  b.addEventListener('mousedown', (e) => {
    if (e.button === 0) e.preventDefault();
  });
  b.addEventListener('click', () => {
    b.disabled = true;
    void run()
      .catch((cause) => {
        error = String(cause);
      })
      .finally(() => {
        b.disabled = false;
      });
  });
  actions.append(b);
};
button('Next AI wave', async () => {
  previous = await manualWave(pair.editors[0]!, waveCount === 0 ? 'ONE' : 'TWO', previous);
  await pair.settle();
  waveCount++;
});
for (const action of ['accept', 'reject'] as const)
  button(action === 'accept' ? 'Accept AI proposal' : 'Reject AI proposal', async () => {
    if (!previous) throw new Error('No AI proposal.');
    const editor = pair.editors[0]!,
      decide = action === 'accept' ? acceptReviewGroup : rejectReviewGroup;
    const result = decide(
      editor,
      previous.ownership.at(-1)!.groupId,
      manualToken(editor),
      manualToken(editor)
    );
    if (!result.success) throw new Error(result.error);
    await pair.settle();
  });
document.body.prepend(actions);
const api = {
  select(index: number, start: number, end = start) {
    selectManual(pair.editors[0]!, index, start, end);
    pair.editors[0]!.surface!.focus();
  },
  state(peer = 0) {
    const editor = pair.editors[peer]!,
      s = editor.surface!;
    s.flushPendingInput();
    return {
      mode: manualSuggestionState(editor).active,
      author: editor.getConfiguredAuthor(),
      manual: manualRevisionFingerprints(editor, 'Alice'),
      manualParagraphs: manualParagraphFingerprints(editor),
      waveCount,
      error,
      sites: collectRevisionSites(s.session.part()).length,
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
Object.assign(window, { manualTrackingControls: api });
declare global {
  interface Window {
    manualTrackingControls: typeof api;
  }
}
