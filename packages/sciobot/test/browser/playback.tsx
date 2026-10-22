// Test-only collaboration transport under the EigenPal Pro License.
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../support/collaboration';
import { docxFixture } from '../support/document';
import { readEditorRevision } from '../../src/projection/revision-cache';
import { readWorkingView } from '../../src/review/working-view';
import { planTargetDocument } from '../../src/diff/plan-target';
import { applyDirectPlan } from '../../src/ai/apply-direct';
import { DirectChangeLedger } from '../../src/ai/change-ledger';
import { directHighlights } from '../../src/ai/acknowledge';
import { playDirectChanges, type DirectPlayback } from '../../src/ai/playback';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';

const fonts = await loadDefaultFonts({ families: ['Calibri'] });
if (fonts.failures.length) throw new Error('Packaged browser fonts failed.');
const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const filler = Array.from({ length: 30 }, (_, index) => p(`Filler paragraph ${index + 1}.`)).join(
  ''
);
const pair = await createTwoEditorFixture({
  document: docxFixture(
    p('Alpha original.') + p('Beta original.') + filler + p('Omega original.') + p('Peer line.')
  ),
  clientIds: [73, 41],
  editor: { fonts },
});
for (const [index, container] of pair.containers.entries()) {
  const pane = document.createElement('section');
  pane.dataset.pane = String(index);
  // The native reveal API scrolls the nearest `.docx-editor__scroll-container`.
  pane.className = 'editor-pane docx-editor__scroll-container';
  document.body.append(pane);
  pane.append(container);
}
const editor = pair.editors[0]!,
  ledger = new DirectChangeLedger(editor),
  context = { documentId: 'playback', epoch: 'mount' };
let playback: DirectPlayback | null = null;
let abort: AbortController | null = null;

async function markdown() {
  const token = readEditorRevision(editor, context.documentId, context.epoch)!,
    view = await readWorkingView(editor.surface!.session, token);
  if (!view.success) throw new Error(view.error);
  return { token, projection: view.view.projection };
}
const texts = () =>
  editor
    .surface!.session.paragraphIds()
    .map((id) => paragraphTextOf(editor.surface!.session.part(), id));

const api = {
  async source() {
    return (await markdown()).projection.text;
  },
  /** Plan a complete target, commit it directly as one AI edit, and start playback. */
  async apply(target: string, duration = 2400, session = 'request-1') {
    const { token, projection } = await markdown();
    const planned = planTargetDocument(
      { revision: token, pkg: editor.surface!.session.currentPackage(), projection },
      target
    );
    if (!planned.success) throw new Error(planned.error);
    const before = editor.surface!.session.part();
    const applied = ledger.mute(() => applyDirectPlan(editor, planned.plan, context));
    if (!applied.success) throw new Error(applied.error);
    ledger.record(before, { session, pending: true });
    const started = performance.now();
    abort = new AbortController();
    playback = playDirectChanges(editor, ledger, {
      duration,
      label: 'ScioBot',
      signal: abort.signal,
    });
    void playback.finished.then(() => {
      api.elapsed = performance.now() - started;
    });
    return { kinds: planned.plan.operations.map((edit) => edit.kind), motion: playback.motion };
  },
  elapsed: -1,
  async finished() {
    await playback?.finished;
    return api.elapsed;
  },
  cancel() {
    playback?.cancel();
  },
  state() {
    const visible = (selector: string) =>
      [...document.querySelectorAll<HTMLElement>(selector)].filter(
        (element) => element.style.display !== 'none'
      ).length;
    return {
      texts: texts(),
      sweeps: visible('[data-sciobot-ai-sweep]'),
      carets: visible('[data-sciobot-ai-caret]'),
      pending: ledger.changes.filter((change) => change.pending).length,
      changes: ledger.changes.length,
      highlights: directHighlights(editor, ledger).map((highlight) =>
        paragraphTextOf(editor.surface!.session.part(), highlight.paragraphId)
      ),
      canUndo: editor.surface!.session.canUndo(),
      scroll: (document.querySelector('[data-pane="0"]') as HTMLElement).scrollTop,
      hostScroll: editor.surface!.hostElement().scrollTop,
    };
  },
  /** A peer deletes a paragraph while Alice's playback runs. */
  async peerDelete(text: string) {
    const bob = pair.editors[1]!,
      part = bob.surface!.session.part(),
      id = bob
        .surface!.session.paragraphIds()
        .find((item) => paragraphTextOf(part, item) === text)!;
    const result = bob.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: id }]);
    await pair.settle();
    return result.committed;
  },
  peerTexts: () =>
    pair.editors[1]!.surface!.session.paragraphIds().map((id) =>
      paragraphTextOf(pair.editors[1]!.surface!.session.part(), id)
    ),
  savedDigests: () => pair.savedDigests(),
};
declare global {
  interface Window {
    playback: typeof api;
  }
}
window.playback = api;
const controls = document.createElement('div');
Object.assign(controls.style, { position: 'fixed', top: '0', right: '0', zIndex: '100' });
for (const [label, action] of [
  ['Abort playback', () => abort?.abort()],
  ['Detach editor', () => editor.detach()],
  ['Attach editor', () => editor.attach(pair.containers[0]!)],
] as const) {
  const button = document.createElement('button');
  button.textContent = label;
  button.addEventListener('click', action);
  controls.append(button);
}
document.body.append(controls);
window.addEventListener('pagehide', () => pair.destroy(), { once: true });
