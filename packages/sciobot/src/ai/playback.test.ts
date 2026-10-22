import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { paragraphTextOf, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { protectedTargetInput, targetParagraph as p } from '../../test/support/target-document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { planTargetDocument } from '../diff/plan-target';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { directHighlights } from './acknowledge';
import { applyDirectPlan } from './apply-direct';
import { DirectChangeLedger } from './change-ledger';
import { PLAYBACK_CAP_MS, playDirectChanges } from './playback';

const context = { documentId: 'playback', epoch: 'mount' };
const bytes = async () =>
  writeOoxmlPackage((await protectedTargetInput(p('Alpha.') + p('Beta.') + p('Gamma.'))).pkg);
function mount(document_: Uint8Array) {
  const container = document.createElement('div');
  container.className = 'docx-editor__scroll-container';
  document.body.append(container);
  return createDocxEditor({ container, document: document_ });
}
async function commit(editor: DocxEditorInstance, ledger: DirectChangeLedger, target: string) {
  const token = readEditorRevision(editor, context.documentId, context.epoch)!,
    view = await readWorkingView(editor.surface!.session, token);
  if (!view.success) throw new Error(view.error);
  const planned = planTargetDocument(
    {
      revision: token,
      pkg: editor.surface!.session.currentPackage(),
      projection: view.view.projection,
    },
    target
  );
  if (!planned.success) throw new Error(planned.error);
  const before = editor.surface!.session.part();
  const result = ledger.mute(() => applyDirectPlan(editor, planned.plan, context));
  if (!result.success) throw new Error(result.error);
  ledger.record(before, { session: 'request', pending: true });
}
const texts = (editor: DocxEditorInstance) =>
  editor
    .surface!.session.paragraphIds()
    .map((id) => paragraphTextOf(editor.surface!.session.part(), id));
const highlighted = (editor: DocxEditorInstance, ledger: DirectChangeLedger) =>
  directHighlights(editor, ledger).map((item) =>
    paragraphTextOf(editor.surface!.session.part(), item.paragraphId)
  );
const overlays = () =>
  document.querySelectorAll('[data-sciobot-ai-sweep], [data-sciobot-ai-caret]').length;

for (const transition of ['detach', 'switch', 'destroy'] as const)
  test(`playback stops on ${transition} without waiting for its timer or rolling back the edit`, async () => {
    const editor = mount(await bytes());
    const ledger = new DirectChangeLedger(editor);
    try {
      await commit(editor, ledger, 'Alpha changed.\n\nBeta.\n\nGamma.');
      const saved = new Uint8Array(await editor.save());
      const replacement = transition === 'switch' ? await bytes() : null;
      const surface = editor.surface!;
      const container = surface.hostElement().parentElement!;
      const running = playDirectChanges(editor, ledger, { duration: 4000, reducedMotion: false });
      expect(overlays()).toBeGreaterThan(0);
      let finished = false;
      void running.finished.then(() => {
        finished = true;
      });
      if (transition === 'detach') editor.detach();
      else if (transition === 'switch') {
        editor.load(replacement!);
        await editor.save(); // Settle any scheduled document open.
      } else editor.destroy();
      await new Promise<void>((resolve) => queueMicrotask(resolve));
      await new Promise<void>((resolve) => queueMicrotask(resolve));
      expect(overlays()).toBe(0);
      expect(finished).toBe(true);
      expect(ledger.changes.every((change) => !change.pending)).toBe(true);
      if (transition === 'detach') {
        editor.attach(container);
        expect(texts(editor)).toEqual(['Alpha changed.', 'Beta.', 'Gamma.']);
        expect(new Uint8Array(await editor.save())).toEqual(saved);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(texts(editor)).toEqual(['Alpha.', 'Beta.', 'Gamma.']);
      } else if (transition === 'switch') {
        expect(texts(editor)).toEqual(['Alpha.', 'Beta.', 'Gamma.']);
        expect(highlighted(editor, ledger)).toEqual([]);
      }
    } finally {
      ledger.dispose();
      editor.destroy();
    }
  }, 60_000);

for (const receiver of [0, 1] as const)
  test(`detaching playback keeps concurrent peer text and shared undo/redo; receiver ${receiver}`, async () => {
    const pair = await createTwoEditorFixture({ document: await bytes(), clientIds: [73, 41] });
    const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
    const ledger = new DirectChangeLedger(alice);
    try {
      await commit(alice, ledger, 'Alpha changed.\n\nBeta.\n\nGamma.');
      const running = playDirectChanges(alice, ledger, { duration: 4000, reducedMotion: false });
      pair.pause();
      const gamma = bob
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Gamma.')!;
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: gamma, offset: 6, text: ' Bob' },
        ]).committed
      ).toBe(true);
      alice.detach();
      await running.finished;
      expect(overlays()).toBe(0);
      pair.reconnect(receiver);
      alice.attach(pair.containers[0]!);
      await pair.settle();
      for (const editor of pair.editors)
        expect(texts(editor)).toEqual(['Alpha changed.', 'Beta.', 'Gamma. Bob']);
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(texts(editor)).toEqual(['Alpha.', 'Beta.', 'Gamma. Bob']);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
    } finally {
      ledger.dispose();
      pair.destroy();
    }
  }, 60_000);

test('playback commits nothing, reveals every change, and clamps a burst to one 4000 ms budget', async () => {
  const editor = mount(await bytes()),
    ledger = new DirectChangeLedger(editor);
  try {
    await commit(editor, ledger, 'Alpha changed.\n\nBeta.\n\nGamma changed.');
    const revision = editor.surface!.session.packageRevision(),
      started = Date.now();
    const first = playDirectChanges(editor, ledger, { duration: 60_000, reducedMotion: false });
    expect(first.motion).toBe(true);
    // The committed target is complete before playback; pending changes are not painted yet.
    expect(texts(editor)).toEqual(['Alpha changed.', 'Beta.', 'Gamma changed.']);
    expect(highlighted(editor, ledger)).toEqual([]);
    expect(overlays()).toBeGreaterThan(0);
    // A second playback in the same burst cancels the first and inherits its deadline.
    await new Promise((resolve) => setTimeout(resolve, 300));
    await commit(editor, ledger, 'Alpha changed.\n\nBeta changed.\n\nGamma changed.');
    const second = playDirectChanges(editor, ledger, { duration: 60_000, reducedMotion: false });
    await first.finished;
    await second.finished;
    expect(Date.now() - started).toBeLessThan(PLAYBACK_CAP_MS + 500);
    expect(overlays()).toBe(0);
    expect(ledger.changes.every((change) => !change.pending)).toBe(true);
    expect(highlighted(editor, ledger)).toEqual([
      'Alpha changed.',
      'Beta changed.',
      'Gamma changed.',
    ]);
    expect(editor.surface!.session.packageRevision()).not.toBe(revision);
  } finally {
    ledger.dispose();
    editor.destroy();
  }
}, 30_000);

test('cancel, abort, reduced motion, and undo during playback keep the document and leave no stale highlight', async () => {
  const editor = mount(await bytes()),
    ledger = new DirectChangeLedger(editor);
  try {
    await commit(editor, ledger, 'Alpha changed.\n\nGamma.');
    const revision = editor.surface!.session.packageRevision();
    const playback = playDirectChanges(editor, ledger, { duration: 3000, reducedMotion: false });
    playback.cancel();
    await playback.finished;
    expect(overlays()).toBe(0);
    expect(editor.surface!.session.packageRevision()).toBe(revision);
    expect(texts(editor)).toEqual(['Alpha changed.', 'Gamma.']);
    expect(highlighted(editor, ledger)).toEqual(['Alpha changed.']);

    // Reduced motion reveals at once without overlays.
    await commit(editor, ledger, 'Alpha changed.\n\nGamma changed.');
    const still = playDirectChanges(editor, ledger, { reducedMotion: true });
    expect(still.motion).toBe(false);
    expect(overlays()).toBe(0);
    expect(highlighted(editor, ledger)).toEqual(['Alpha changed.', 'Gamma changed.']);

    // An abort signal stops playback; undo during playback removes the change it touched.
    await commit(editor, ledger, 'Alpha again.\n\nGamma changed.');
    const controller = new AbortController(),
      running = playDirectChanges(editor, ledger, {
        duration: 3000,
        reducedMotion: false,
        signal: controller.signal,
      });
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    controller.abort();
    await running.finished;
    expect(overlays()).toBe(0);
    expect(texts(editor)).toEqual(['Alpha changed.', 'Gamma changed.']);
    expect(highlighted(editor, ledger)).not.toContain('Alpha again.');
    expect(ledger.changes.every((change) => !change.pending)).toBe(true);
  } finally {
    ledger.dispose();
    editor.destroy();
  }
}, 30_000);

for (const first of [0, 1] as const)
  test(`two sessions: a peer edit and deletion during playback converge without stale highlights; receiver ${first}`, async () => {
    const pair = await createTwoEditorFixture({
      document: await bytes(),
      documentId: `playback-${first}`,
      clientIds: [73, 41],
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
        ledger = new DirectChangeLedger(alice);
      await commit(alice, ledger, 'Alpha changed.\n\nBeta changed.\n\nGamma.');
      const playback = playDirectChanges(alice, ledger, { duration: 1500, reducedMotion: false });
      pair.pause();
      const part = bob.surface!.session.part(),
        ids = bob.surface!.session.paragraphIds(),
        gamma = ids.find((id) => paragraphTextOf(part, id) === 'Gamma.')!;
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: gamma, offset: 6, text: ' Bob' },
        ]).committed
      ).toBe(true);
      pair.reconnect(first);
      await pair.settle();
      const beta = alice
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(alice.surface!.session.part(), id) === 'Beta changed.')!;
      expect(
        bob.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: beta }]).committed
      ).toBe(true);
      await pair.settle();
      await playback.finished;
      expect(texts(alice)).toEqual(texts(bob));
      expect(texts(alice)).toEqual(['Alpha changed.', 'Gamma. Bob']);
      expect(overlays()).toBe(0);
      expect(highlighted(alice, ledger)).toEqual(['Alpha changed.']);
      // Alice's undo keeps the peer's work.
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(texts(alice)).toEqual(texts(bob));
      expect(texts(alice)).toContain('Gamma. Bob');
      expect(texts(alice)).toContain('Alpha.');
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      ledger.dispose();
    } finally {
      pair.destroy();
    }
  }, 60_000);
