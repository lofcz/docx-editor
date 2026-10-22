import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { paragraphTextOf, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { protectedTargetInput, targetParagraph as p } from '../../test/support/target-document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { planTargetDocument } from '../diff/plan-target';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { applyDirectPlan } from './apply-direct';
import { DirectTargetBaseline } from './direct-baseline';

const context = { documentId: 'baseline', epoch: 'mount' };
const texts = (editor: DocxEditorInstance) =>
  editor
    .surface!.session.paragraphIds()
    .map((id) => paragraphTextOf(editor.surface!.session.part(), id));
function mount(bytes: Uint8Array) {
  const container = document.createElement('div');
  document.body.append(container);
  return createDocxEditor({
    container,
    document: bytes,
    author: 'Teacher',
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
}
/** One agent call: prepare, plan, apply, and remember the landing. */
async function call(editor: DocxEditorInstance, baseline: DirectTargetBaseline, target: string) {
  const prepared = await baseline.prepare(editor, context);
  if (!prepared.success) throw new Error(prepared.error);
  const planned = planTargetDocument(prepared.input, target);
  if (!planned.success) throw new Error(planned.error);
  const applied = applyDirectPlan(editor, planned.plan, context);
  if (!applied.success) throw new Error(`${applied.code}: ${applied.error}`);
  const landed = await baseline.remember(editor, context);
  if (!landed.success) throw new Error(landed.error);
  return { prepared, landed };
}
const OPENING = 'Opening.\n\nMiddle.\n\nClosing.';
const INTERMEDIATE = 'Opening.\n\nMiddle, revised.\n\nClosing.';
const SECOND = 'Opening.\n\nMiddle, revised.\n\nClosing, revised.';
const document_ = async (markdown: string) =>
  writeOoxmlPackage(
    (
      await protectedTargetInput(
        markdown
          .split('\n\n')
          .map((line) => p(line))
          .join('')
      )
    ).pkg
  );

test('consecutive direct calls reuse the landed source until the document changes', async () => {
  const editor = mount(await document_(OPENING)),
    baseline = new DirectTargetBaseline();
  try {
    const first = await call(editor, baseline, INTERMEDIATE);
    expect(first.prepared.reused).toBe(false);
    const cached = await baseline.prepare(editor, context);
    if (!cached.success) throw new Error(cached.error);
    expect(cached.reused).toBe(true);
    expect(cached.input.projection).toBe(first.landed.input.projection);
    // The reused source equals a fresh read of a document that holds the intermediate target.
    const fresh = mount(await document_(INTERMEDIATE));
    try {
      const reread = await new DirectTargetBaseline().prepare(fresh, context);
      if (!reread.success) throw new Error(reread.error);
      expect(cached.input.projection.text).toBe(reread.input.projection.text);
    } finally {
      fresh.destroy();
    }
    const second = await call(editor, baseline, SECOND);
    expect(second.prepared.reused).toBe(true);
    expect(texts(editor)).toEqual(['Opening.', 'Middle, revised.', 'Closing, revised.']);

    // A manual sentence invalidates the landing and survives the next direct edit.
    const middle = editor.surface!.session.paragraphIds()[1]!;
    expect(
      editor.surface!.applyAutomationOps(() => [
        { op: 'insertText', paragraphId: middle, offset: 16, text: ' Manual sentence.' },
      ]).committed
    ).toBe(true);
    const third = await call(
      editor,
      baseline,
      'Opening, revised.\n\nMiddle, revised. Manual sentence.\n\nClosing, revised.'
    );
    expect(third.prepared.reused).toBe(false);
    expect(texts(editor)).toEqual([
      'Opening, revised.',
      'Middle, revised. Manual sentence.',
      'Closing, revised.',
    ]);

    // Undo and redo each invalidate; save/reopen keeps the landed text.
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    const afterUndo = await baseline.prepare(editor, context);
    expect(afterUndo.success && afterUndo.reused).toBe(false);
    expect(texts(editor)[1]).toBe('Middle, revised. Manual sentence.');
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const afterRedo = await baseline.prepare(editor, context);
    expect(afterRedo.success && afterRedo.reused).toBe(false);
    const reopened = mount(new Uint8Array(await editor.save()));
    try {
      expect(texts(reopened)).toEqual(texts(editor));
    } finally {
      reopened.destroy();
    }
  } finally {
    editor.destroy();
  }
}, 60_000);

test('a review decision invalidates the remembered landing', async () => {
  const body =
    p('Keep.') +
    '<w:p><w:ins w:id="7" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Proposal.</w:t></w:r></w:ins></w:p>';
  const editor = mount(writeOoxmlPackage((await protectedTargetInput(body)).pkg)),
    baseline = new DirectTargetBaseline();
  try {
    await call(editor, baseline, 'Kept.\n\nProposal.');
    expect(editor.acceptReviewItem(editor.getReviewItems()[0]!.key).ok).toBe(true);
    const prepared = await baseline.prepare(editor, context);
    expect(prepared.success && prepared.reused).toBe(false);
  } finally {
    editor.destroy();
  }
}, 60_000);

for (const first of [0, 1] as const)
  test(`a remote revision invalidates the landing and peer text survives the next call; receiver ${first}`, async () => {
    const pair = await createTwoEditorFixture({ document: await document_(OPENING) });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
        baseline = new DirectTargetBaseline();
      await call(alice, baseline, INTERMEDIATE);
      await pair.settle();
      const closing = bob.surface!.session.paragraphIds()[2]!;
      pair.pause();
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: closing, offset: 8, text: ' Bob.' },
        ]).committed
      ).toBe(true);
      pair.reconnect(first);
      await pair.settle();
      const prepared = await baseline.prepare(alice, context);
      expect(prepared.success && prepared.reused).toBe(false);
      await call(alice, baseline, 'Opening, revised.\n\nMiddle, revised.\n\nClosing. Bob.');
      await pair.settle();
      for (const editor of pair.editors)
        expect(texts(editor)).toEqual(['Opening, revised.', 'Middle, revised.', 'Closing. Bob.']);
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(texts(editor)).toEqual(['Opening.', 'Middle, revised.', 'Closing. Bob.']);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const [savedAlice, savedBob] = await pair.savedDigests();
      expect(savedAlice).toEqual(savedBob);
    } finally {
      pair.destroy();
    }
  }, 60_000);
