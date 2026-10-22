import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { findNode, paragraphTextOf, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { protectedTargetInput, targetParagraph as p } from '../../test/support/target-document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { planTargetDocument } from '../diff/plan-target';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { directHighlights } from './acknowledge';
import { applyDirectPlan } from './apply-direct';
import { DirectChangeLedger } from './change-ledger';

const context = { documentId: 'ack', epoch: 'mount' };
const bytes = async () =>
  writeOoxmlPackage((await protectedTargetInput(p('Alpha.') + p('Beta.') + p('Gamma.'))).pkg);
async function direct(editor: DocxEditorInstance, ledger: DirectChangeLedger, target: string) {
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
  ledger.record(before, { session: 'request' });
}
/** Every painted range addresses a live paragraph within its current text. */
function assertLive(editor: DocxEditorInstance, ledger: DirectChangeLedger) {
  const part = editor.surface!.session.part();
  for (const highlight of directHighlights(editor, ledger)) {
    expect(findNode(part, highlight.paragraphId)?.kind).toBe('paragraph');
    expect(highlight.end).toBe(paragraphTextOf(part, highlight.paragraphId)!.length);
  }
}
const texts = (editor: DocxEditorInstance, ledger: DirectChangeLedger) =>
  directHighlights(editor, ledger).map((highlight) =>
    paragraphTextOf(editor.surface!.session.part(), highlight.paragraphId)
  );

test('typing in one highlighted block settles only that change; undo and deletion leave no dangling range', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: await bytes() }),
    ledger = new DirectChangeLedger(editor);
  try {
    await direct(editor, ledger, 'Alpha!\n\nBeta.\n\nGamma!');
    expect(texts(editor, ledger)).toEqual(['Alpha!', 'Gamma!']);
    const [alpha, gamma] = directHighlights(editor, ledger);
    editor.surface!.setSelection({
      anchor: { paragraphId: alpha!.paragraphId, offset: 6 },
      head: { paragraphId: alpha!.paragraphId, offset: 6 },
    });
    editor.surface!.type(' typed');
    // Painting flushes any queued typing first; the typed commit settles the touched change.
    expect(texts(editor, ledger)).toEqual(['Gamma!']);
    expect(paragraphTextOf(editor.surface!.session.part(), alpha!.paragraphId)).toBe(
      'Alpha! typed'
    );
    expect(directHighlights(editor, ledger)[0]!.changeId).toBe(gamma!.changeId);
    assertLive(editor, ledger);
    // Deleting the remaining highlighted paragraph drops its change.
    expect(
      editor.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: gamma!.paragraphId }])
        .committed
    ).toBe(true);
    expect(directHighlights(editor, ledger)).toEqual([]);
    // Undo of the deletion does not resurrect a stale highlight; undo of the AI edit leaves none.
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    assertLive(editor, ledger);
    expect(directHighlights(editor, ledger)).toEqual([]);
    await direct(editor, ledger, 'Alpha! typed\n\nBeta?\n\nGamma!');
    expect(texts(editor, ledger)).toEqual(['Beta?']);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(directHighlights(editor, ledger)).toEqual([]);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    assertLive(editor, ledger);
    // Highlights are presentation state: a reopened document has none and keeps the text.
    const reopenedContainer = document.createElement('div');
    document.body.append(reopenedContainer);
    const reopened = createDocxEditor({
      container: reopenedContainer,
      document: new Uint8Array(await editor.save()),
    });
    const fresh = new DirectChangeLedger(reopened);
    try {
      expect(directHighlights(reopened, fresh)).toEqual([]);
      expect(
        reopened
          .surface!.session.paragraphIds()
          .map((id) => paragraphTextOf(reopened.surface!.session.part(), id))
      ).toEqual(['Alpha! typed', 'Beta?', 'Gamma!']);
    } finally {
      fresh.dispose();
      reopened.destroy();
    }
  } finally {
    ledger.dispose();
    editor.destroy();
  }
}, 60_000);

for (const first of [0, 1] as const)
  test(`a remote edit inside a highlighted block settles it on the author's peer; receiver ${first}`, async () => {
    const pair = await createTwoEditorFixture({ document: await bytes() });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
        ledger = new DirectChangeLedger(alice);
      await direct(alice, ledger, 'Alpha!\n\nBeta.\n\nGamma!');
      await pair.settle();
      const [alpha, gamma] = directHighlights(alice, ledger);
      pair.pause();
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: alpha!.paragraphId, offset: 0, text: 'Bob ' },
        ]).committed
      ).toBe(true);
      pair.reconnect(first);
      await pair.settle();
      expect(directHighlights(alice, ledger).map((highlight) => highlight.changeId)).toEqual([
        gamma!.changeId,
      ]);
      assertLive(alice, ledger);
      // Remote deletion of the remaining highlighted paragraph leaves no dangling range.
      expect(
        bob.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: gamma!.paragraphId }])
          .committed
      ).toBe(true);
      await pair.settle();
      expect(directHighlights(alice, ledger)).toEqual([]);
      const [savedAlice, savedBob] = await pair.savedDigests();
      expect(savedAlice).toEqual(savedBob);
      ledger.dispose();
    } finally {
      pair.destroy();
    }
  }, 60_000);
