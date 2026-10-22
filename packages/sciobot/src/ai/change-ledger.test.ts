import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { paragraphTextOf, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { C01_SOURCES } from '../../test/support/c01-sources';
import { protectedTargetInput, targetParagraph as p } from '../../test/support/target-document';
import { planTargetDocument } from '../diff/plan-target';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readWorkingView } from '../review/working-view';
import { applyDirectPlan } from './apply-direct';
import { DirectChangeLedger } from './change-ledger';

const context = { documentId: 'ledger', epoch: 'mount' };
const LESSON =
  '<lesson_title grade="7">Lesson</lesson_title>\n\n<phase type="evocation">\n<activity name="Warmup" duration="10">\n{first}\n\n<source title="Source" url="https://example.com" type="verified" />\n\n{last}\n</activity>\n</phase>';
const lesson = (first: string, last: string) =>
  LESSON.replace('{first}', first).replace('{last}', last);

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
async function direct(
  editor: DocxEditorInstance,
  ledger: DirectChangeLedger,
  target: string,
  session: string,
  pending = false
) {
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
  ledger.record(before, { session, pending });
}
const textOf = (editor: DocxEditorInstance, id: string) =>
  paragraphTextOf(editor.surface!.session.part(), id);
const lessonEditor = async () => {
  const [body, options] = C01_SOURCES['nested MDX and attributes']!();
  return mount(writeOoxmlPackage((await protectedTargetInput(body, options)).pkg));
};

test('two direct edits to one activity in one request form one change', async () => {
  const editor = await lessonEditor(),
    ledger = new DirectChangeLedger(editor);
  try {
    await direct(editor, ledger, lesson('A better question.', 'Last question.'), 'request-1');
    expect(ledger.changes).toHaveLength(1);
    const first = ledger.changes[0]!;
    expect(first).toMatchObject({ ownerBlockId: 'activity', kind: 'text', status: 'new' });
    expect(first.nodeIds.map((id) => textOf(editor, id))).toEqual(['A better question.']);
    await direct(editor, ledger, lesson('A better question.', 'Final reflection.'), 'request-1');
    expect(ledger.changes).toHaveLength(1);
    const merged = ledger.changes[0]!;
    expect(merged.id).toBe(first.id);
    expect(merged.ownerBlockId).toBe('activity');
    expect(merged.nodeIds.map((id) => textOf(editor, id))).toEqual([
      'A better question.',
      'Final reflection.',
    ]);
    // Repeating an edit of the same paragraph keeps one change.
    await direct(
      editor,
      ledger,
      lesson('An even better question.', 'Final reflection.'),
      'request-1'
    );
    expect(ledger.changes.map((change) => change.id)).toEqual([first.id]);
    expect(ledger.changes[0]!.nodeIds.map((id) => textOf(editor, id))).toEqual([
      'An even better question.',
      'Final reflection.',
    ]);
  } finally {
    ledger.dispose();
    editor.destroy();
  }
});

test('a new request settles earlier highlights without deciding unrelated Word review items', async () => {
  const body =
    p('Keep.') +
    '<w:p><w:ins w:id="7" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Imported proposal.</w:t></w:r></w:ins></w:p>' +
    p('Tail.');
  const editor = mount(writeOoxmlPackage((await protectedTargetInput(body)).pkg)),
    ledger = new DirectChangeLedger(editor);
  try {
    const review = editor.getReviewItems().map((item) => item.key);
    expect(review).toHaveLength(1);
    await direct(editor, ledger, 'Kept.\n\nImported proposal.\n\nTail.', 'request-1');
    expect(ledger.changes.map((change) => change.nodeIds.map((id) => textOf(editor, id)))).toEqual([
      ['Kept.'],
    ]);
    await direct(editor, ledger, 'Kept.\n\nImported proposal.\n\nTail end.', 'request-2');
    expect(ledger.changes.map((change) => change.nodeIds.map((id) => textOf(editor, id)))).toEqual([
      ['Tail end.'],
    ]);
    expect(editor.getReviewItems().map((item) => item.key)).toEqual(review);
    ledger.settleAll();
    expect(ledger.changes).toEqual([]);
    expect(editor.getReviewItems().map((item) => item.key)).toEqual(review);
    expect(textOf(editor, editor.surface!.session.paragraphIds()[0]!)).toBe('Kept.');
  } finally {
    ledger.dispose();
    editor.destroy();
  }
});

test('insertions, removals, pending playback, acknowledgement, and fade settle correctly', async () => {
  const editor = mount(writeOoxmlPackage((await protectedTargetInput(p('One.') + p('Two.'))).pkg)),
    ledger = new DirectChangeLedger(editor, 5);
  try {
    await direct(editor, ledger, 'One.', 'request-1');
    // Removals leave nothing to highlight.
    expect(ledger.changes).toEqual([]);
    await direct(editor, ledger, 'One.\n\nNew one.\n\nNew two.', 'request-2', true);
    expect(ledger.changes).toHaveLength(2);
    expect(ledger.changes.every((change) => change.kind === 'insert' && change.pending)).toBe(true);
    ledger.reveal([ledger.changes[0]!.nodeIds[0]!]);
    expect(ledger.changes.map((change) => change.pending)).toEqual([false, true]);
    ledger.reveal();
    expect(ledger.changes.some((change) => change.pending)).toBe(false);
    const [first, second] = ledger.changes;
    ledger.acknowledge(first!.id);
    expect(ledger.changes.find((change) => change.id === first!.id)!.status).toBe('seen');
    await new Promise((done) => setTimeout(done, 30));
    expect(ledger.changes.map((change) => change.id)).toEqual([second!.id]);
    // Settled content stays in the document and in the baseline.
    expect(editor.surface!.session.paragraphIds().map((id) => textOf(editor, id))).toEqual([
      'One.',
      'New one.',
      'New two.',
    ]);
    await direct(editor, ledger, 'One.\n\nNew one.\n\nNew two!', 'request-2');
    expect(ledger.changes).toHaveLength(1);
    expect(ledger.changes[0]!.nodeIds.map((id) => textOf(editor, id))).toEqual(['New two!']);
  } finally {
    ledger.dispose();
    editor.destroy();
  }
});

test('user typing, undo, and remote-style commits seal the baseline and drop only touched changes', async () => {
  const editor = mount(
      writeOoxmlPackage((await protectedTargetInput(p('Alpha.') + p('Beta.') + p('Gamma.'))).pkg)
    ),
    ledger = new DirectChangeLedger(editor);
  try {
    await direct(editor, ledger, 'Alpha!\n\nBeta.\n\nGamma!', 'request-1');
    expect(ledger.changes).toHaveLength(2);
    const [alpha, gamma] = ledger.changes;
    const typed = alpha!.nodeIds[0]!;
    expect(
      editor.surface!.applyAutomationOps(() => [
        { op: 'insertText', paragraphId: typed, offset: 0, text: 'User ' },
      ]).committed
    ).toBe(true);
    expect(ledger.changes.map((change) => change.id)).toEqual([gamma!.id]);
    // The manual sentence is kept; a later direct edit in the same request starts a new baseline.
    await direct(editor, ledger, 'User Alpha!\n\nBeta?\n\nGamma!', 'request-1');
    expect(ledger.changes.map((change) => change.nodeIds.map((id) => textOf(editor, id)))).toEqual([
      ['Beta?'],
    ]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(ledger.changes).toEqual([]);
    expect(textOf(editor, typed)).toBe('User Alpha!');
  } finally {
    ledger.dispose();
    editor.destroy();
  }
});
