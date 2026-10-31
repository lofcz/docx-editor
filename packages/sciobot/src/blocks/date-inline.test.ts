import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { paragraphTextOf, readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { projectSourcePackage, sourceRevision } from '../../test/support/target-source-fixture';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { previewDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from '../diff/plan-target';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { buildBlockIndex } from './block-index';
import { insertInlineComponent } from './inline-insert';

const MARKDOWN = 'Meeting on <date value="2024-05-01" /> in class';

async function dated() {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Seed</w:t></w:r></w:p>');
  const planned = planTargetDocument(input, MARKDOWN, { allocateBlockId: () => 'date-1' });
  if (!planned.success) throw new Error(planned.error);
  const applied = previewDirectPlan(input.pkg, planned.plan);
  if (!applied.success) throw new Error(applied.error);
  return writeOoxmlPackage(applied.pkg);
}

async function markdownOf(editor: DocxEditorInstance) {
  const view = await readWorkingView(
    editor.surface!.session,
    readEditorRevision(editor, 'date', 'date')!
  );
  if (!view.success) throw new Error(view.error);
  return view.view.projection.text;
}

const texts = (editor: DocxEditorInstance) =>
  editor
    .surface!.session.paragraphIds()
    .map((id) => paragraphTextOf(editor.surface!.session.part(), id));

test('a date chip keeps its ISO value, shows it in Word, and updates in place', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: await dated() });
  try {
    expect(await markdownOf(editor)).toBe(MARKDOWN);
    expect(texts(editor)).toEqual(['Meeting on 2024-05-01 in class']);
    const index = buildBlockIndex(editor.surface!.session.currentPackage());
    expect(index.byId.get('date-1')?.payload).toMatchObject({
      type: 'date',
      attributes: { value: '2024-05-01' },
    });
    // An agent change of the value plans an in-place update, not a new chip.
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    const source = {
      revision: sourceRevision,
      pkg: reopened.package,
      projection: await projectSourcePackage(reopened.package),
    };
    expect(source.projection.text).toBe(MARKDOWN);
    const changed = planTargetDocument(source, MARKDOWN.replace('2024-05-01', '2024-06-15'), {
      allocateBlockId: () => 'unused',
    });
    if (!changed.success) throw new Error(changed.error);
    const applied = previewDirectPlan(source.pkg, changed.plan);
    if (!applied.success) throw new Error(applied.error);
    expect(buildBlockIndex(applied.pkg).byId.get('date-1')?.payload.attributes).toEqual({
      value: '2024-06-15',
    });
    // Invalid values refuse.
    const paragraph = editor.surface!.session.paragraphIds()[0]!;
    expect(
      insertInlineComponent(editor, {
        payload: { version: 1, id: 'bad', type: 'date', attributes: { value: '15. 6. 2024' } },
        paragraphId: paragraph,
        offset: 0,
      })
    ).toMatchObject({ success: false, code: 'invalid-input' });
  } finally {
    editor.destroy();
    container.remove();
  }
});

test('inserting a date chip is one undo step and survives reopen', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: await dated() });
  try {
    const paragraph = editor.surface!.session.paragraphIds()[0]!;
    const inserted = insertInlineComponent(editor, {
      payload: { version: 1, id: 'date-2', type: 'date', attributes: { value: '2025-01-02' } },
      paragraphId: paragraph,
      offset: 0,
    });
    if (!inserted.success) throw new Error(inserted.error);
    expect(await markdownOf(editor)).toBe(`<date value="2025-01-02" />${MARKDOWN}`);
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(buildBlockIndex(reopened.package).byId.get('date-2')?.payload.attributes).toEqual({
      value: '2025-01-02',
    });
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(await markdownOf(editor)).toBe(MARKDOWN);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(await markdownOf(editor)).toBe(`<date value="2025-01-02" />${MARKDOWN}`);
  } finally {
    editor.destroy();
    container.remove();
  }
});

for (const first of [0, 1] as const)
  test(`a date chip converges beside peer typing; receiver ${first}`, async () => {
    const pair = await createTwoEditorFixture({ document: await dated() });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      pair.pause();
      const paragraph = alice.surface!.session.paragraphIds()[0]!;
      expect(
        insertInlineComponent(alice, {
          payload: { version: 1, id: 'date-3', type: 'date', attributes: { value: '2025-03-04' } },
          paragraphId: paragraph,
          offset: 0,
        }).success
      ).toBe(true);
      const bobParagraph = bob.surface!.session.paragraphIds()[0]!;
      const length = paragraphTextOf(bob.surface!.session.part(), bobParagraph)!.length;
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: bobParagraph, offset: length, text: ' Bob' },
        ]).committed
      ).toBe(true);
      pair.reconnect(first);
      await pair.settle();
      for (const editor of pair.editors)
        expect(await markdownOf(editor)).toBe(`<date value="2025-03-04" />${MARKDOWN} Bob`);
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) expect(await markdownOf(editor)).toBe(`${MARKDOWN} Bob`);
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
