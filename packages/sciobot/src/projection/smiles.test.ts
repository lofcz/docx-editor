import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { paragraphTextOf, readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { projectSourcePackage, sourceRevision } from '../../test/support/target-source-fixture';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { previewDirectPlan } from '../ai/apply-direct';
import { buildBlockIndex } from '../blocks/block-index';
import { updateAtomicBlock } from '../blocks/update-atomic';
import { planTargetDocument } from '../diff/plan-target';

// `[C@@H](C)` reads as a Markdown link in element text; the attribute keeps it exact.
const ALANINE = 'N[C@@H](C)C(=O)O';
const ASPIRIN = 'CC(=O)OC1=CC=CC=C1C(=O)O';
const MARKDOWN = `Intro\n\n<smiles smiles="${ALANINE}" />\n\nAfter`;

async function lesson(markdown = MARKDOWN) {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Seed</w:t></w:r></w:p>');
  const planned = planTargetDocument(input, markdown, { allocateBlockId: () => 'smiles-1' });
  if (!planned.success) throw new Error(planned.error);
  const applied = previewDirectPlan(input.pkg, planned.plan);
  if (!applied.success) throw new Error(applied.error);
  return applied.pkg;
}
const texts = (editor: DocxEditorInstance) => {
  const session = editor.surface!.session;
  return session.paragraphIds().map((id) => paragraphTextOf(session.part(), id));
};
async function projected(editor: DocxEditorInstance) {
  const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
  if (!reopened.ok) throw new Error(reopened.reason);
  return { pkg: reopened.package, projection: await projectSourcePackage(reopened.package) };
}

test('a SMILES block keeps its exact string through save, reopen, and an unchanged plan', async () => {
  const pkg = await lesson();
  expect(buildBlockIndex(pkg).byId.get('smiles-1')?.payload).toMatchObject({
    type: 'smiles',
    attributes: { smiles: ALANINE },
  });
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(pkg),
  });
  try {
    // Word shows the exact source.
    expect(texts(editor)).toEqual(['Intro', ALANINE, 'After']);
    const { pkg: reopened, projection } = await projected(editor);
    expect(projection.text).toBe(MARKDOWN);
    const again = planTargetDocument(
      { revision: sourceRevision, pkg: reopened, projection },
      MARKDOWN,
      { allocateBlockId: () => 'unused' }
    );
    if (!again.success) throw new Error(again.error);
    expect(again.plan.operations).toHaveLength(0);
  } finally {
    editor.destroy();
  }
});

test('the element-text form reads as the same structure', async () => {
  const pkg = await lesson(`Intro\n\n<smiles>\n${ASPIRIN}\n</smiles>\n\nAfter`);
  expect(buildBlockIndex(pkg).byId.get('smiles-1')?.payload.attributes).toEqual({
    smiles: ASPIRIN,
  });
});

test('an edit changes source and Word fallback together; undo restores both', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(await lesson()),
  });
  try {
    const updated = updateAtomicBlock(editor, {
      blockId: 'smiles-1',
      attributes: { smiles: ASPIRIN },
    });
    if (!updated.success) throw new Error(updated.error);
    expect(texts(editor)).toEqual(['Intro', ASPIRIN, 'After']);
    expect((await projected(editor)).projection.text).toContain(`<smiles smiles="${ASPIRIN}" />`);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(editor)).toEqual(['Intro', ALANINE, 'After']);
    expect((await projected(editor)).projection.text).toBe(MARKDOWN);
    // A string Word cannot carry exactly is refused.
    expect(
      updateAtomicBlock(editor, { blockId: 'smiles-1', attributes: { smiles: 'C C' } })
    ).toMatchObject({ success: false });
  } finally {
    editor.destroy();
  }
});

for (const first of [0, 1] as const)
  test(`a SMILES edit converges beside peer typing; receiver ${first}`, async () => {
    const pair = await createTwoEditorFixture({ document: writeOoxmlPackage(await lesson()) });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      pair.pause();
      expect(
        updateAtomicBlock(alice, { blockId: 'smiles-1', attributes: { smiles: ASPIRIN } }).success
      ).toBe(true);
      const after = bob
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'After')!;
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: after, offset: 5, text: ' Bob' },
        ]).committed
      ).toBe(true);
      pair.reconnect(first);
      await pair.settle();
      for (const editor of pair.editors)
        expect(texts(editor)).toEqual(['Intro', ASPIRIN, 'After Bob']);
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(texts(editor)).toEqual(['Intro', ALANINE, 'After Bob']);
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
