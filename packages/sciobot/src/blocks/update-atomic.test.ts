import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { previewDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from '../diff/plan-target';
import { buildBlockIndex } from './block-index';
import { readPayloadStore } from './payload-opc-codec';
import { updateAtomicBlock } from './update-atomic';

const ORIGINAL = 'def f():\n\treturn 1';
const CHANGED = 'def f():\n\tif  x:\n\t\treturn  2\n\n\treturn 3';

async function codeDocument() {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Seed text</w:t></w:r></w:p>');
  const planned = planTargetDocument(input, `\`\`\`python\n${ORIGINAL}\n\`\`\`\n\nSeed text`, {
    allocateBlockId: () => 'code-1',
  });
  if (!planned.success) throw new Error(planned.error);
  const applied = previewDirectPlan(input.pkg, planned.plan);
  if (!applied.success) throw new Error(applied.error);
  return writeOoxmlPackage(applied.pkg);
}

const code = (pkg: OoxmlPackage) => buildBlockIndex(pkg).byId.get('code-1')?.payload.attributes;
const texts = (editor: DocxEditorInstance) =>
  editor
    .surface!.session.paragraphIds()
    .map((id) => paragraphTextOf(editor.surface!.session.part(), id));

test('an atomic code block update keeps its ID, exact whitespace, Word lines, and history', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: await codeDocument(), author: 'Ann' });
  try {
    const surface = editor.surface!;
    expect(
      updateAtomicBlock(editor, {
        blockId: 'code-1',
        attributes: { language: 'python', value: CHANGED },
      })
    ).toMatchObject({ success: true });
    expect(code(surface.session.currentPackage())).toEqual({ language: 'python', value: CHANGED });
    expect(texts(editor)).toEqual([...CHANGED.split('\n'), 'Seed text']);
    expect(
      updateAtomicBlock(editor, {
        blockId: 'code-1',
        attributes: { language: 'typescript', value: CHANGED },
      }).success
    ).toBe(true);

    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(code(reopened.package)).toEqual({ language: 'typescript', value: CHANGED });
    const store = readPayloadStore(reopened.package);
    expect(store.status === 'ready' && [...store.entries.keys()]).toEqual(['code-1']);
    expect((await projectSourcePackage(reopened.package)).text).toBe(
      `\`\`\`typescript\n${CHANGED}\n\`\`\`\n\nSeed text`
    );

    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(code(surface.session.currentPackage())?.language).toBe('python');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(code(surface.session.currentPackage())).toEqual({ language: 'python', value: ORIGINAL });
    expect(texts(editor)).toEqual([...ORIGINAL.split('\n'), 'Seed text']);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(code(surface.session.currentPackage())?.value).toBe(CHANGED);

    expect(
      updateAtomicBlock(editor, { blockId: 'code-1', attributes: { language: 1 } })
    ).toMatchObject({ success: false, code: 'invalid-input' });
    expect(updateAtomicBlock(editor, { blockId: 'missing', attributes: {} })).toMatchObject({
      success: false,
      code: 'invalid-input',
    });
  } finally {
    editor.destroy();
    container.remove();
  }
});

test('an inline atomic control refuses a paragraph fallback update', async () => {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Seed text</w:t></w:r></w:p>');
  const planned = planTargetDocument(input, 'Area $a^2$ here.\n\nSeed text', {
    allocateBlockId: () => 'math-1',
  });
  if (!planned.success) throw new Error(planned.error);
  const applied = previewDirectPlan(input.pkg, planned.plan);
  if (!applied.success) throw new Error(applied.error);
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: writeOoxmlPackage(applied.pkg),
    author: 'Ann',
  });
  try {
    const before = editor.surface!.session.packageRevision();
    expect(
      updateAtomicBlock(editor, { blockId: 'math-1', attributes: { texExpression: 'b^2' } })
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(editor.surface!.session.packageRevision()).toBe(before);
  } finally {
    editor.destroy();
    container.remove();
  }
});

for (const first of [0, 1] as const)
  test(`two sessions converge on a code update beside peer typing; receiver ${first}`, async () => {
    const pair = await createTwoEditorFixture({ document: await codeDocument() });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const seed = bob
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Seed text')!;
      pair.pause();
      expect(
        updateAtomicBlock(alice, {
          blockId: 'code-1',
          attributes: { language: 'python', value: CHANGED },
        }).success
      ).toBe(true);
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: seed, offset: 9, text: ' Bob' },
        ]).committed
      ).toBe(true);
      pair.reconnect(first);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(code(editor.surface!.session.currentPackage())?.value).toBe(CHANGED);
        expect(texts(editor)).toEqual([...CHANGED.split('\n'), 'Seed text Bob']);
      }
      // Alice's undo restores the old code and keeps Bob's text; redo restores her edit.
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(code(editor.surface!.session.currentPackage())?.value).toBe(ORIGINAL);
        expect(texts(editor)).toEqual([...ORIGINAL.split('\n'), 'Seed text Bob']);
      }
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(code(editor.surface!.session.currentPackage())?.value).toBe(CHANGED);
      // Bob removes the block while Alice changes it: both converge.
      pair.pause();
      expect(
        updateAtomicBlock(alice, {
          blockId: 'code-1',
          attributes: { language: 'python', value: 'x = 1' },
        }).success
      ).toBe(true);
      const control = buildBlockIndex(bob.surface!.session.currentPackage()).byId.get('code-1')!;
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'removeContentControl', controlId: control.controlId, keepContent: false },
        ]).committed
      ).toBe(true);
      pair.reconnect(first);
      await pair.settle();
      expect(texts(alice)).toEqual(texts(bob));
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
