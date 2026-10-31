import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { paragraphTextOf, readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { previewDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from '../diff/plan-target';
import { updateAtomicBlock } from './update-atomic';

const visual = (src: string, fileId: string) =>
  `<lesson_visual kind="explainer" phase="understanding" src="${src}" alt="Leaf diagram" prompt="A leaf" fileId="${fileId}" />`;
const MARKDOWN = `Intro\n\n${visual('visuals/a.png', 'f1')}\n\nAfter`;

async function markdown(editor: DocxEditorInstance) {
  const read = readOoxmlPackage(new Uint8Array(await editor.save()));
  if (!read.ok) throw new Error(read.reason);
  return (await projectSourcePackage(read.package)).text;
}

test('an atomic block without its own fallback updates with the default Word line', async () => {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Seed</w:t></w:r></w:p>');
  const planned = planTargetDocument(input, MARKDOWN, { allocateBlockId: () => 'lv-1' });
  if (!planned.success) throw new Error(planned.error);
  const applied = previewDirectPlan(input.pkg, planned.plan);
  if (!applied.success) throw new Error(applied.error);
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(applied.pkg),
  });
  try {
    const texts = () =>
      editor
        .surface!.session.paragraphIds()
        .map((id) => paragraphTextOf(editor.surface!.session.part(), id));
    expect(texts()).toEqual(['Intro', 'Leaf diagram', 'After']);
    const updated = updateAtomicBlock(editor, {
      blockId: 'lv-1',
      attributes: {
        kind: 'explainer',
        phase: 'understanding',
        src: 'visuals/b.png',
        alt: 'Leaf diagram',
        prompt: 'A leaf, red veins',
        fileId: 'f2',
      },
    });
    if (!updated.success) throw new Error(updated.error);
    expect(texts()).toEqual(['Intro', 'Leaf diagram', 'After']);
    expect(await markdown(editor)).toContain('src="visuals/b.png"');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(await markdown(editor)).toBe(MARKDOWN);
  } finally {
    editor.destroy();
  }
});
