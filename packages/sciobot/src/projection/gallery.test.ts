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

const ONE = { url: 'https://img.example/one.png', caption: 'One' };
const TWO = { url: 'https://img.example/two.png' };
const tag = (images: object[]) =>
  `<image_gallery layout="row-fit" images={${JSON.stringify(images)}} />`;
const MARKDOWN = `Intro\n\n${tag([ONE, TWO])}\n\nAfter`;

async function lesson(markdown = MARKDOWN) {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Seed</w:t></w:r></w:p>');
  const planned = planTargetDocument(input, markdown, { allocateBlockId: () => 'gallery-1' });
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
  const read = readOoxmlPackage(new Uint8Array(await editor.save()));
  if (!read.ok) throw new Error(read.reason);
  return { pkg: read.package, projection: await projectSourcePackage(read.package) };
}

test('a gallery keeps its images and order through save, reopen, and an unchanged plan', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(await lesson()),
  });
  try {
    // Word shows one line per image: its caption, or its address.
    expect(texts(editor)).toEqual(['Intro', 'One', TWO.url, 'After']);
    const { pkg, projection } = await projected(editor);
    expect(projection.text).toBe(MARKDOWN);
    for (const target of [MARKDOWN, `Intro\n\n![One](${ONE.url})\n![](${TWO.url})\n\nAfter`]) {
      const again = planTargetDocument({ revision: sourceRevision, pkg, projection }, target, {
        allocateBlockId: () => 'unused',
      });
      if (!again.success) throw new Error(again.error);
      expect(again.plan.operations).toHaveLength(0);
    }
  } finally {
    editor.destroy();
  }
});

test('Markdown images become a gallery; unsafe addresses refuse', async () => {
  const pkg = await lesson(`Intro\n\n![One](${ONE.url})\n\nAfter`);
  expect(buildBlockIndex(pkg).byId.get('gallery-1')?.payload.attributes).toEqual({
    layout: 'row-fit',
    images: [ONE],
  });
  const input = await targetDocumentInput('<w:p><w:r><w:t>Seed</w:t></w:r></w:p>');
  const planned = planTargetDocument(input, tag([{ url: 'javascript:alert(1)' }]), {
    allocateBlockId: () => 'bad',
  });
  const applied = planned.success ? previewDirectPlan(input.pkg, planned.plan) : planned;
  expect(applied.success).toBe(false);
});

test('reordering is one undo unit that changes payload and Word lines together', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(await lesson()),
  });
  try {
    const moved = updateAtomicBlock(editor, {
      blockId: 'gallery-1',
      attributes: { layout: 'row-fit', images: [TWO, ONE] },
    });
    if (!moved.success) throw new Error(moved.error);
    expect(texts(editor)).toEqual(['Intro', TWO.url, 'One', 'After']);
    expect((await projected(editor)).projection.text).toContain(tag([TWO, ONE]));
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect((await projected(editor)).projection.text).toBe(MARKDOWN);
  } finally {
    editor.destroy();
  }
});

for (const first of [0, 1] as const)
  test(`a reorder converges beside peer typing; receiver ${first}`, async () => {
    const pair = await createTwoEditorFixture({ document: writeOoxmlPackage(await lesson()) });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      pair.pause();
      expect(
        updateAtomicBlock(alice, {
          blockId: 'gallery-1',
          attributes: { layout: 'row-fit', images: [TWO, ONE] },
        }).success
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
        expect(texts(editor)).toEqual(['Intro', TWO.url, 'One', 'After Bob']);
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
