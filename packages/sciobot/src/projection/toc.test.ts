import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { paragraphTextOf, readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { projectSourcePackage, sourceRevision } from '../../test/support/target-source-fixture';
import { previewDirectPlan } from '../ai/apply-direct';
import { buildBlockIndex } from '../blocks/block-index';
import { planTargetDocument } from '../diff/plan-target';

const MARKDOWN = '# Lesson\n\n<toc />\n\n## Part one\n\nBody';

test('a table of contents is an atomic block that round-trips and keeps a Word title line', async () => {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Seed</w:t></w:r></w:p>');
  const planned = planTargetDocument(input, MARKDOWN, { allocateBlockId: () => 'toc-1' });
  if (!planned.success) throw new Error(planned.error);
  const applied = previewDirectPlan(input.pkg, planned.plan);
  if (!applied.success) throw new Error(applied.error);
  expect(buildBlockIndex(applied.pkg).byId.get('toc-1')?.payload).toMatchObject({
    type: 'toc',
    attributes: {},
  });
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(applied.pkg),
  });
  try {
    const session = editor.surface!.session;
    expect(session.paragraphIds().map((id) => paragraphTextOf(session.part(), id))).toContain(
      'Table of contents'
    );
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    const projection = await projectSourcePackage(reopened.package);
    expect(projection.text).toBe(MARKDOWN);
    const again = planTargetDocument(
      { revision: sourceRevision, pkg: reopened.package, projection },
      MARKDOWN,
      { allocateBlockId: () => 'unused' }
    );
    if (!again.success) throw new Error(again.error);
    expect(again.plan.operations).toHaveLength(0);
  } finally {
    editor.destroy();
  }
});
