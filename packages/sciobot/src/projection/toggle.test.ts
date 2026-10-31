import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { projectSourcePackage, sourceRevision } from '../../test/support/target-source-fixture';
import { previewDirectPlan } from '../ai/apply-direct';
import { buildBlockIndex } from '../blocks/block-index';
import { planTargetDocument } from '../diff/plan-target';
import { readWorkingView } from '../review/working-view';

const MARKDOWN = 'Intro\n\n<toggle>\n\nSummary line\n\nHidden body\n\n</toggle>\n\nAfter';

test('a toggle is a container whose summary and body stay native text', async () => {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Seed</w:t></w:r></w:p>');
  const planned = planTargetDocument(input, MARKDOWN, { allocateBlockId: () => 'toggle-1' });
  if (!planned.success) throw new Error(planned.error);
  const applied = previewDirectPlan(input.pkg, planned.plan);
  if (!applied.success) throw new Error(applied.error);
  const blocks = [...buildBlockIndex(applied.pkg).byId.values()];
  expect(blocks.map((block) => block.payload.type)).toEqual(['toggle']);
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(applied.pkg),
  });
  try {
    const session = editor.surface!.session;
    const view = await readWorkingView(session, {
      documentId: 'toggle',
      epoch: '1',
      revision: session.packageRevision(),
    });
    if (!view.success) throw new Error(view.error);
    expect(view.view.projection.text).toBe(MARKDOWN);
    // Both paragraphs are ordinary native text: an edit to the body is a text edit.
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    const source = {
      revision: sourceRevision,
      pkg: reopened.package,
      projection: await projectSourcePackage(reopened.package),
    };
    const edited = planTargetDocument(source, MARKDOWN.replace('Hidden body', 'Hidden text'), {
      allocateBlockId: () => 'unused',
    });
    if (!edited.success) throw new Error(edited.error);
    expect(edited.plan.operations.map((operation) => operation.kind)).toEqual(['text']);
  } finally {
    editor.destroy();
  }
});
