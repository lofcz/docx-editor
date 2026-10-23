import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { docxFixture } from '../../test/support/document';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { applyDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from './plan-target';
import { decodeEditPlan, serializeEditPlan, type CanonicalEditPlan } from './plan';

const context = { documentId: 'retained-identities', epoch: 'native' };
for (const target of ['Keep $x$\n\nOther $y$', 'Keep $x$\n\nOther\n\nNew $y$'])
  test(`cross-operation birth identities reject atomically: ${target}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture(
        '<w:p><w:r><w:t>Keep</w:t></w:r></w:p><w:p><w:r><w:t>Other</w:t></w:r></w:p>'
      ),
    });
    try {
      const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
      const view = await readWorkingView(editor.surface!.session, revision);
      if (!view.success) throw Error(view.error);
      let next = 0;
      const planned = planTargetDocument(
        {
          revision,
          pkg: editor.surface!.session.currentPackage(),
          projection: view.view.projection,
        },
        target,
        { allocateBlockId: () => `born-${++next}` }
      );
      if (!planned.success) throw Error(planned.error);
      expect(decodeEditPlan(serializeEditPlan(planned.plan))).toEqual({
        success: true,
        plan: planned.plan,
      });
      const copy = JSON.parse(serializeEditPlan(planned.plan)) as CanonicalEditPlan;
      const births = copy.operations
        .flatMap((edit) =>
          edit.kind === 'inline-atoms' || edit.kind === 'siblings' ? (edit.atoms ?? []) : []
        )
        .filter((atom) => atom.kind === 'inline-control');
      expect(births).toHaveLength(2);
      (births[1]!.payload as { id: string }).id = births[0]!.payload.id;
      expect(decodeEditPlan(JSON.stringify(copy))).toEqual({
        success: false,
        code: 'invalid-input',
        error: 'duplicate-inserted-block-id',
      });
      const before = editor.surface!.session.currentPackage();
      expect(applyDirectPlan(editor, copy, context)).toMatchObject({
        success: false,
        code: 'invalid-input',
      });
      expect(editor.surface!.session.currentPackage()).toBe(before);
      expect(editor.surface!.session.canUndo()).toBe(false);
      expect(readEditorRevision(editor, context.documentId, context.epoch)).toEqual(revision);
    } finally {
      editor.destroy();
    }
  });
