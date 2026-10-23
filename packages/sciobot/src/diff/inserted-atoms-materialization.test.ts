import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { buildBlockIndex } from '../blocks/block-index';
import { applyDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from './plan-target';
import { decodeEditPlan, serializeEditPlan } from './plan';
const context = { documentId: 'atoms', epoch: 'native' };
const sentence =
  'A $x^2$ [Ada](mention:user-7) <source url="https://example.test" title="Paper" /> Z';
async function plan(editor: DocxEditorInstance, target: string) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const view = await readWorkingView(editor.surface!.session, revision);
  if (!view.success) throw Error(view.error);
  let next = 0;
  const result = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: view.view.projection },
    target,
    { allocateBlockId: () => `new-${++next}` }
  );
  if (!result.success) throw Error(result.error);
  return result.plan;
}
for (const target of [
  `Keep.\n\n${sentence}`,
  'Keep.\n\n# $x^2$',
  'Keep.\n\n**$x^2$**',
  'Keep.\n\n[Link](https://link.test)$x^2$',
  '<phase type="evocation">\n[Link](https://link.test)$x^2$\n</phase>',
  'Keep.\n\n<scio-inline-equation texExpression="" />',
  `Keep.\n\n<phase type="evocation">\n${sentence}\n</phase>`,
  'Keep.\n\n| H |\n| - |\n| A $x^2$ [Ada](mention:user-7) |',
])
  test(`new inline controls land through native plans, history and DOCX: ${target.slice(0, 70)}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
    });
    try {
      const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
      const planned = await plan(editor, target);
      const operations = planned.operations.filter((edit) => edit.kind === 'siblings');
      expect(operations.some((edit) => edit.atoms?.length)).toBe(true);
      expect(decodeEditPlan(serializeEditPlan(planned))).toEqual({ success: true, plan: planned });
      const applied = applyDirectPlan(editor, planned, context);
      if (!applied.success) throw Error(applied.error);
      expect((await plan(editor, target)).operations).toEqual([]);
      const index = buildBlockIndex(editor.surface!.session.currentPackage());
      expect(index.diagnostics).toEqual([]);
      expect(
        [...index.byId.values()].filter((block) =>
          ['inline_equation', 'mention', 'source'].includes(block.payload.type)
        ).length
      ).toBeGreaterThan(0);
      if (target.includes(sentence))
        expect(
          editor
            .surface!.session.paragraphIds()
            .map((id) => paragraphTextOf(editor.surface!.session.part(), id))
        ).toContain('A x^2 Ada Paper Z');
      const after = canonicalOoxmlFingerprint(editor.surface!.session.part());
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
      expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(0);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(after);
      const bytes = new Uint8Array(await editor.save()),
        loaded = readOoxmlPackage(bytes);
      if (!loaded.ok) throw Error(loaded.reason);
      expect(
        [...buildBlockIndex(loaded.package).byId].map(([id, block]) => [id, block.payload])
      ).toEqual([...index.byId].map(([id, block]) => [id, block.payload]));
      expect(
        canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
      ).toBe(after);
      const reopened = createDocxEditor({
        container: document.createElement('div'),
        document: bytes,
      });
      try {
        expect(canonicalOoxmlFingerprint(reopened.surface!.session.part())).toBe(after);
        expect((await plan(reopened, target)).operations).toEqual([]);
      } finally {
        reopened.destroy();
      }
    } finally {
      editor.destroy();
    }
  });

test('D01 atom bindings reject missing records, changed fallbacks, moved paths and reused identities', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
  });
  try {
    const planned = await plan(editor, `Keep.\n\n${sentence}`);
    for (const change of ['missing', 'fallback', 'path', 'identity']) {
      const copy = JSON.parse(serializeEditPlan(planned));
      const edit = copy.operations.find((edit: { atoms?: unknown[] }) => edit.atoms?.length);
      if (change === 'missing') edit.atoms.pop();
      if (change === 'fallback') edit.atoms[0].fallback = 'changed';
      if (change === 'path') edit.atoms[0].location.path = [999];
      if (change === 'identity') edit.atoms[1].payload.id = edit.atoms[0].payload.id;
      expect(decodeEditPlan(JSON.stringify(copy)).success).toBe(false);
    }
  } finally {
    editor.destroy();
  }
});

for (const target of [
  'Keep.\n\n[$x^2$](https://example.test)',
  'Keep.\n\nA ![Figure](asset:pending) Z',
])
  test(`unsupported atom materialization refuses before publication: ${target}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
    });
    try {
      const planned = await plan(editor, target);
      const before = editor.surface!.session.currentPackage();
      const revision = readEditorRevision(editor, context.documentId, context.epoch);
      const applied = applyDirectPlan(editor, planned, context);
      expect(applied).toMatchObject({ success: false, code: 'unsupported-operation' });
      expect(editor.surface!.session.currentPackage()).toBe(before);
      expect(readEditorRevision(editor, context.documentId, context.epoch)).toEqual(revision);
      expect(editor.surface!.session.canUndo()).toBe(false);
      expect(buildBlockIndex(before).byId.size).toBe(0);
    } finally {
      editor.destroy();
    }
  });
