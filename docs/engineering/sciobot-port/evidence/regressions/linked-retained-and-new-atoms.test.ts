// Open D30 regression: fresh atoms within retained linked atom sequences.
import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  readOoxmlPackage,
  type OoxmlPackage,
  type ImageDecodePort,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../../../../packages/sciobot/test/support/document';
import { readWorkingView } from '../../../../../packages/sciobot/src/review/working-view';
import { readEditorRevision } from '../../../../../packages/sciobot/src/projection/revision-cache';
import { buildBlockIndex } from '../../../../../packages/sciobot/src/blocks/block-index';
import { applyDirectPlan } from '../../../../../packages/sciobot/src/ai/apply-direct';
import { prepareInlineImageResources } from '../../../../../packages/sciobot/src/native/inline-image-resources';
import { planTargetDocument } from '../../../../../packages/sciobot/src/diff/plan-target';
const context = { documentId: 'inline-images', epoch: 'native' };
const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
const decode: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
const media = (pkg: OoxmlPackage) =>
  [...pkg.partBytes].filter(([name]) => /(^|\/)media\//.test(name));
async function plan(editor: DocxEditorInstance, target: string) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const view = await readWorkingView(editor.surface!.session, revision);
  if (!view.success) throw Error(view.error);
  let next = 0;
  const result = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: view.view.projection },
    target,
    { allocateBlockId: () => `control-${++next}` }
  );
  if (!result.success) throw Error(result.error);
  return result.plan;
}

for (const retained of ['math', 'image'] as const)
  test(`a fresh linked atom preserves the existing linked ${retained} control, media and native history`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      imageDecodePort: decode,
      document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
    });
    const seed =
      retained === 'math'
        ? 'Keep. [$x$](https://outer.test "Read") Z'
        : 'Keep. [![Cat](asset:cat "Photo")](https://outer.test "Read") Z';
    const target = 'Keep. [$x$ ![Cat](asset:cat "Photo")](https://outer.test "Read") Z';
    const apply = async (markdown: string) => {
      const planned = await plan(editor, markdown),
        resolved: string[] = [];
      const prepared = await prepareInlineImageResources(
        planned,
        async (reference) => {
          resolved.push(reference);
          return { bytes: png, mime: 'image/png', widthPoints: 12, heightPoints: 8 };
        },
        decode
      );
      if (!prepared.success) throw Error(prepared.error);
      const result = applyDirectPlan(editor, planned, { ...context, images: prepared.images });
      if (!result.success) throw Error(result.error);
      return resolved;
    };
    try {
      expect(await apply(seed)).toEqual(retained === 'image' ? ['asset:cat'] : []);
      const initial = editor.surface!.session.currentPackage(),
        before = canonicalOoxmlFingerprint(editor.surface!.session.part());
      const original = [...buildBlockIndex(initial).byId.values()][0]!;
      const originalPayload = original.payload;
      const originalWord = canonicalOoxmlFingerprint(
        findNode(editor.surface!.session.part(), original.controlId)!
      );
      expect(await apply(target)).toEqual(retained === 'math' ? ['asset:cat'] : []);
      const pkg = editor.surface!.session.currentPackage(),
        index = buildBlockIndex(pkg);
      expect(index.diagnostics).toEqual([]);
      expect(index.byId.size).toBe(2);
      expect(index.byId.get(original.id)?.controlId).toBe(original.controlId);
      expect(index.byId.get(original.id)?.payload).toEqual(originalPayload);
      expect(
        canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), original.controlId)!)
      ).toBe(originalWord);
      expect(media(pkg)).toHaveLength(1);
      if (retained === 'image') expect(media(pkg)).toEqual(media(initial));
      expect((await plan(editor, target)).operations).toEqual([]);
      const after = canonicalOoxmlFingerprint(editor.surface!.session.part());
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
      expect(media(editor.surface!.session.currentPackage())).toEqual(media(initial));
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(after);
      const saved = new Uint8Array(await editor.save()),
        loaded = readOoxmlPackage(saved);
      if (!loaded.ok) throw Error(loaded.reason);
      expect(
        canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
      ).toBe(after);
      expect(media(loaded.package)).toEqual(media(pkg));
      const reopened = createDocxEditor({
        container: document.createElement('div'),
        imageDecodePort: decode,
        document: saved,
      });
      try {
        expect((await plan(reopened, target)).operations).toEqual([]);
      } finally {
        reopened.destroy();
      }
    } finally {
      editor.destroy();
    }
  });
