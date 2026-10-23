// Open D30 regression: linked image births still require native reservation ownership.
import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
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
for (const target of [
  'Keep. [![Cat](asset:cat "Photo")](https://outer.test "Read") Z',
  'Keep.\n\n[![Cat](asset:cat "Photo")](https://outer.test "Read")',
  'Keep.\n\n| H |\n| - |\n| [![Cat](asset:cat "Photo")](https://outer.test "Read") |',
  'Keep.\n\n<phase type="evocation">\n[![Cat](asset:cat "Photo")](https://outer.test "Read")\n</phase>',
])
  test(`new inline drawings retain source references, media, native history and DOCX: ${target.slice(0, 130)}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      imageDecodePort: decode,
      document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
    });
    try {
      const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
      const originalId = editor.surface!.session.paragraphIds()[0]!;
      const planned = await plan(editor, target);
      const original = editor.surface!.session.currentPackage();
      const resolved: string[] = [];
      const prepared = await prepareInlineImageResources(
        planned,
        async (reference) => {
          resolved.push(reference);
          return { bytes: png, mime: 'image/png', widthPoints: 12, heightPoints: 8 };
        },
        decode
      );
      if (!prepared.success) throw Error(prepared.error);
      expect(editor.surface!.session.currentPackage()).toBe(original);
      expect(media(original)).toEqual([]);
      expect(resolved).toEqual(
        target.includes('asset:dog') ? ['asset:cat', 'asset:dog'] : ['asset:cat']
      );
      let publications = 0;
      const stop = editor.surface!.session.subscribe(() => publications++);
      const applied = applyDirectPlan(editor, planned, { ...context, images: prepared.images });
      stop();
      expect(publications).toBe(1);
      if (!applied.success) throw Error(applied.error);
      expect(editor.surface!.session.paragraphIds()).toContain(originalId);
      const pkg = editor.surface!.session.currentPackage();
      const index = buildBlockIndex(pkg);
      expect(index.diagnostics).toEqual([]);
      const images = [...index.byId.values()].filter(
        (block) => block.payload.type === 'inline_image'
      );
      expect(images).toHaveLength(target.includes('asset:dog') ? 2 : 1);
      expect(media(pkg)).toHaveLength(images.length);
      for (const [, bytes] of media(pkg)) expect(bytes).toEqual(png);
      for (const image of images) {
        expect(image.payload.attributes.widthPoints).toBe(12);
        expect(image.payload.attributes.heightPoints).toBe(8);
        expect(String(image.payload.attributes.contentHash)).toHaveLength(64);
      }
      expect((await plan(editor, target)).operations).toEqual([]);
      const after = canonicalOoxmlFingerprint(editor.surface!.session.part());
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
      expect(media(editor.surface!.session.currentPackage())).toEqual([]);
      expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(0);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(after);
      expect(media(editor.surface!.session.currentPackage())).toEqual(media(pkg));
      const bytes = new Uint8Array(await editor.save());
      const loaded = readOoxmlPackage(bytes);
      if (!loaded.ok) throw Error(loaded.reason);
      expect(
        canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
      ).toBe(after);
      expect(media(loaded.package)).toEqual(media(pkg));
      expect(
        [...buildBlockIndex(loaded.package).byId].map(([id, block]) => [id, block.payload])
      ).toEqual([...index.byId].map(([id, block]) => [id, block.payload]));
      const reopened = createDocxEditor({
        container: document.createElement('div'),
        document: bytes,
        imageDecodePort: decode,
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
