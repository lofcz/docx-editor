// Fresh atoms preserve native controls and resources within retained inline sequences.
import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  readOoxmlPackage,
  type OoxmlPackage,
  type ImageDecodePort,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { buildBlockIndex } from '../blocks/block-index';
import { applyDirectPlan } from '../ai/apply-direct';
import { prepareInlineImageResources } from '../native/inline-image-resources';
import { planTargetDocument } from './plan-target';
import { decodeEditPlan, serializeEditPlan } from './plan';
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
  const result = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: view.view.projection },
    target,
    {}
  );
  if (!result.success) throw Error(result.error);
  const decoded = decodeEditPlan(serializeEditPlan(result.plan));
  expect(decoded).toEqual({ success: true, plan: result.plan });
  if (!decoded.success) throw Error(decoded.error);
  return decoded.plan;
}

for (const retained of ['math', 'image', 'mention', 'source'] as const)
  for (const side of ['before', 'after'] as const)
    for (const ownership of ['plain', 'linked', 'code-linked'] as const)
      for (const gapEdits of [false, true])
        test(`fresh atoms preserve ${retained} native history: ${side}/${ownership}/gaps=${gapEdits}`, async () => {
          const editor = createDocxEditor({
            container: document.createElement('div'),
            imageDecodePort: decode,
            document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
          });
          const atom =
            retained === 'math'
              ? '$x$'
              : retained === 'image'
                ? '![Cat](asset:cat "Photo")'
                : retained === 'mention'
                  ? '<scio-mention key="user" value="Ada" />'
                  : '<source title="Paper" url="https://source.test" />';
          const fresh = retained === 'image' ? '$x$' : '![Cat](asset:cat "Photo")';
          const wrap = (value: string) =>
            ownership === 'plain'
              ? value
              : `[${ownership === 'code-linked' ? `<code>${value}</code>` : value}](https://outer.test "Read")`;
          const seed = `Keep. ${gapEdits ? 'x x ' : ''}${wrap(atom)}${gapEdits ? ' x x' : ''} Z`;
          const target = `Keep. ${gapEdits ? 'x ' : ''}${wrap(side === 'after' ? `${atom} ${fresh}` : `${fresh} ${atom}`)}${gapEdits ? ' x' : ''} Z`;
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
            const result = applyDirectPlan(editor, planned, {
              ...context,
              images: prepared.images,
            });
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
            expect(await apply(target)).toEqual(retained === 'image' ? [] : ['asset:cat']);
            const pkg = editor.surface!.session.currentPackage(),
              index = buildBlockIndex(pkg);
            expect(index.diagnostics).toEqual([]);
            expect(index.byId.size).toBe(2);
            expect(index.byId.get(original.id)?.controlId).toBe(original.controlId);
            expect(index.byId.get(original.id)?.payload).toEqual(originalPayload);
            expect(
              canonicalOoxmlFingerprint(
                findNode(editor.surface!.session.part(), original.controlId)!
              )
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
