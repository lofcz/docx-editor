// Fresh atoms preserve native controls and resources within retained inline sequences.
import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  readOoxmlPackage,
  paragraphOffsetIndex,
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
  for (const side of ['prefix', 'middle', 'suffix'] as const)
    for (const ownership of ['plain', 'linked', 'code-linked'] as const)
      test(`identical atoms preserve both ${retained} identities: ${side}/${ownership}`, async () => {
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
        const wrap = (value: string) =>
          ownership === 'plain'
            ? value
            : `[${ownership === 'code-linked' ? `<code>${value}</code>` : value}](https://outer.test "Read")`;
        const body = `Lead. ${atom} bridge ${atom} Tail.`;
        const seed = wrap(body);
        const target = wrap(
          side === 'prefix'
            ? `${atom} Fresh ${body}`
            : side === 'middle'
              ? `Lead. ${atom} ${atom} bridge ${atom} Tail.`
              : `${body} Fresh ${atom}`
        );
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
          const paragraphId = editor.surface!.session.paragraphIds()[0]!;
          const native = findNode(editor.surface!.session.part(), paragraphId)!;
          if (native.kind !== 'paragraph') throw Error('Missing paragraph');
          const offsets = paragraphOffsetIndex(native);
          const originals = [...buildBlockIndex(initial).byId.values()]
            .map((block) => ({
              block,
              span: offsets.spanOf(block.controlId)!,
              word: canonicalOoxmlFingerprint(
                findNode(editor.surface!.session.part(), block.controlId)!
              ),
            }))
            .sort((a, b) => a.span.start - b.span.start);
          expect(originals).toHaveLength(2);
          const planned = await plan(editor, target);
          expect(
            planned.operations.flatMap((edit) => (edit.kind === 'inline-atoms' ? edit.atoms : []))
          ).toHaveLength(1);
          expect(await apply(target)).toEqual(retained === 'image' ? ['asset:cat'] : []);
          const pkg = editor.surface!.session.currentPackage(),
            index = buildBlockIndex(pkg);
          expect(index.diagnostics).toEqual([]);
          expect(index.byId.size).toBe(3);
          const afterParagraph = findNode(editor.surface!.session.part(), paragraphId)!;
          if (afterParagraph.kind !== 'paragraph') throw Error('Missing paragraph');
          const afterOffsets = paragraphOffsetIndex(afterParagraph);
          for (const [ordinal, original] of originals.entries()) {
            const { block, span, word } = original;
            expect(index.byId.get(block.id)?.controlId).toBe(block.controlId);
            expect(index.byId.get(block.id)?.payload).toEqual(block.payload);
            expect(
              canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), block.controlId)!)
            ).toBe(word);
            const width = span.end - span.start;
            const shift =
              side === 'prefix'
                ? width + ' Fresh '.length
                : side === 'middle' && ordinal === 1
                  ? width + 1
                  : 0;
            expect(afterOffsets.spanOf(block.controlId)).toEqual({
              start: span.start + shift,
              end: span.end + shift,
            });
          }
          expect(media(pkg)).toHaveLength(retained === 'image' ? 3 : 0);
          for (const [name, bytes] of media(initial))
            expect(pkg.partBytes.get(name)).toEqual(bytes);
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
