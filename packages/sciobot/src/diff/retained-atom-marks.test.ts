import { SCIOBOT_CODE_STYLE_ID } from '../native/inline-code';
import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphOffsetIndex,
  readOoxmlPackage,
  runPropertyEdits,
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
const contentLockPolicy = { allowFormattingInLockedContent: true };
const context = { contentLockPolicy, documentId: 'retained-atom-marks', epoch: 'native' };
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
async function plan(editor: DocxEditorInstance, target: string) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const source = await readWorkingView(editor.surface!.session, revision);
  if (!source.success) throw Error(source.error);
  let next = 0;
  const planned = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: source.view.projection },
    target,
    { allocateBlockId: () => `atom-mark-${++next}` }
  );
  if (!planned.success) throw Error(planned.error);
  return planned.plan;
}
for (const atom of [
  '$x$',
  '[Ada](mention:user-7)',
  '<source title="Paper" url="https://source.test" />',
  '![Cat](asset:cat "Photo")',
])
  for (const [wrapper, property] of [
    ['**', 'b'],
    ['_', 'i'],
    ['~~', 'strike'],
    ['<code>', 'rStyle'],
  ] as const)
    test(`retained ${atom} supports reversible ${property} marks without replacing native content`, async () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        imageDecodePort: decode,
        contentLockPolicy,
        document: docxFixture(
          '<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:color w:val="112233"/></w:rPr><w:t>Keep.</w:t></w:r></w:p>'
        ),
      });
      const plain = `Keep. A ${atom} Z`,
        marked = `Keep. A ${wrapper}${atom}${wrapper === '<code>' ? '</code>' : wrapper} Z`;
      try {
        const birth = await plan(editor, plain);
        const prepared = await prepareInlineImageResources(
          birth,
          async () => ({ bytes: png, mime: 'image/png', widthPoints: 12, heightPoints: 8 }),
          decode
        );
        if (!prepared.success) throw Error(prepared.error);
        const seeded = applyDirectPlan(editor, birth, { ...context, images: prepared.images });
        if (!seeded.success) throw Error(seeded.error);
        const pkg = editor.surface!.session.currentPackage(),
          index = buildBlockIndex(pkg);
        const control = [...index.byId.values()][0]!;
        const paragraphId = editor.surface!.session.paragraphIds()[0]!;
        const originalBytes = [...pkg.partBytes];
        let original = editor.surface!.session.part();
        for (const [target, value] of [
          [marked, '1'],
          [plain, '0'],
        ] as const) {
          const before = canonicalOoxmlFingerprint(original);
          const paragraph = findNode(original, paragraphId)!;
          if (paragraph.kind !== 'paragraph') throw Error('Missing paragraph');
          const span = paragraphOffsetIndex(paragraph).spanOf(
            findNode(original, control.controlId)!
          )!;
          let expected = original;
          for (const op of runPropertyEdits(
            original,
            paragraphId,
            span.start,
            span.end,
            property === 'rStyle' ? [] : [{ localName: property, attributes: { val: value } }],
            'proposed'
          )) {
            const result = applyTreeOp(
              expected,
              {
                op: 'setRunProperties',
                paragraphId,
                ...op,
                ...(property === 'rStyle'
                  ? { characterStyle: value === '1' ? SCIOBOT_CODE_STYLE_ID : null }
                  : {}),
              },
              { contentLockPolicy }
            );
            if (!result.ok) throw Error(result.reason);
            expected = result.part;
          }
          const planned = await plan(editor, target);
          expect(planned.operations.length).toBeGreaterThan(0);
          expect(planned.operations.every((op) => op.kind === 'format')).toBe(true);
          expect(decodeEditPlan(serializeEditPlan(planned))).toEqual({
            success: true,
            plan: planned,
          });
          const resources = await prepareInlineImageResources(
            planned,
            async () => {
              throw Error('Retained media must not resolve');
            },
            decode
          );
          expect(resources.success).toBe(true);
          const applied = applyDirectPlan(editor, planned, context);
          if (!applied.success) throw Error(applied.error);
          const actual = editor.surface!.session.part();
          expect(canonicalOoxmlFingerprint(actual)).toBe(canonicalOoxmlFingerprint(expected));
          expect([...buildBlockIndex(editor.surface!.session.currentPackage()).byId]).toEqual([
            ...index.byId,
          ]);
          expect(
            [...editor.surface!.session.currentPackage().partBytes].filter(
              ([name]) =>
                property !== 'rStyle' ||
                ![
                  '/word/styles.xml',
                  '/word/_rels/document.xml.rels',
                  '/[Content_Types].xml',
                ].includes(name.startsWith('/') ? name : '/' + name)
            )
          ).toEqual(
            originalBytes.filter(
              ([name]) =>
                property !== 'rStyle' ||
                ![
                  '/word/styles.xml',
                  '/word/_rels/document.xml.rels',
                  '/[Content_Types].xml',
                ].includes(name.startsWith('/') ? name : '/' + name)
            )
          );
          expect((await plan(editor, target)).operations).toEqual([]);
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
            canonicalOoxmlFingerprint(expected)
          );
          const saved = new Uint8Array(await editor.save());
          const loaded = readOoxmlPackage(saved);
          if (!loaded.ok) throw Error(loaded.reason);
          expect(
            canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
          ).toBe(canonicalOoxmlFingerprint(expected));
          const reopened = createDocxEditor({
            container: document.createElement('div'),
            document: saved,
            imageDecodePort: decode,
            contentLockPolicy,
          });
          try {
            expect((await plan(reopened, target)).operations).toEqual([]);
          } finally {
            reopened.destroy();
          }
          original = editor.surface!.session.part();
        }
      } finally {
        editor.destroy();
      }
    });

for (const nativePolicy of [false, true])
  test(`format previews cannot override the host policy: native opt-in ${nativePolicy}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
      ...(nativePolicy ? { contentLockPolicy } : {}),
    });
    try {
      const birth = applyDirectPlan(editor, await plan(editor, 'Keep. $x$ Z'), context);
      if (!birth.success) throw Error(birth.error);
      const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
      const planned = await plan(editor, 'Keep. **$x$** Z');
      let publications = 0;
      const stop = editor.surface!.session.subscribe(() => publications++);
      const result = applyDirectPlan(editor, planned, {
        documentId: context.documentId,
        epoch: context.epoch,
        ...(!nativePolicy ? { contentLockPolicy } : {}),
      });
      stop();
      expect(result.success).toBe(false);
      expect(publications).toBe(0);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
    } finally {
      editor.destroy();
    }
  });
