import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  contentControlsIn,
  readOoxmlPackage,
  stylesPartOf,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { buildBlockIndex } from '../blocks/block-index';
import { applyDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from './plan-target';
import { decodeEditPlan, serializeEditPlan } from './plan';
import { SCIOBOT_CODE_STYLE_ID, inlineCodeStyleId } from '../native/inline-code';
const context = { documentId: 'code-atoms', epoch: 'native' };
const sentence = 'A <code>**$x$** and <source title="Paper" url="https://source.test" /></code> Z';
async function plan(editor: DocxEditorInstance, target: string) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const source = await readWorkingView(editor.surface!.session, revision);
  if (!source.success) throw Error(source.error);
  let next = 0;
  const planned = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: source.view.projection },
    target,
    { allocateBlockId: () => `code-${++next}` }
  );
  if (!planned.success) throw Error(planned.error);
  return planned.plan;
}
const descendants = (node: OoxmlNode): OoxmlNode[] => [
  node,
  ...(node.kind === 'textValue' ? [] : node.children.flatMap(descendants)),
];
for (const [kind, target] of [
  ['paragraph', `Keep.\n\n${sentence}`],
  ['heading', `Keep.\n\n# ${sentence}`],
  ['table', `Keep.\n\n| H |\n| - |\n| ${sentence} |`],
  ['registered', `Keep.\n\n<phase type="evocation">\n${sentence}\n</phase>`],
  ['retained', `Keep. ${sentence} End.`],
  ['mention', 'Keep. A <code>**$x$** and [Ada](mention:user-7)</code> Z End.'],
  ['imported', `Keep. ${sentence.replace('<code>', '<code>lead ')} End.`],
  [
    'linked',
    'Keep. [A <code>**$x$** and <source title="Paper" url="https://source.test" /></code> Z](https://outer.test "Read") End.',
  ],
] as const)
  test(`code-marked inline controls retain native style and payload ownership in ${kind}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document:
        kind === 'imported'
          ? docxFixture(
              '<w:p><w:r><w:rPr><w:rStyle w:val="ImportedBase"/><w:color w:val="112233"/></w:rPr><w:t>Keep.</w:t></w:r></w:p>',
              {
                styles:
                  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="character" w:styleId="ImportedBase"><w:name w:val="Imported base"/><w:rPr><w:color w:val="112233"/></w:rPr></w:style></w:styles>',
              }
            )
          : docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
    });
    try {
      const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
      const planned = await plan(editor, target);
      expect(decodeEditPlan(serializeEditPlan(planned))).toEqual({ success: true, plan: planned });
      const applied = applyDirectPlan(editor, planned, context);
      if (!applied.success) throw Error(applied.error);
      const pkg = editor.surface!.session.currentPackage();
      const part = editor.surface!.session.part();
      const index = buildBlockIndex(pkg);
      expect(index.diagnostics).toEqual([]);
      const atoms = [...index.byId.values()].filter((block) =>
        ['inline_equation', 'mention', 'source'].includes(block.payload.type)
      );
      expect(atoms).toHaveLength(2);
      if (kind === 'mention')
        expect(atoms.find((atom) => atom.payload.type === 'mention')?.payload.attributes).toEqual({
          key: 'user-7',
          value: 'Ada',
        });
      for (const atom of atoms) {
        const control = contentControlsIn(part.root).find(
          ({ node }) => node.id === atom.controlId
        )!.node;
        const styles = descendants(control)
          .filter((node) => node.kind !== 'textValue' && node.localName === 'rStyle')
          .map((node) =>
            node.kind !== 'textValue'
              ? node.attributes.find((attribute) => attribute.localName === 'val')?.value
              : null
          );
        expect(styles).toEqual([
          kind === 'imported' ? inlineCodeStyleId('ImportedBase') : SCIOBOT_CODE_STYLE_ID,
        ]);
        if (kind === 'imported') {
          expect(
            descendants(control)
              .filter((node) => node.kind !== 'textValue' && node.localName === 'color')
              .map((node) =>
                node.kind !== 'textValue'
                  ? node.attributes.find((attribute) => attribute.localName === 'val')?.value
                  : null
              )
          ).toEqual(['112233']);
        }
        expect(
          descendants(control).filter(
            (node) => node.kind !== 'textValue' && node.localName === 'strike'
          )
        ).toHaveLength(0);
      }
      const styles = stylesPartOf(pkg);
      expect(styles).toBeDefined();
      expect(
        descendants(styles!.root).filter(
          (node) =>
            node.kind !== 'textValue' &&
            node.localName === 'style' &&
            node.attributes.some(
              (attribute) =>
                attribute.localName === 'styleId' && attribute.value === SCIOBOT_CODE_STYLE_ID
            )
        )
      ).toHaveLength(1);
      expect((await plan(editor, target)).operations).toEqual([]);
      const after = canonicalOoxmlFingerprint(part);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
      expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(0);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(after);
      const bytes = new Uint8Array(await editor.save());
      const opened = readOoxmlPackage(bytes);
      if (!opened.ok) throw Error(opened.reason);
      expect(
        canonicalOoxmlFingerprint(opened.package.parts.get(opened.package.mainDocumentPart)!)
      ).toBe(after);
      expect(
        [...buildBlockIndex(opened.package).byId].map(([id, block]) => [id, block.payload])
      ).toEqual([...index.byId].map(([id, block]) => [id, block.payload]));
      const reopened = createDocxEditor({
        container: document.createElement('div'),
        document: bytes,
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
