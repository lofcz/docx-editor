import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  contentControlsIn,
  paragraphTextOf,
  readOoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { buildBlockIndex } from '../blocks/block-index';
import { applyDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from './plan-target';
import { decodeEditPlan, serializeEditPlan } from './plan';
const context = { documentId: 'linked-atoms', epoch: 'native' };
const sentence =
  '[**$x$** and <source title="Paper" url="https://source.test" />](https://outer.test "Read")';
async function plan(editor: DocxEditorInstance, target: string) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const view = await readWorkingView(editor.surface!.session, revision);
  if (!view.success) throw Error(view.error);
  let next = 0;
  const planned = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: view.view.projection },
    target,
    { allocateBlockId: () => `linked-${++next}` }
  );
  if (!planned.success) throw Error(planned.error);
  return planned.plan;
}
function descendants(node: OoxmlNode): OoxmlNode[] {
  return [node, ...(node.kind === 'textValue' ? [] : node.children.flatMap(descendants))];
}
for (const [kind, target] of [
  ['paragraph', `Keep.\n\n${sentence}`],
  ['heading', `Keep.\n\n# ${sentence}`],
  ['table', `Keep.\n\n| H |\n| - |\n| ${sentence} |`],
  ['registered', `Keep.\n\n<phase type="evocation">\n${sentence}\n</phase>`],
  ['retained', `Keep. ${sentence} End.`],
  ['imported', `Keep. ${sentence} End.`],
] as const)
  test(`linked inline controls retain exact native ownership in ${kind}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document:
        kind === 'imported'
          ? docxFixture(
              '<w:p><w:r><w:rPr><w:color w:val="445566"/></w:rPr><w:t xml:space="preserve">Keep. </w:t></w:r><w:hyperlink r:id="existing" w:tooltip="Read" w:history="1"><w:r><w:rPr><w:color w:val="112233"/></w:rPr><w:t xml:space="preserve"> and </w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> End.</w:t></w:r></w:p>',
              {
                relationships:
                  '<Relationship Id="existing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://outer.test" TargetMode="External"/>',
              }
            )
          : docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
    });
    try {
      const initial = canonicalOoxmlFingerprint(editor.surface!.session.part());
      const imported = descendants(editor.surface!.session.part().root).find(
        (node) => node.kind === 'hyperlink'
      );
      const planned = await plan(editor, target);
      expect(decodeEditPlan(serializeEditPlan(planned))).toEqual({ success: true, plan: planned });
      const applied = applyDirectPlan(editor, planned, context);
      if (!applied.success) throw Error(applied.error);
      const pkg = editor.surface!.session.currentPackage();
      const part = editor.surface!.session.part();
      const index = buildBlockIndex(pkg);
      expect(index.diagnostics).toEqual([]);
      const links = descendants(part.root).filter(
        (node) => node.kind !== 'textValue' && node.localName === 'hyperlink'
      );
      expect(links).toHaveLength(1);
      const link = links[0]!;
      if (link.kind === 'textValue') throw Error('Missing native hyperlink');
      expect(contentControlsIn(link)).toHaveLength(2);
      expect(link.attributes.find((a) => a.localName === 'tooltip')?.value).toBe('Read');
      const relationship = link.attributes.find((a) => a.localName === 'id')?.value;
      expect(pkg.relationships.get(part.name)?.find((r) => r.id === relationship)?.rawTarget).toBe(
        'https://outer.test'
      );
      const controls = contentControlsIn(link).map(({ node }) => node);
      if (kind === 'imported') {
        expect(link.id).toBe(imported!.id);
        expect(link.attributes).toEqual(
          (imported as Exclude<OoxmlNode, { kind: 'textValue' }>).attributes
        );
        const colors = controls.map((control) =>
          descendants(control)
            .filter((node) => node.kind !== 'textValue' && node.localName === 'color')
            .map((node) =>
              node.kind !== 'textValue'
                ? node.attributes.find((attribute) => attribute.localName === 'val')?.value
                : null
            )
        );
        expect(colors).toEqual([['445566'], ['112233']]);
      }
      const bold = descendants(controls[0]!).filter(
        (node) => node.kind !== 'textValue' && node.localName === 'b'
      );
      expect(bold).toHaveLength(1);
      expect(
        descendants(controls[1]!).filter(
          (node) => node.kind !== 'textValue' && node.localName === 'b'
        )
      ).toHaveLength(0);
      expect(
        [...index.byId.values()].filter((block) => block.payload.id.startsWith('linked-'))
      ).toHaveLength(kind === 'registered' ? 3 : 2);
      expect(
        editor
          .surface!.session.paragraphIds()
          .map((id) => paragraphTextOf(part, id))
          .join('\n')
      ).toContain('x and Paper');
      expect((await plan(editor, target)).operations).toEqual([]);
      const fingerprint = canonicalOoxmlFingerprint(part);
      const identities = controls.map((node) => node.id);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(initial);
      expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(0);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(fingerprint);
      const reopenedBytes = new Uint8Array(await editor.save());
      const loaded = readOoxmlPackage(reopenedBytes);
      if (!loaded.ok) throw Error(loaded.reason);
      expect(
        canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
      ).toBe(fingerprint);
      expect(
        [...buildBlockIndex(loaded.package).byId].map(([id, block]) => [id, block.payload])
      ).toEqual([...index.byId].map(([id, block]) => [id, block.payload]));
      const reopened = createDocxEditor({
        container: document.createElement('div'),
        document: reopenedBytes,
      });
      try {
        expect((await plan(reopened, target)).operations).toEqual([]);
      } finally {
        reopened.destroy();
      }
      expect(
        contentControlsIn(editor.surface!.session.part().root).filter(({ node }) =>
          identities.includes(node.id)
        )
      ).toHaveLength(2);
    } finally {
      editor.destroy();
    }
  });
