// Open D30 regression: repeated-space insertion discards an imported native hyperlink.
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
test('a linked image birth preserves imported link metadata, inherited native properties and reversible code styles', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    imageDecodePort: decode,
    document: docxFixture(
      '<w:p><w:r><w:t xml:space="preserve">Keep. </w:t></w:r><w:hyperlink r:id="existing" w:tooltip="Read" w:history="1" w:tgtFrame="_blank"><w:r><w:rPr xmlns:proof="urn:sciobot:proof"><w:rStyle w:val="ImportedBase"/><w:color w:val="112233"/><proof:metadata proof:keep="yes"/></w:rPr><w:t>label</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> End.</w:t></w:r></w:p>',
      {
        relationships:
          '<Relationship Id="existing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://outer.test" TargetMode="External"/>',
        styles:
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="character" w:styleId="ImportedBase"><w:name w:val="Imported base"/><w:rPr><w:color w:val="112233"/></w:rPr></w:style></w:styles>',
      }
    ),
  });
  const nodes = (
    node: import('@docx-editor.dev/core/store').OoxmlNode
  ): import('@docx-editor.dev/core/store').OoxmlNode[] => [
    node,
    ...(node.kind === 'textValue' ? [] : node.children.flatMap(nodes)),
  ];
  const target =
    'Keep. [label <code>![Cat](asset:cat "Photo")</code>](https://outer.test "Read") End.';
  try {
    const before = editor.surface!.session.part(),
      beforeWord = canonicalOoxmlFingerprint(before);
    const link = nodes(before.root).find((node) => node.kind === 'hyperlink')!;
    if (link.kind !== 'hyperlink') throw Error('Missing native link');
    const metadata = nodes(link).find(
      (node) => node.kind !== 'textValue' && node.namespaceUri === 'urn:sciobot:proof'
    )!;
    const planned = await plan(editor, target);
    const prepared = await prepareInlineImageResources(
      planned,
      async () => ({ bytes: png, mime: 'image/png', widthPoints: 12, heightPoints: 8 }),
      decode
    );
    if (!prepared.success) throw Error(prepared.error);
    const applied = applyDirectPlan(editor, planned, { ...context, images: prepared.images });
    if (!applied.success) throw Error(applied.error);
    const part = editor.surface!.session.part(),
      currentLink = findNode(part, link.id);
    expect(currentLink?.kind).toBe('hyperlink');
    if (currentLink?.kind !== 'hyperlink') throw Error('The imported link disappeared');
    expect(currentLink.attributes).toEqual(link.attributes);
    const image = [...buildBlockIndex(editor.surface!.session.currentPackage()).byId.values()].find(
      (block) => block.payload.type === 'inline_image'
    )!;
    const control = findNode(part, image.controlId)!;
    const elements = nodes(control).filter((node) => node.kind !== 'textValue');
    const property = (name: string) =>
      elements.find((node) => node.kind !== 'textValue' && node.localName === name);
    const style = property('rStyle'),
      color = property('color');
    expect(
      style?.kind !== 'textValue' &&
        style?.attributes.find((attribute) => attribute.localName === 'val')?.value
    ).toBe('SciobotCode__ImportedBase');
    expect(
      color?.kind !== 'textValue' &&
        color?.attributes.find((attribute) => attribute.localName === 'val')?.value
    ).toBe('112233');
    const copied = elements.find(
      (node) => node.kind !== 'textValue' && node.namespaceUri === 'urn:sciobot:proof'
    )!;
    expect(canonicalOoxmlFingerprint(copied)).toBe(canonicalOoxmlFingerprint(metadata));
    expect(nodes(currentLink).some((node) => node.id === image.controlId)).toBe(true);
    expect((await plan(editor, target)).operations).toEqual([]);
    const after = canonicalOoxmlFingerprint(part);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(beforeWord);
    expect(media(editor.surface!.session.currentPackage())).toEqual([]);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(after);
    const saved = new Uint8Array(await editor.save()),
      loaded = readOoxmlPackage(saved);
    if (!loaded.ok) throw Error(loaded.reason);
    expect(
      canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
    ).toBe(after);
    const reopened = createDocxEditor({
      container: document.createElement('div'),
      document: saved,
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
