import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  readOoxmlPackage,
  resolveImageRelationship,
  type ImageDecodePort,
  type OoxmlNode,
  type OoxmlPart,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { buildBlockIndex } from '../blocks/block-index';
import { planTargetDocument } from './plan-target';
import { serializeEditPlan, decodeEditPlan } from './plan';
import { applyDirectPlan } from '../ai/apply-direct';
import { prepareInlineImageResources } from '../native/inline-image-resources';
const context = { documentId: 'image-replacement', epoch: 'native' };
const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
const data = new TextEncoder().encode('Comment\0Replacement');
const chunk = new Uint8Array(12 + data.length);
new DataView(chunk.buffer).setUint32(0, data.length);
chunk.set([116, 69, 88, 116], 4);
chunk.set(data, 8);
let crc = 0xffffffff;
for (const byte of chunk.subarray(4, -4)) {
  crc ^= byte;
  for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
}
new DataView(chunk.buffer).setUint32(chunk.length - 4, (crc ^ 0xffffffff) >>> 0);
const replacement = new Uint8Array(png.length + chunk.length);
replacement.set(png.subarray(0, 33));
replacement.set(chunk, 33);
replacement.set(png.subarray(33), 33 + chunk.length);
const decode: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
const media = (pkg: OoxmlPackage) =>
  [...pkg.partBytes]
    .filter(([name]) => /(^|\/)media\//.test(name))
    .sort(([a], [b]) => a.localeCompare(b));
const descendants = (node: OoxmlNode): OoxmlNode[] => [
  node,
  ...(node.kind === 'textValue' ? [] : node.children.flatMap(descendants)),
];
function op(part: OoxmlPart, value: TreeDocOp): OoxmlPart {
  const applied = applyTreeOp(part, value);
  if (!applied.ok) throw Error(applied.reason);
  return applied.part;
}
async function plan(editor: DocxEditorInstance, target: string) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const source = await readWorkingView(editor.surface!.session, revision);
  if (!source.success) throw Error(source.error);
  const planned = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: source.view.projection },
    target
  );
  if (!planned.success) throw Error(planned.error);
  const decoded = decodeEditPlan(serializeEditPlan(planned.plan));
  expect(decoded).toEqual({ success: true, plan: planned.plan });
  if (!decoded.success) throw Error(decoded.error);
  return decoded.plan;
}
async function apply(editor: DocxEditorInstance, target: string) {
  const planned = await plan(editor, target),
    resolved: string[] = [];
  const prepared = await prepareInlineImageResources(
    planned,
    async (reference) => {
      resolved.push(reference);
      return {
        bytes: reference === 'asset:cat' ? png : replacement,
        mime: 'image/png',
        widthPoints: reference === 'asset:cat' ? 12 : 40,
        heightPoints: reference === 'asset:cat' ? 8 : 30,
      };
    },
    decode
  );
  if (!prepared.success) throw Error(prepared.error);
  const result = applyDirectPlan(editor, planned, { ...context, images: prepared.images });
  if (!result.success) throw Error(result.error);
  return { planned, resolved };
}
for (const mode of ['resource', 'metadata', 'both'] as const)
  for (const ownership of ['plain', 'linked', 'code-linked'] as const)
    for (const text of [false, true])
      for (const alternate of [false, true])
        test(`image replacement preserves native identity, geometry and history: ${mode}/${ownership}/text=${text}/alternate=${alternate}`, async () => {
          let documentBytes = docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>');
          if (alternate) {
            const entries = unzipSync(documentBytes);
            entries['custom/main.xml'] = entries['word/document.xml']!;
            entries['custom/_rels/main.xml.rels'] = entries['word/_rels/document.xml.rels']!;
            delete entries['word/document.xml'];
            delete entries['word/_rels/document.xml.rels'];
            entries['_rels/.rels'] = strToU8(
              strFromU8(entries['_rels/.rels']!).replace('word/document.xml', 'custom/main.xml')
            );
            entries['[Content_Types].xml'] = strToU8(
              strFromU8(entries['[Content_Types].xml']!).replace(
                '/word/document.xml',
                '/custom/main.xml'
              )
            );
            documentBytes = zipSync(entries);
          }
          const editor = createDocxEditor({
            container: document.createElement('div'),
            document: documentBytes,
            imageDecodePort: decode,
          });
          const wrap = (image: string, changed = false) =>
            ownership === 'plain'
              ? image
              : `[${ownership === 'code-linked' ? `<code>${image}</code>` : image}](${changed ? 'https://changed.test' : 'https://outer.test'} "Read")`;
          const seed = `Keep. ${wrap('![Cat](asset:cat "Photo")')} End.`;
          const reference = mode === 'metadata' ? 'asset:cat' : 'asset:dog',
            alt = mode === 'resource' ? 'Cat' : 'Dog',
            title = mode === 'resource' ? 'Photo' : mode === 'metadata' ? null : 'Changed';
          const target = `Keep${text ? ' AI' : ''}. ${wrap(`![${alt}](${reference}${title === null ? '' : ` "${title}"`})`, mode === 'both')} End.`;
          try {
            expect((await apply(editor, seed)).resolved).toEqual(['asset:cat']);
            const initial = editor.surface!.session.currentPackage(),
              before = editor.surface!.session.part();
            const block = [...buildBlockIndex(initial).byId.values()][0]!,
              control = findNode(before, block.controlId)!;
            const drawing = descendants(control).find((node) => node.kind === 'drawing')!,
              paragraphId = editor.surface!.session.paragraphIds()[0]!;
            const link = descendants(before.root).find((node) => node.kind === 'hyperlink');
            const { planned, resolved } = await apply(editor, target);
            expect(resolved).toEqual(mode === 'metadata' ? [] : ['asset:dog']);
            expect(planned.operations.filter((edit) => edit.kind === 'inline-image')).toHaveLength(
              1
            );
            expect(planned.operations.some((edit) => edit.kind === 'inline-atoms')).toBe(false);
            const current = editor.surface!.session.currentPackage(),
              index = buildBlockIndex(current);
            expect(index.diagnostics).toEqual([]);
            expect(index.byId.size).toBe(1);
            const updated = index.byId.get(block.id)!;
            expect(updated.controlId).toBe(block.controlId);
            expect(findNode(editor.surface!.session.part(), drawing.id)?.kind).toBe('drawing');
            expect(updated.payload.attributes.reference).toBe(reference);
            expect(updated.payload.attributes.alt).toBe(alt);
            expect(updated.payload.attributes.title).toBe(title);
            expect(updated.payload.attributes.widthPoints).toBe(
              block.payload.attributes.widthPoints
            );
            expect(updated.payload.attributes.heightPoints).toBe(
              block.payload.attributes.heightPoints
            );
            const mediaName = updated.payload.attributes.mediaPartName as string;
            expect(
              current.partBytes.get(mediaName) ?? current.partBytes.get(mediaName.slice(1))
            ).toEqual(mode === 'metadata' ? png : replacement);
            expect(media(current)).toHaveLength(1);
            if (mode === 'metadata') expect(media(current)).toEqual(media(initial));
            const embed = descendants(
              findNode(editor.surface!.session.part(), drawing.id)!
            ).flatMap((node) =>
              node.kind === 'textValue'
                ? []
                : node.attributes.filter((a) => a.localName === 'embed').map((a) => a.value)
            );
            expect(embed).toHaveLength(1);
            const resource = resolveImageRelationship(
              current.relationships.get(current.mainDocumentPart),
              current.mainDocumentPart,
              embed[0]!
            );
            expect(resource.mode).toBe('internal');
            if (resource.mode !== 'internal') throw Error('Missing native media');
            expect(resource.partName).toBe(mediaName);
            // Independent native Word operations define exactly the allowed subtree changes.
            let expected = before;
            if (text)
              expected = op(expected, { op: 'insertText', paragraphId, offset: 4, text: ' AI' });
            expected = op(expected, {
              op: 'setContentControlProperties',
              controlId: block.controlId,
              lock: 'unlocked',
            });
            if (mode !== 'metadata')
              expected = op(expected, {
                op: 'replaceDrawingResource',
                drawingNodeId: drawing.id,
                relationshipId: embed[0]!,
              });
            if (mode !== 'resource')
              expected = op(expected, {
                op: 'setDrawingMetadata',
                drawingNodeId: drawing.id,
                title: title ?? '',
                description: alt,
              });
            expected = op(expected, {
              op: 'setContentControlProperties',
              controlId: block.controlId,
              lock: 'sdtContentLocked',
            });
            if (mode === 'both' && link) {
              const relationship = current.relationships
                .get(current.mainDocumentPart)!
                .find((r) => r.rawTarget === 'https://changed.test')!;
              expected = op(expected, {
                op: 'setHyperlinkTarget',
                linkId: link.id,
                relationshipId: relationship.id,
                tooltip: 'Read',
              });
            }
            const word = canonicalOoxmlFingerprint(expected);
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(word);
            expect((await plan(editor, target)).operations).toEqual([]);
            expect(editor.exec({ type: 'undo' }).ok).toBe(true);
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
              canonicalOoxmlFingerprint(before)
            );
            expect(media(editor.surface!.session.currentPackage())).toEqual(media(initial));
            expect(editor.exec({ type: 'redo' }).ok).toBe(true);
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(word);
            const saved = new Uint8Array(await editor.save()),
              loaded = readOoxmlPackage(saved);
            if (!loaded.ok) throw Error(loaded.reason);
            expect(
              canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
            ).toBe(word);
            expect(media(loaded.package)).toEqual(media(current));
            const reopened = createDocxEditor({
              container: document.createElement('div'),
              document: saved,
              imageDecodePort: decode,
            });
            try {
              expect((await plan(reopened, target)).operations).toEqual([]);
              expect(
                buildBlockIndex(reopened.surface!.session.currentPackage()).byId.get(block.id)
                  ?.payload
              ).toEqual(updated.payload);
            } finally {
              reopened.destroy();
            }
          } finally {
            editor.destroy();
          }
        });

test('byte-identical replacement updates provenance without replacing native media or geometry', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    imageDecodePort: decode,
    document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
  });
  try {
    await apply(editor, 'Keep. ![Cat](asset:cat "Photo") End.');
    const before = editor.surface!.session.currentPackage();
    const word = canonicalOoxmlFingerprint(editor.surface!.session.part());
    const block = [...buildBlockIndex(before).byId.values()][0]!;
    const target = 'Keep. ![Cat](asset:alias "Photo") End.';
    const planned = await plan(editor, target);
    const prepared = await prepareInlineImageResources(
      planned,
      async () => ({
        bytes: png,
        mime: 'image/png',
        widthPoints: 400,
        heightPoints: 300,
      }),
      decode
    );
    if (!prepared.success) throw Error(prepared.error);
    expect(applyDirectPlan(editor, planned, { ...context, images: prepared.images }).success).toBe(
      true
    );
    const after = editor.surface!.session.currentPackage();
    const updated = buildBlockIndex(after).byId.get(block.id)!;
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(word);
    expect(after.relationships).toEqual(before.relationships);
    expect(media(after)).toEqual(media(before));
    expect(updated.payload.attributes).toEqual({
      ...block.payload.attributes,
      reference: 'asset:alias',
    });
    expect((await plan(editor, target)).operations).toEqual([]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(
      buildBlockIndex(editor.surface!.session.currentPackage()).byId.get(block.id)?.payload
    ).toEqual(block.payload);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const loaded = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!loaded.ok) throw Error(loaded.reason);
    expect(media(loaded.package)).toEqual(media(before));
    expect(buildBlockIndex(loaded.package).byId.get(block.id)?.payload).toEqual(updated.payload);
    expect(
      canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
    ).toBe(word);
  } finally {
    editor.destroy();
  }
});

test('image replacement refuses missing preparation, another plan handle, and stale native revisions atomically', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    imageDecodePort: decode,
    document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
  });
  try {
    await apply(editor, 'Keep. ![Cat](asset:cat "Photo") End.');
    const target = 'Keep. ![Dog](asset:dog "Changed") End.';
    const planned = await plan(editor, target);
    const other = await plan(editor, target.replace('Changed', 'Other'));
    const prepared = await prepareInlineImageResources(
      other,
      async () => ({ bytes: replacement, mime: 'image/png' }),
      decode
    );
    if (!prepared.success) throw Error(prepared.error);
    const before = editor.surface!.session.currentPackage();
    for (const images of [undefined, prepared.images]) {
      const result = applyDirectPlan(editor, planned, { ...context, images });
      expect(result.success).toBe(false);
      expect(editor.surface!.session.currentPackage()).toBe(before);
    }
    const paragraphId = editor.surface!.session.paragraphIds()[0]!;
    expect(
      editor.surface!.applyAutomationOps(
        () => [{ op: 'insertText', paragraphId, offset: 0, text: 'Peer ' }],
        { kind: 'body' },
        [],
        false,
        true
      ).rejected
    ).toBe(false);
    const changed = editor.surface!.session.currentPackage();
    expect(applyDirectPlan(editor, planned, context).success).toBe(false);
    expect(editor.surface!.session.currentPackage()).toBe(changed);
  } finally {
    editor.destroy();
  }
});

test('image plan codec rejects malformed resource intentions and removes exact no-ops', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    imageDecodePort: decode,
    document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
  });
  try {
    await apply(editor, 'Keep. ![Cat](asset:cat "Photo") End.');
    const planned = await plan(editor, 'Keep. ![Dog](asset:dog "Changed") End.');
    const raw = serializeEditPlan(planned);
    for (const patch of [
      { resource: { reference: '', alt: 'Dog', title: null } },
      { resource: { reference: 'asset:dog', alt: '\u0000', title: null } },
      { resource: { reference: 'asset:dog', alt: 'Dog', title: '\ud800' } },
      { resource: { reference: 'asset:dog', alt: 'Dog', title: null, widthPoints: 99 } },
      { blockId: 'wrong-owner' },
      { drawingNodeId: '' },
    ]) {
      const value = JSON.parse(raw);
      Object.assign(
        value.operations.find((op: { kind: string }) => op.kind === 'inline-image'),
        patch
      );
      expect(decodeEditPlan(JSON.stringify(value)).success).toBe(false);
    }
    const value = JSON.parse(raw);
    const image = value.operations.find((op: { kind: string }) => op.kind === 'inline-image');
    const { reference, alt, title } = image.before.attributes;
    image.resource = { reference, alt, title };
    const decoded = decodeEditPlan(JSON.stringify(value));
    expect(decoded.success).toBe(true);
    if (decoded.success) expect(decoded.plan.operations).toEqual([]);
  } finally {
    editor.destroy();
  }
});
