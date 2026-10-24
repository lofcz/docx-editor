import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  cleanupOrphanImageMedia,
  findNode,
  replaceNode,
  withPart,
  type ImageDecodePort,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from './document';
import { planTargetDocument } from '../../src/diff/plan-target';
import { previewDirectPlan } from '../../src/ai/apply-direct';
import { readWorkingView } from '../../src/review/working-view';
import { readEditorRevision } from '../../src/projection/revision-cache';
import { prepareInlineImageResources } from '../../src/native/inline-image-resources';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { withPayloadRecord } from '../../src/blocks/payload-opc-codec';
export const imageRemovalContext = {
  documentId: 'tracked-image-removal',
  epoch: 'native',
  author: 'AI',
  date: '2026-10-24T12:00:00.000Z',
};
export const imageRemovalPng = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
export const imageRemovalDecodePort: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
export const imageRemovalMedia = (pkg: OoxmlPackage) =>
  [...pkg.partBytes]
    .filter(([name]) => /(^|\/)media\//.test(name))
    .sort(([a], [b]) => a.localeCompare(b));
export const imageRemovalNodes = (node: OoxmlNode): OoxmlNode[] => [
  node,
  ...(node.kind === 'textValue' ? [] : node.children.flatMap(imageRemovalNodes)),
];
export async function imageRemovalFixture(linked: boolean, shared: boolean): Promise<OoxmlPackage> {
  const seed = createDocxEditor({
    container: document.createElement('div'),
    imageDecodePort: imageRemovalDecodePort,
    document: docxFixture('<w:p><w:r><w:t>Lead.  Tail.</w:t></w:r></w:p>'),
  });
  try {
    const revision = readEditorRevision(
      seed,
      imageRemovalContext.documentId,
      imageRemovalContext.epoch
    )!;
    const source = seed.surface!.session.currentPackage();
    const view = await readWorkingView(seed.surface!.session, revision);
    if (!view.success) throw Error(view.error);
    const cat = '![Cat](asset:cat "Photo")';
    const target = `Lead. ${linked ? `[${cat}](https://keep.test "Keep")` : cat} / ![Dog](asset:dog "Other") Tail.`;
    const planned = planTargetDocument(
      { pkg: source, revision, projection: view.view.projection },
      target
    );
    if (!planned.success) throw Error(planned.error);
    const resources = await prepareInlineImageResources(
      planned.plan,
      async (reference) => ({
        bytes: imageRemovalPng,
        mime: 'image/png',
        widthPoints: reference === 'asset:cat' ? 24 : 32,
        heightPoints: reference === 'asset:cat' ? 18 : 20,
      }),
      imageRemovalDecodePort
    );
    if (!resources.success) throw Error(resources.error);
    const lowered = previewDirectPlan(source, planned.plan, resources.images);
    if (!lowered.success) throw Error(lowered.error);
    if (!shared) return lowered.pkg;
    const blocks = [...buildBlockIndex(lowered.pkg).byId.values()];
    const catBlock = blocks.find((block) => block.payload.attributes.reference === 'asset:cat')!;
    const dogBlock = blocks.find((block) => block.payload.attributes.reference === 'asset:dog')!;
    const part = lowered.pkg.parts.get(dogBlock.partName)!;
    const blip = (id: string) =>
      imageRemovalNodes(findNode(part, id)!).find(
        (node) => node.kind !== 'textValue' && node.localName === 'blip'
      )!;
    const first = blip(catBlock.controlId),
      second = blip(dogBlock.controlId);
    if (first.kind !== 'pictureBlip' || second.kind !== 'pictureBlip')
      throw Error('Missing image blip.');
    const embed = first.attributes.find((attr) => attr.localName === 'embed')!.value;
    const previous = second.attributes.find((attr) => attr.localName === 'embed')!.value;
    const replaced = replaceNode(part, second.id, {
      ...second,
      attributes: second.attributes.map((attr) =>
        attr.localName === 'embed' ? { ...attr, value: embed } : attr
      ),
    });
    if (!replaced.ok) throw Error('Cannot prepare shared image reference.');
    const stored = withPayloadRecord(
      withPart(lowered.pkg, replaced.part),
      {
        ...dogBlock.payload,
        attributes: {
          ...dogBlock.payload.attributes,
          mediaPartName: catBlock.payload.attributes.mediaPartName!,
        },
      },
      'Image'
    );
    if (!stored.ok) throw Error(stored.reason);
    return cleanupOrphanImageMedia(
      stored.pkg,
      part.name,
      String(dogBlock.payload.attributes.mediaPartName),
      previous
    );
  } finally {
    seed.destroy();
  }
}
