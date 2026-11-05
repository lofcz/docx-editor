import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  bodyStoryRoot,
  contentControlsIn,
  createNodeIdAllocator,
  escapeXml,
  findNode,
  insertChildren,
  parentNodeOf,
  readOoxmlPart,
  usedParaIds,
  validateTreeOp,
  w14RootPrefix,
  withoutCustomXmlNode,
  withPart,
  WML_NAMESPACE_URI,
  type ImageDecodePort,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { embedInlineImage } from '../native/inline-image-media';
import { withInlineControlId } from '../native/inline-image-control-id';
import { inlinePicture } from '../native/inline-picture';
import {
  resolveInlineImageResources,
  type InlineImageResolver,
} from '../native/inline-image-resources';
import { activeMdxRegistry } from '../registry/active';
import { withFreshIds } from './fresh-ids';
import { readGallery } from './gallery';
import { createInlineImagePayload } from './inline-image-codec';
import { blockTag, readPayloadStore, withPayloadRecord } from './payload-opc-codec';

const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

/** Explicitly fetch gallery assets, then replace the gallery with native images in one transaction. */
export async function ungroupGallery(
  editor: DocxEditorInstance,
  request: {
    blockId: string;
    resolve: InlineImageResolver;
    decodePort: ImageDecodePort;
    signal?: AbortSignal;
  }
): Promise<{ success: true; revision: number } | NativeEditFailure> {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (surface.editingMode() !== 'edit')
    return fail('unsupported-operation', 'Gallery changes require edit mode.');
  const session = surface.session;
  const revision = session.packageRevision();
  const pkg = session.currentPackage();
  const gallery = readGallery(pkg, request.blockId);
  if (!gallery || gallery.partName !== session.part().name)
    return fail('invalid-input', 'The gallery is missing from the active story.');
  const remove = {
    op: 'removeContentControl' as const,
    controlId: gallery.controlId,
    keepContent: false,
  };
  const refusal = validateTreeOp(session.part(), remove);
  if (refusal) return fail('unsupported-operation', `The gallery cannot be ungrouped: ${refusal}.`);
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready')
    return fail('invalid-document', 'The gallery metadata is unavailable.');
  try {
    const resources = await resolveInlineImageResources(
      gallery.images.map((image) => String(image.attributes.url)),
      request.resolve,
      request.decodePort,
      request.signal
    );
    if (request.signal?.aborted)
      return fail('stale-revision', 'The gallery operation was canceled.');
    if (
      editor.surface !== surface ||
      surface.session !== session ||
      session.packageRevision() !== revision ||
      surface.editingMode() !== 'edit'
    )
      return fail(
        'stale-revision',
        'The document changed while the images loaded. Open the gallery control again.'
      );
    const imageIds = gallery.images.map(() => crypto.randomUUID());
    const transform = (working: OoxmlPackage) => {
      let next = working;
      const records: ReturnType<typeof createInlineImagePayload>[] = [];
      const used = new Set(usedParaIds(session.part().root));
      for (const [index, image] of gallery.images.entries()) {
        const part = next.parts.get(gallery.partName)!;
        const control = findNode(part, gallery.controlId)!;
        const parent = parentNodeOf(part, control.id);
        if (!parent || parent.kind === 'paragraph')
          throw Error('The gallery has no block container.');
        const reference = String(image.attributes.url);
        const resource = resources.get(reference)!;
        const widthPoints =
          typeof image.attributes.width === 'number'
            ? image.attributes.width * 0.75
            : resource.widthPoints;
        const heightPoints =
          typeof image.attributes.height === 'number'
            ? image.attributes.height * 0.75
            : resource.heightPoints;
        if (
          ![widthPoints, heightPoints].every(
            (size) => Number.isFinite(size) && Math.round(size * 12700) > 0 && size <= 158400
          )
        )
          throw Error('The image dimensions exceed the native drawing bounds.');
        const id = imageIds[index]!;
        const embedded = embedInlineImage(next, gallery.partName, id, {
          ...resource,
          widthPoints,
          heightPoints,
        });
        next = embedded.pkg;
        const alt = String(image.attributes.caption ?? '').replace(/\r\n|\r/g, '\n');
        const drawing = inlinePicture({
          docPrId: embedded.docPrId,
          relationshipId: embedded.relationshipId,
          widthPoints,
          heightPoints,
          title: null,
          alt,
        });
        const align = String(gallery.payload.attributes.align ?? 'left');
        const caption = alt
          ? `<w:p><w:pPr><w:pStyle w:val="Caption"/><w:jc w:val="${escapeXml(align)}"/></w:pPr><w:r>${alt
              .split('\n')
              .map((line) => `<w:t xml:space="preserve">${escapeXml(line)}</w:t>`)
              .join('<w:br/>')}</w:r></w:p>`
          : '';
        const parsed = readOoxmlPart(
          `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body><w:p><w:pPr><w:jc w:val="${escapeXml(align)}"/></w:pPr><w:sdt><w:sdtPr><w:tag w:val="${escapeXml(blockTag(id))}"/><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent><w:r><w:t>\uFFFC</w:t></w:r></w:sdtContent></w:sdt></w:p>${caption}</w:body></w:document>`,
          { name: '/sciobot-gallery-ungroup.xml', contentType: part.contentType }
        );
        const body = parsed.ok && bodyStoryRoot(parsed.part);
        const template = body && body.kind !== 'textValue' && body.children[0];
        if (!template || template.kind !== 'paragraph')
          throw Error('The native image paragraph is invalid.');
        const replacePicture = (node: OoxmlNode): OoxmlNode => {
          if (node.kind === 'text') return drawing;
          if (node.kind === 'textValue') return node;
          return { ...node, children: node.children.map(replacePicture) } as OoxmlNode;
        };
        const allocate = createNodeIdAllocator(part, 'new');
        const identity = { prefix: w14RootPrefix(part.root), used, seed: `ungroup:${id}` };
        const paragraph = withFreshIds(replacePicture(template), allocate, identity);
        const paragraphs = [paragraph];
        if (body && body.children[1])
          paragraphs.push(withFreshIds(body.children[1], allocate, identity));
        const inserted = insertChildren(
          part,
          parent.id,
          parent.children.findIndex((child) => child.id === control.id),
          paragraphs
        );
        if (!inserted.ok) throw Error('The native image cannot replace the gallery.');
        next = withPart(next, inserted.part);
        const imageControl = contentControlsIn(paragraph)[0]?.node;
        if (!imageControl) throw Error('The native image control is missing.');
        next = withInlineControlId(next, gallery.partName, imageControl.id);
        const payload = createInlineImagePayload(id, {
          reference,
          alt,
          title: null,
          contentHash: resource.contentHash,
          mediaPartName: embedded.partName.startsWith('/')
            ? embedded.partName
            : '/' + embedded.partName,
          widthPoints,
          heightPoints,
        });
        records.push(payload);
      }
      const removed = applyTreeOp(next.parts.get(gallery.partName)!, remove);
      if (!removed.ok) throw Error(`The gallery removal failed: ${removed.reason}.`);
      next = withPart(next, removed.part);
      for (const id of [gallery.id, ...gallery.images.map((image) => image.id)])
        next = withoutCustomXmlNode(
          next,
          store.part.partName,
          activeMdxRegistry().storage.nodePrefix + id
        );
      for (const payload of records) {
        const saved = withPayloadRecord(next, payload, String(payload.attributes.alt || 'Image'));
        if (!saved.ok) throw Error(saved.reason);
        next = saved.pkg;
      }
      if (readPayloadStore(next).status !== 'ready')
        throw Error('The native image metadata is invalid.');
      return next;
    };
    transform(pkg);
    const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [transform], false, true);
    if (result.rejected || !result.committed)
      return fail('apply-failed', String(result.reason ?? 'The gallery did not change.'));
    return { success: true, revision: session.packageRevision() };
  } catch (error) {
    return fail('apply-failed', error instanceof Error ? error.message : String(error));
  }
}
