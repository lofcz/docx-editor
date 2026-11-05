import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  contentControlPropertiesOf,
  contentControlContentNodeOf,
  contentControlsIn,
  createNodeIdAllocator,
  escapeXml,
  findNode,
  parentNodeOf,
  readOoxmlPart,
  replaceNode,
  withPart,
  WML_NAMESPACE_URI,
  usedParaIds,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { captureReviewGroupHistory, reviewGroupHistoryRefusal } from '../review/group-history';
import { groupRemovalBudget } from '../review/group-removal-budget';
import { groupReviewFieldBudget } from '../review/group-field-budget';
import { ancestorContentWriteRefusal } from './ancestor-content-write';
import { blockTag, readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import { buildBlockIndex } from './block-index';
import { withFreshIds } from './fresh-ids';
import {
  hasNativeGalleryCaptionStyle,
  isNativeGalleryCaption,
  nativeGalleryPicture,
} from './gallery-native-record';
import { NATIVE_GALLERY_RESOURCE } from './gallery-native-shape';
import { readGallery } from './gallery';
import { readManagedImage } from './inline-image-update';
import { emptyNativeGalleryCaption } from './gallery-native-caption';

const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

export interface AdjacentGalleryImage {
  readonly drawingNodeId: string;
  readonly paragraphId: string;
  readonly captionId: string | null;
}

/** Read a complete adjacent run. Native Caption paragraphs remain attached to their pictures. */
export function readAdjacentGalleryImages(
  pkg: OoxmlPackage,
  partName: string,
  drawingNodeId: string
): readonly AdjacentGalleryImage[] {
  const part = pkg.parts.get(partName);
  const drawing = part && findNode(part, drawingNodeId);
  if (!part || drawing?.kind !== 'drawing') return [];
  let paragraph: OoxmlNode | null = drawing;
  while (paragraph && paragraph.kind !== 'paragraph') paragraph = parentNodeOf(part, paragraph.id);
  if (!paragraph) return [];
  const parent = parentNodeOf(part, paragraph.id);
  if (!parent || !['body', 'contentControlContent', 'tableCell'].includes(parent.kind)) return [];
  if (ancestorContentWriteRefusal(part, paragraph.id)) return [];
  const index = buildBlockIndex(pkg);
  // Existing registered galleries and other atomic hosts keep their own mutation boundary.
  for (
    let ancestor: OoxmlNode | null = parent;
    ancestor;
    ancestor = parentNodeOf(part, ancestor.id)
  ) {
    if (ancestor.kind !== 'contentControl') continue;
    const tag = contentControlPropertiesOf(ancestor).tag;
    if (
      [...index.byId.values()].some(
        (block) =>
          blockTag(block.id) === tag &&
          ['image_gallery', 'image_gallery_item'].includes(block.payload.type)
      )
    )
      return [];
  }
  const picture = (node: OoxmlNode) => {
    const native = nativeGalleryPicture(pkg, part, node);
    if (!native) return null;
    const controls = contentControlsIn(node);
    if (controls.length) {
      const managed = readManagedImage(pkg, part.name, native.fields.drawing.id);
      if (controls.length !== 1 || !managed || managed.control.id !== controls[0]!.node.id)
        return null;
      const properties = contentControlPropertiesOf(managed.control);
      if (
        properties.dataBinding ||
        (properties.lock && !['unlocked', 'sdtContentLocked'].includes(properties.lock))
      )
        return null;
    }
    return native;
  };
  const runs: AdjacentGalleryImage[][] = [];
  let run: AdjacentGalleryImage[] = [];
  for (let at = 0; at < parent.children.length; at++) {
    const node = parent.children[at]!;
    const image = picture(node);
    if (!image) {
      if (run.length) runs.push(run);
      run = [];
      continue;
    }
    const following = parent.children[at + 1];
    const caption = following && isNativeGalleryCaption(part, following) ? following : null;
    const nestedCaption = following && contentControlContentNodeOf(following);
    if (
      following &&
      !caption &&
      (hasNativeGalleryCaptionStyle(part, following) ||
        nestedCaption?.children.some((child) => hasNativeGalleryCaptionStyle(part, child)))
    ) {
      if (run.length) runs.push(run);
      run = [];
      continue;
    }
    run.push({
      drawingNodeId: image.fields.drawing.id,
      paragraphId: node.id,
      captionId: caption?.id ?? null,
    });
    if (caption) at++;
  }
  if (run.length) runs.push(run);
  return runs.find((images) => images.some((image) => image.drawingNodeId === drawingNodeId)) ?? [];
}

/** Build new ownership wrappers around unchanged native descendants. */
export function nativeGalleryControl(
  part: OoxmlPart,
  id: string,
  children: readonly OoxmlNode[],
  allocate: () => string
): OoxmlElement {
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body><w:sdt><w:sdtPr><w:tag w:val="${escapeXml(blockTag(id))}"/></w:sdtPr><w:sdtContent/></w:sdt></w:body></w:document>`,
    { name: '/sciobot-native-gallery.xml', contentType: part.contentType }
  );
  const body = parsed.ok && bodyStoryRoot(parsed.part);
  const template = body && body.kind !== 'textValue' && body.children[0];
  if (!template || template.kind !== 'contentControl')
    throw Error('The native gallery wrapper is invalid.');
  const control = withFreshIds(template, allocate, { prefix: null, used: new Set(), seed: id });
  if (control.kind !== 'contentControl') throw Error('The native gallery wrapper is missing.');
  return {
    ...control,
    children: control.children.map((node) =>
      node.kind === 'contentControlContent' ? { ...node, children: [...children] } : node
    ),
  } as OoxmlElement;
}

/** Group adjacent native pictures in one shared transaction without copying their resources or fields. */
export function groupAdjacentImages(
  editor: DocxEditorInstance,
  request: {
    readonly drawingNodeId: string;
    readonly expectedRevision?: number;
  }
): { success: true; revision: number; blockId: string } | NativeEditFailure {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (surface.activeScope().kind !== 'body')
    return fail('unsupported-operation', 'Gallery grouping requires the body story.');
  if (surface.editingMode() !== 'edit')
    return fail('unsupported-operation', 'Gallery changes require edit mode.');
  const revision = surface.session.packageRevision();
  if (request.expectedRevision !== undefined && request.expectedRevision !== revision)
    return fail('stale-revision', 'Read the current images before grouping them.');
  const history = reviewGroupHistoryRefusal(surface.collaborationSession());
  if (history) return fail('unsupported-operation', history);
  const pkg = surface.session.currentPackage(),
    part = surface.session.part();
  const images = readAdjacentGalleryImages(pkg, part.name, request.drawingNodeId);
  if (images.length < 2 || images.length > 100)
    return fail('unsupported-operation', 'Choose two to 100 adjacent native pictures.');
  const parent = parentNodeOf(part, images[0]!.paragraphId)!;
  const blockId = crypto.randomUUID();
  const imageIds = images.map(() => crypto.randomUUID());
  const transform = (working: OoxmlPackage) => {
    const source = working.parts.get(part.name)!;
    let next = working;
    const allocate = createNodeIdAllocator(source, 'new');
    const used = new Set(usedParaIds(source.root));
    const owners = images.map((image, at) => {
      const children = [
        findNode(source, image.paragraphId)!,
        image.captionId
          ? findNode(source, image.captionId)!
          : emptyNativeGalleryCaption(source, imageIds[at]!, allocate, used),
      ];
      const stored = withPayloadRecord(
        next,
        {
          id: imageIds[at]!,
          version: 3,
          type: 'image_gallery_item',
          galleryId: blockId,
          attributes: { resource: NATIVE_GALLERY_RESOURCE },
        },
        'Image'
      );
      if (!stored.ok) throw Error(stored.reason);
      next = stored.pkg;
      return nativeGalleryControl(source, imageIds[at]!, children, allocate);
    });
    const stored = withPayloadRecord(
      next,
      { id: blockId, type: 'image_gallery', version: 3, attributes: { layout: 'row-fit' } },
      'Gallery'
    );
    if (!stored.ok) throw Error(stored.reason);
    next = stored.pkg;
    const control = nativeGalleryControl(source, blockId, owners, allocate);
    const selected = new Set(
      images.flatMap((image) => [image.paragraphId, ...(image.captionId ? [image.captionId] : [])])
    );
    const replaced = replaceNode(source, parent.id, {
      ...parent,
      children: parent.children.flatMap((node) =>
        node.id === images[0]!.paragraphId ? [control] : selected.has(node.id) ? [] : [node]
      ),
    } as OoxmlElement);
    if (!replaced.ok) throw Error('The image paragraphs cannot be grouped.');
    return withPart(next, replaced.part);
  };
  try {
    const candidate = transform(pkg);
    if (!readGallery(candidate, blockId) || readPayloadStore(candidate).status !== 'ready')
      throw Error('The gallery has invalid native image ownership.');
    const result = captureReviewGroupHistory(
      surface.collaborationSession(),
      () => surface.applyAutomationOps(() => [], { kind: 'body' }, [transform], false, true),
      groupRemovalBudget(pkg, candidate),
      groupReviewFieldBudget(pkg, candidate)
    );
    if (result.rejected || !result.committed)
      return fail('apply-failed', String(result.reason ?? 'The gallery did not change.'));
    return { success: true, revision: surface.session.packageRevision(), blockId };
  } catch (cause) {
    return fail('apply-failed', cause instanceof Error ? cause.message : String(cause));
  }
}
