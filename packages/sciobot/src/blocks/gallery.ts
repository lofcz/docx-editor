import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  contentControlContentNodeOf,
  contentControlPropertiesOf,
  createNodeIdAllocator,
  findNode,
  insertChildren,
  lockForbidsEdit,
  ATOMIC_CONTENT_NAMESPACE,
  replaceNode,
  withPart,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import type { JsonObject } from '../data/json';
import { galleryImage } from '../projection/gallery';
import { buildBlockIndex } from './block-index';
import { galleryImageControl, galleryImageText } from './gallery-record-build';
import { GALLERY_IMAGE_RECORD } from './gallery-record-shape';
import { moveBlock } from './move';
import { blockIdFromTag, readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import { isBlockId, type BlockPayload } from './payload-schema';
import { removeBlock } from './remove';
import { ancestorContentWriteRefusal } from './ancestor-content-write';

export interface GalleryImageState {
  readonly id: string;
  readonly controlId: string;
  readonly attributes: JsonObject;
}
export interface GalleryState {
  readonly id: string;
  readonly controlId: string;
  readonly partName: string;
  readonly payload: BlockPayload;
  readonly images: readonly GalleryImageState[];
}

/** Stable native image identities supplement the unchanged public images array. */
export function readGallery(pkg: OoxmlPackage, blockId: string): GalleryState | null {
  const index = buildBlockIndex(pkg);
  const gallery = index.byId.get(blockId);
  if (!gallery || gallery.payload.type !== 'image_gallery' || gallery.payload.version !== 2)
    return null;
  const part = pkg.parts.get(gallery.partName);
  const control = part && findNode(part, gallery.controlId);
  const content = control && contentControlContentNodeOf(control);
  if (!content) return null;
  const images: GalleryImageState[] = [];
  for (const child of content.children) {
    if (child.kind !== 'contentControl') continue;
    const id = blockIdFromTag(contentControlPropertiesOf(child).tag ?? '');
    const image = id && index.byId.get(id);
    if (
      !image ||
      image.parentId !== blockId ||
      image.payload.type !== GALLERY_IMAGE_RECORD ||
      image.payload.version !== 2 ||
      image.payload.galleryId !== blockId
    )
      return null;
    images.push({ id: image.id, controlId: child.id, attributes: image.payload.attributes });
  }
  return { ...gallery, images };
}

export type GalleryChange =
  | {
      readonly type: 'caption';
      readonly imageId: string;
      readonly caption: string;
      readonly expectedCaption: string;
    }
  | {
      readonly type: 'move';
      readonly imageId: string;
      readonly beforeImageId: string | null;
      readonly expectedOrder: readonly string[];
    }
  | { readonly type: 'remove'; readonly imageId: string }
  | {
      readonly type: 'append';
      readonly images: readonly { readonly id: string; readonly attributes: JsonObject }[];
    }
  | {
      readonly type: 'presentation';
      readonly layout?: 'row-fit' | 'grid';
      readonly align?: 'left' | 'center' | 'right';
    };

export interface GalleryUpdate {
  readonly blockId: string;
  readonly change: GalleryChange;
  readonly expectedRevision?: number;
}
export type GalleryUpdateResult =
  | { readonly success: true; readonly revision: number }
  | NativeEditFailure;
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

/** One image edit, membership change, or native move is one validated shared transaction. */
export function updateGallery(
  editor: DocxEditorInstance,
  request: GalleryUpdate
): GalleryUpdateResult {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (surface.editingMode() !== 'edit')
    return fail('unsupported-operation', 'Gallery changes require edit mode.');
  if (
    request.expectedRevision !== undefined &&
    request.expectedRevision !== surface.session.packageRevision()
  )
    return fail('stale-revision', 'Read the current gallery before changing it.');
  const pkg = surface.session.currentPackage();
  const gallery = readGallery(pkg, request.blockId);
  if (!gallery)
    return fail('invalid-input', 'The gallery is missing or needs preparation before sharing.');
  if (gallery.partName !== surface.session.part().name)
    return fail('unsupported-operation', 'The gallery is outside the active story.');
  const change = request.change;
  const item =
    'imageId' in change ? gallery.images.find((image) => image.id === change.imageId) : undefined;
  if ('imageId' in change && !item)
    return fail('stale-revision', 'The image no longer belongs to this gallery.');
  if (change.type === 'remove')
    return removeBlock(editor, {
      controlId: gallery.images.length === 1 ? gallery.controlId : item!.controlId,
      mode: 'delete',
    });
  if (change.type === 'move') {
    if (
      JSON.stringify(change.expectedOrder) !==
      JSON.stringify(gallery.images.map((image) => image.id))
    )
      return fail('stale-revision', 'The gallery order changed while this move was open.');
    const before =
      change.beforeImageId === null
        ? null
        : gallery.images.find((image) => image.id === change.beforeImageId);
    if (before === undefined)
      return fail('stale-revision', 'The movement target no longer belongs to this gallery.');
    return moveBlock(editor, {
      controlId: item!.controlId,
      parentControlId: gallery.controlId,
      beforeId: before?.controlId ?? null,
    });
  }
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready')
    return fail('invalid-document', 'The gallery payload store is unavailable.');
  let transform: (working: OoxmlPackage) => OoxmlPackage;
  const part = surface.session.part();
  const target = findNode(part, item?.controlId ?? gallery.controlId)!;
  // The validated image owner permits its atomic fallback write, never an authored Word lock.
  const declared = contentControlPropertiesOf(
    item && target.kind === 'contentControl'
      ? {
          ...target,
          attributes: target.attributes.filter(
            (attribute) =>
              attribute.namespaceUri !== ATOMIC_CONTENT_NAMESPACE ||
              attribute.localName !== 'atomicContent'
          ),
        }
      : target
  );
  const protectedAncestor = ancestorContentWriteRefusal(part, target.id);
  if (lockForbidsEdit(declared.lock) || declared.dataBinding || protectedAncestor)
    return fail(
      'unsupported-operation',
      protectedAncestor ?? 'The gallery content is protected by a native control.'
    );
  if (change.type === 'caption') {
    if (typeof change.caption !== 'string' || change.caption.length > 2048)
      return fail('invalid-input', 'The image caption is too long.');
    if ((item!.attributes.caption ?? '') !== change.expectedCaption)
      return fail('stale-revision', 'The image caption changed while this draft was open.');
    const previous = store.entries.get(item!.id)!;
    if (previous.decoded.status !== 'supported')
      return fail('invalid-document', 'The image record is unavailable.');
    const { caption: _caption, ...attributes } = item!.attributes;
    const payload: BlockPayload = {
      ...previous.decoded.value,
      attributes: {
        ...attributes,
        ...(change.caption ? { caption: change.caption } : {}),
      },
    };
    transform = (working) => {
      let part = working.parts.get(gallery.partName)!;
      const control = findNode(part, item!.controlId);
      if (!control) throw Error('The image fallback disappeared.');
      const texts: { id: string; kind: 'textValue'; value: string }[] = [];
      const visit = (node: OoxmlNode) => {
        if (node.kind === 'textValue') texts.push(node);
        else node.children.forEach(visit);
      };
      const content = contentControlContentNodeOf(control);
      if (!content) throw Error('The image fallback is missing.');
      visit(content);
      if (!texts.length) throw Error('The image fallback text is missing.');
      for (const [at, text] of texts.entries()) {
        const changed = replaceNode(part, text.id, {
          ...text,
          value: at === 0 ? galleryImageText(payload.attributes) : '',
        });
        if (!changed.ok) throw Error('The image caption could not change.');
        part = changed.part;
      }
      const stored = withPayloadRecord(withPart(working, part), payload, previous.label);
      if (!stored.ok) throw Error(stored.reason);
      return stored.pkg;
    };
  } else if (change.type === 'presentation') {
    const { images: _images, ...attributes } = gallery.payload.attributes;
    const payload: BlockPayload = {
      ...gallery.payload,
      attributes: {
        ...attributes,
        ...(change.layout === undefined ? {} : { layout: change.layout }),
        ...(change.align === undefined ? {} : { align: change.align }),
      },
    };
    transform = (working) => {
      const stored = withPayloadRecord(working, payload, store.entries.get(gallery.id)!.label);
      if (!stored.ok) throw Error(stored.reason);
      return stored.pkg;
    };
  } else {
    if (
      !change.images.length ||
      gallery.images.length + change.images.length > 100 ||
      new Set(change.images.map((image) => image.id)).size !== change.images.length ||
      change.images.some(
        (image) =>
          !isBlockId(image.id) || store.entries.has(image.id) || !galleryImage(image.attributes)
      )
    )
      return fail(
        'invalid-input',
        'Add one to 100 images with distinct unused IDs and valid addresses.'
      );
    transform = (working) => {
      let next = working;
      const part = working.parts.get(gallery.partName)!;
      const control = findNode(part, gallery.controlId)!;
      const content = contentControlContentNodeOf(control)!;
      const allocate = createNodeIdAllocator(part, 'new');
      const children = change.images.map((image) => {
        const stored = withPayloadRecord(
          next,
          {
            id: image.id,
            type: GALLERY_IMAGE_RECORD,
            version: 2,
            galleryId: gallery.id,
            attributes: image.attributes,
          },
          galleryImageText(image.attributes)
        );
        if (!stored.ok) throw Error(stored.reason);
        next = stored.pkg;
        return galleryImageControl(part, image.id, image.attributes, undefined, allocate);
      });
      const inserted = insertChildren(part, content.id, content.children.length, children);
      if (!inserted.ok) throw Error('The native gallery image insertion failed.');
      return withPart(next, inserted.part);
    };
  }
  try {
    const candidate = transform(pkg);
    if (!readGallery(candidate, gallery.id))
      throw Error('The gallery update produced invalid image ownership.');
    const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [transform], false, true);
    if (result.rejected || !result.committed)
      return fail('apply-failed', String(result.reason ?? 'The gallery update did not commit.'));
    return { success: true, revision: surface.session.packageRevision() };
  } catch (cause) {
    return fail('apply-failed', cause instanceof Error ? cause.message : String(cause));
  }
}
