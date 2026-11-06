import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  contentControlPropertiesOf,
  contentControlsIn,
  createNodeIdAllocator,
  findNode,
  freshenCopiedRunOrigins,
  freshenCopiedDrawingPropertyIds,
  insertChildren,
  parentNodeOf,
  usedParaIds,
  w14RootPrefix,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { buildBlockIndex } from './block-index';
import { withFreshIds } from './fresh-ids';
import { blockIdFromTag, blockTag, readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import type { BlockPayload } from './payload-schema';
import { copiedGalleryPayload, galleryRecordOrder } from './gallery-copy';
import { resolveExactBlockSelection } from './selection-exact';
import { blockRangeCopyRefusal } from './selection-copy-safety';
import { ancestorContentWriteRefusal } from './ancestor-content-write';
import { drawingCopyRefusal } from './selection-drawing-copy';

export interface BlockDuplication {
  readonly controlId: string;
  /** A new block ID for each copied block, by payload type. IDs must be unused. */
  readonly allocateBlockId: (type: string) => string;
  readonly expectedRevision?: number;
}
export type BlockDuplicateResult =
  | {
      readonly success: true;
      readonly revision: number;
      readonly controlId: string;
      /** New payload IDs by the IDs they copy. */
      readonly payloadIds: Readonly<Record<string, string>>;
    }
  | NativeEditFailure;
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

/**
 * The copy of a control tree: block tags renamed, control `w:id` values and paragraph `w14`
 * identities dropped (fresh ones are minted on placement), everything else kept.
 */
function copyOf(node: OoxmlNode, tags: ReadonlyMap<string, string>): OoxmlNode {
  if (node.kind === 'textValue') return node;
  let attributes = node.attributes;
  if (node.kind === 'paragraph')
    attributes = attributes.filter(
      (attribute) =>
        attribute.namespaceUri !== W14 ||
        (attribute.localName !== 'paraId' && attribute.localName !== 'textId')
    ) as typeof attributes;
  if (node.namespaceUri === WML_NAMESPACE_URI && node.localName === 'tag')
    attributes = attributes.map((attribute) =>
      attribute.localName === 'val' && tags.has(attribute.value)
        ? { ...attribute, value: tags.get(attribute.value)! }
        : attribute
    ) as typeof attributes;
  const children = node.children
    .filter(
      (child) =>
        !(
          node.namespaceUri === WML_NAMESPACE_URI &&
          node.localName === 'sdtPr' &&
          child.kind !== 'textValue' &&
          child.namespaceUri === WML_NAMESPACE_URI &&
          child.localName === 'id'
        )
    )
    .map((child) => copyOf(child, tags));
  return { ...node, attributes, children } as OoxmlNode;
}

/**
 * Insert a copy of a registered block after it, in one native undo unit. The copy and every
 * block inside it get fresh payload IDs, so it never shares metadata with the original.
 */
export function duplicateBlock(
  editor: DocxEditorInstance,
  request: BlockDuplication
): BlockDuplicateResult {
  const result = duplicateSelection(
    editor,
    {
      nodeIds: [request.controlId],
      allocateBlockId: request.allocateBlockId,
      expectedRevision: request.expectedRevision,
    },
    false
  );
  return result.success
    ? {
        success: true,
        revision: result.revision,
        controlId: result.nodeIds[0]!,
        payloadIds: result.payloadIds,
      }
    : result;
}

export interface BlockSelectionDuplication {
  readonly nodeIds: readonly string[];
  readonly allocateBlockId: (type: string) => string;
  readonly expectedRevision?: number;
}
export type BlockSelectionDuplicateResult =
  | {
      readonly success: true;
      readonly revision: number;
      readonly nodeIds: readonly string[];
      readonly payloadIds: Readonly<Record<string, string>>;
    }
  | NativeEditFailure;

/**
 * Copy selected complete blocks after the last selected root, with independent payload and review identities.
 * Ordinary ranges refuse unsupported drawings, document-wide references, section marks, and pending revisions.
 * A single registered block retains the existing registered-block copy support.
 */
export function duplicateBlockSelection(
  editor: DocxEditorInstance,
  request: BlockSelectionDuplication
): BlockSelectionDuplicateResult {
  return duplicateSelection(editor, request, true);
}

function duplicateSelection(
  editor: DocxEditorInstance,
  request: BlockSelectionDuplication,
  native: boolean
): BlockSelectionDuplicateResult {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  if (native && surface.activeScope().kind !== 'body')
    return fail('unsupported-operation', 'Block selection changes require the document body.');
  surface.flushPendingInput();
  if (
    request.expectedRevision !== undefined &&
    request.expectedRevision !== surface.session.packageRevision()
  )
    return fail('stale-revision', 'Read the current document before duplicating this block.');
  if (surface.editingMode() !== 'edit')
    return fail('unsupported-operation', 'Block duplication requires edit mode.');
  const pkg = surface.session.currentPackage(),
    part = surface.session.part();
  const control = findNode(part, request.nodeIds[0]!);
  const parent = control && parentNodeOf(part, control.id);
  if (
    !native &&
    (!control || control.kind !== 'contentControl' || !parent || parent.kind === 'paragraph')
  )
    return fail('invalid-input', 'The registered block is missing from the active story.');
  const range = resolveExactBlockSelection(part, request.nodeIds);
  if (!range.success) return range;
  const drawingRefusal = drawingCopyRefusal(pkg, part.name, range.nodes);
  if (drawingRefusal) return fail('unsupported-operation', drawingRefusal);
  // Pure preflight validates source identities before the caller allocates payload IDs.
  // Placement repeats allocation inside the actual transaction's actor scope.
  try {
    freshenCopiedDrawingPropertyIds(pkg, range.nodes, 500_000, 64);
  } catch (error) {
    return fail('unsupported-operation', error instanceof Error ? error.message : String(error));
  }
  const index = buildBlockIndex(pkg);
  const selected = new Set(range.nodes.map((node) => node.id));
  const covered = (id: string) => {
    for (let node = findNode(part, id); node; node = parentNodeOf(part, node.id))
      if (selected.has(node.id)) return true;
    return false;
  };
  for (const block of index.byId.values()) {
    if (block.payload.type !== 'image_gallery_item' || !covered(block.controlId)) continue;
    const owner =
      typeof block.payload.galleryId === 'string' && index.byId.get(block.payload.galleryId);
    if (!owner || !covered(owner.controlId))
      return fail('unsupported-operation', 'Copy gallery images together with their gallery.');
  }
  const registeredSingle =
    range.nodes.length === 1 &&
    [...index.byId.values()].some((block) => block.controlId === range.nodes[0]!.id);
  const registeredControls = new Set([...index.byId.values()].map((block) => block.controlId));
  if (
    surface.collaborationSession() &&
    !surface.collaborationSession()?.resolveNodeId &&
    range.nodes.some((node) => !registeredControls.has(node.id))
  )
    return fail(
      'unsupported-operation',
      'This collaboration provider cannot retain copied native block identities.'
    );
  if (native && !registeredSingle) {
    const refusal = blockRangeCopyRefusal(range.nodes);
    if (refusal) return fail('unsupported-operation', refusal);
  }
  if (native)
    for (const node of range.nodes) {
      const refusal = ancestorContentWriteRefusal(part, node.id);
      if (refusal) return fail('unsupported-operation', refusal);
    }
  const store = readPayloadStore(pkg);
  if (!native && store.status !== 'ready')
    return fail('invalid-document', 'The block payload store is unavailable.');
  const tags = new Map<string, string>();
  const payloadIds: Record<string, string> = Object.create(null);
  const copies: { payload: BlockPayload; label: string }[] = [];
  for (const node of range.nodes.flatMap((root) => [
    ...(root.kind === 'contentControl' ? [root] : []),
    ...contentControlsIn(root).map((entry) => entry.node),
  ])) {
    const tag = contentControlPropertiesOf(node).tag ?? '';
    const id = blockIdFromTag(tag);
    if (id === null || Object.hasOwn(payloadIds, id)) continue;
    const block = index.byId.get(id);
    const stored = store.status === 'ready' ? store.entries.get(id) : null;
    if (!block || !stored)
      return fail('invalid-document', 'A copied block has no stored payload record.');
    const next = request.allocateBlockId(block.payload.type);
    if (
      index.byId.has(next) ||
      (store.status === 'ready' && store.entries.has(next)) ||
      Object.values(payloadIds).includes(next)
    )
      return fail('invalid-input', 'A new block ID is already in use.');
    let nextTag: string;
    try {
      nextTag = blockTag(next);
    } catch {
      return fail('invalid-input', 'A new block ID is invalid.');
    }
    tags.set(tag, nextTag);
    payloadIds[id] = next;
    copies.push({ payload: { ...block.payload, id: next }, label: stored.label });
  }
  if (!native && !tags.size) return fail('invalid-input', 'The control is not a registered block.');
  const copiedIds = new Map(Object.entries(payloadIds));
  try {
    for (const copy of copies) copy.payload = copiedGalleryPayload(copy.payload, copiedIds);
    copies.sort((a, b) => galleryRecordOrder(a.payload, b.payload));
  } catch (error) {
    return fail('invalid-input', error instanceof Error ? error.message : String(error));
  }
  let templates: readonly OoxmlNode[];
  try {
    templates = freshenCopiedRunOrigins(
      range.nodes.map((node) => copyOf(node, tags)),
      500_000,
      64
    );
  } catch (error) {
    return fail(
      'unsupported-operation',
      error instanceof Error
        ? error.message
        : 'The copied text cannot retain independent review ownership.'
    );
  }
  const sourceLastId = range.nodes.at(-1)!.id;
  const rootTags = range.nodes.map(
    (node) => tags.get(contentControlPropertiesOf(node).tag ?? '') ?? null
  );
  const identitySeed = `scio-block-copy:${crypto.randomUUID()}`;
  let insertedIds: string[] = [];
  let placedOnce = false;
  const place = (working: OoxmlPackage): OoxmlPackage => {
    const live = working.parts.get(part.name);
    if (!live) throw new Error('The document part disappeared during the duplication.');
    const container = parentNodeOf(live, sourceLastId);
    if (!container) throw new Error('The duplicated block disappeared during the duplication.');
    const at = container.children.findIndex((child) => child.id === sourceLastId);
    const nextId = createNodeIdAllocator(live, 'new');
    const identity = {
      prefix: w14RootPrefix(live.root),
      used: new Set(usedParaIds(live.root)),
      seed: identitySeed,
    };
    const independent = freshenCopiedDrawingPropertyIds(working, templates, 500_000, 64);
    const nodes = independent.map((template) => withFreshIds(template, nextId, identity));
    insertedIds = nodes.map((node) => node.id);
    const placed = insertChildren(live, container.id, at + 1, nodes);
    if (!placed.ok) throw new Error('The copy was refused by the native store.');
    let next = withPart(working, placed.part);
    for (const copy of copies) {
      const stored = withPayloadRecord(next, copy.payload, copy.label);
      if (!stored.ok) throw new Error(`Block payload storage refused: ${stored.reason}`);
      next = stored.pkg;
    }
    placedOnce = true;
    return next;
  };
  // Local node IDs become shared IDs during the synchronous document-port notification.
  // Retain that exact mapping before the collaboration session retires its aliases.
  const stop = surface.session.subscribe(() => {
    const live = surface.session.part();
    const collaboration = surface.collaborationSession();
    insertedIds = insertedIds.map((id) => {
      if (findNode(live, id)) return id;
      const shared = collaboration?.resolveNodeId?.(id) ?? id;
      return findNode(live, shared) ? shared : id;
    });
  });
  try {
    place(pkg);
    placedOnce = false;
    const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [place], false, true);
    if (result.rejected || !result.committed || !placedOnce)
      return fail('apply-failed', String(result.reason ?? 'The duplication did not commit.'));
  } catch (error) {
    return fail('apply-failed', error instanceof Error ? error.message : 'The duplication failed.');
  } finally {
    stop();
  }
  const live = surface.session.part();
  insertedIds = insertedIds.map((id, position) => {
    if (findNode(live, id)) return id;
    const tag = rootTags[position];
    if (!tag) return id;
    const matching = contentControlsIn(live.root).filter(
      ({ node }) => contentControlPropertiesOf(node).tag === tag
    );
    return matching.length === 1 ? matching[0]!.node.id : id;
  });
  if (insertedIds.some((id) => !findNode(live, id)))
    return fail('apply-failed', 'A copied block cannot be found.');
  return {
    success: true,
    revision: surface.session.packageRevision(),
    nodeIds: insertedIds,
    payloadIds,
  };
}
