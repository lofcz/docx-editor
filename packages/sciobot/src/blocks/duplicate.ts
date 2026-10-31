import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  contentControlPropertiesOf,
  contentControlsIn,
  createNodeIdAllocator,
  findNode,
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
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
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
  const control = findNode(part, request.controlId);
  const parent = control && parentNodeOf(part, control.id);
  if (!control || control.kind !== 'contentControl' || !parent || parent.kind === 'paragraph')
    return fail('invalid-input', 'The registered block is missing from the active story.');
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready')
    return fail('invalid-document', 'The block payload store is unavailable.');
  const index = buildBlockIndex(pkg);
  const tags = new Map<string, string>();
  const payloadIds: Record<string, string> = {};
  const copies: { payload: BlockPayload; label: string }[] = [];
  for (const { node } of [{ node: control }, ...contentControlsIn(control)]) {
    const tag = contentControlPropertiesOf(node).tag ?? '';
    const id = blockIdFromTag(tag);
    if (id === null || Object.hasOwn(payloadIds, id)) continue;
    const block = index.byId.get(id);
    const stored = store.entries.get(id);
    if (!block || !stored)
      return fail('invalid-document', 'A copied block has no stored payload record.');
    const next = request.allocateBlockId(block.payload.type);
    if (index.byId.has(next) || store.entries.has(next) || Object.values(payloadIds).includes(next))
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
  if (!tags.size) return fail('invalid-input', 'The control is not a registered block.');
  const template = copyOf(control, tags);
  const rootTag = tags.get(contentControlPropertiesOf(control).tag ?? '')!;
  let placedOnce = false;
  const place = (working: OoxmlPackage): OoxmlPackage => {
    const live = working.parts.get(part.name);
    if (!live) throw new Error('The document part disappeared during the duplication.');
    const container = parentNodeOf(live, control.id);
    if (!container) throw new Error('The duplicated block disappeared during the duplication.');
    const at = container.children.findIndex((child) => child.id === control.id);
    const node = withFreshIds(template, createNodeIdAllocator(live, 'new'), {
      prefix: w14RootPrefix(live.root),
      used: new Set(usedParaIds(live.root)),
      seed: `scio-block:${rootTag}`,
    });
    const placed = insertChildren(live, container.id, at + 1, [node]);
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
  try {
    place(pkg);
    placedOnce = false;
    const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [place], false, true);
    if (result.rejected || !result.committed || !placedOnce)
      return fail('apply-failed', String(result.reason ?? 'The duplication did not commit.'));
  } catch (error) {
    return fail('apply-failed', error instanceof Error ? error.message : 'The duplication failed.');
  }
  // Collaboration can rename the new nodes during the commit, so read the copy back by tag.
  const live = [...contentControlsIn(surface.session.part().root)].find(
    ({ node }) => contentControlPropertiesOf(node).tag === rootTag
  );
  if (!live) return fail('apply-failed', 'The copied block cannot be found.');
  return {
    success: true,
    revision: surface.session.packageRevision(),
    controlId: live.node.id,
    payloadIds,
  };
}
