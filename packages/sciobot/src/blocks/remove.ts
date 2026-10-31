import { activeMdxRegistry } from '../registry/active';
import { isInlineComponentPayload } from '../registry/mdx-read';
import { parentNodeOf } from '@docx-editor.dev/core/store';
import { encodeBlockPayload } from './payload-schema';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  bodyStoryRoot,
  contentControlPropertiesOf,
  contentControlsIn,
  findNode,
  isStoryPart,
  replaceNode,
  storyParagraphs,
  withoutCustomXmlNode,
  withPart,
  type OoxmlNode,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { buildBlockIndex, type IndexedBlock } from './block-index';
import { blockIdFromTag, readPayloadStore } from './payload-opc-codec';

export interface BlockRemoval {
  readonly controlId: string;
  readonly mode: 'delete' | 'unwrap';
  readonly expectedRevision?: number;
}
export type BlockRemoveResult =
  | {
      readonly success: true;
      readonly revision: number;
      readonly removedPayloadIds: readonly string[];
    }
  | NativeEditFailure;
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

/** Several registered wrappers removed together, as one block selection. */
export interface BlocksRemoval {
  readonly controlIds: readonly string[];
  readonly mode: 'delete' | 'unwrap';
  readonly expectedRevision?: number;
}

/** Remove a registered wrapper and its own metadata in one native undo unit. */
export function removeBlock(editor: DocxEditorInstance, request: BlockRemoval): BlockRemoveResult {
  return removeBlocks(editor, {
    controlIds: [request.controlId],
    mode: request.mode,
    ...(request.expectedRevision === undefined
      ? {}
      : { expectedRevision: request.expectedRevision }),
  });
}

/**
 * Remove registered wrappers and their own metadata in one native undo unit. A wrapper inside
 * another selected wrapper is refused: removing the outer one already removes it.
 */
export function removeBlocks(
  editor: DocxEditorInstance,
  request: BlocksRemoval
): BlockRemoveResult {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (
    request.expectedRevision !== undefined &&
    request.expectedRevision !== surface.session.packageRevision()
  )
    return fail('stale-revision', 'Read the current document before removing this block.');
  if (request.mode !== 'delete' && request.mode !== 'unwrap')
    return fail('invalid-input', 'Choose delete or unwrap.');
  if (!request.controlIds.length || new Set(request.controlIds).size !== request.controlIds.length)
    return fail('invalid-input', 'Choose distinct blocks to remove.');
  if (surface.editingMode() === 'suggest')
    return fail('unsupported-operation', 'Tracked block removal requires a review transaction.');
  const pkg = surface.session.currentPackage(),
    part = surface.session.part();
  const index = buildBlockIndex(pkg);
  const targets: { block: IndexedBlock; control: OoxmlNode & { kind: 'contentControl' } }[] = [];
  for (const controlId of request.controlIds) {
    const block = [...index.byId.values()].find((entry) => entry.controlId === controlId);
    const control = findNode(part, controlId);
    if (!block || block.partName !== part.name || !control || control.kind !== 'contentControl')
      return fail('invalid-input', 'The registered block is missing from the active story.');
    targets.push({ block, control });
  }
  for (const { control } of targets)
    for (
      let parent = parentNodeOf(part, control.id);
      parent;
      parent = parentNodeOf(part, parent.id)
    )
      if (request.controlIds.includes(parent.id))
        return fail('invalid-input', 'A selected block is inside another selected block.');
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready')
    return fail('invalid-document', 'The block payload store is unavailable.');
  const ids = new Set<string>();
  // Payload IDs removed with a host-owned inline atom; their records retire in collaboration.
  const inlineIds = new Set<string>();
  const ownsInline = new Map<string, boolean>();
  for (const { block, control } of targets) {
    encodeBlockPayload(block.payload);
    // The host owns these immutable inline atoms. Ordinary text editing retains their lock.
    const ownedInline =
      parentNodeOf(part, control.id)?.kind === 'paragraph' &&
      // Inline controls: images are drawings, so only components with an inline binding count.
      !!activeMdxRegistry().components.get(block.payload.type)?.inline &&
      isInlineComponentPayload(block.payload, activeMdxRegistry());
    ownsInline.set(control.id, ownedInline);
    const owned = [block.id];
    if (request.mode === 'delete')
      for (const { node } of contentControlsIn(control)) {
        const id = blockIdFromTag(contentControlPropertiesOf(node).tag ?? '');
        if (id !== null) owned.push(id);
      }
    for (const id of owned) {
      ids.add(id);
      if (ownedInline) inlineIds.add(id);
    }
  }
  // Never delete a payload used by a second control, including one in another story.
  const counts = new Map<string, number>();
  for (const story of pkg.parts.values()) {
    if (!isStoryPart(story)) continue;
    for (const { node } of contentControlsIn(story.root)) {
      const id = blockIdFromTag(contentControlPropertiesOf(node).tag ?? '');
      if (id !== null && ids.has(id)) counts.set(id, (counts.get(id) ?? 0) + 1);
    }
  }
  if ([...ids].some((id) => (counts.get(id) ?? 0) !== 1))
    return fail('ambiguous-target', 'A removed block ID does not identify one native control.');
  const removedPayloadIds = [...ids].filter((id) => store.entries.has(id));
  const cleanup = (working: OoxmlPackage) => {
    let next = working;
    for (const id of removedPayloadIds) {
      if (!inlineIds.has(id) || editor.snapshot().collaborationStatus === 'inactive') {
        next = withoutCustomXmlNode(next, store.part.partName, `scio_${id}`);
        continue;
      }
      // Adopted custom XML order can differ from storage order. Retire by identity,
      // keeping the same empty history shell that presence decisions already use.
      const part = next.parts.get(store.part.partName)!;
      const node = part.root.children.find(
        (child) =>
          child.kind !== 'textValue' &&
          child.attributes.some(
            (attribute) =>
              attribute.namespaceUri === '' &&
              attribute.localName === 'id' &&
              attribute.value === `scio_${id}`
          )
      );
      if (!node || node.kind === 'textValue') throw Error('The removed payload record is missing.');
      const retired = replaceNode(part, node.id, { ...node, attributes: [], children: [] });
      if (!retired.ok) throw Error('The removed payload record could not retire.');
      next = withPart(next, retired.part);
    }
    const remaining = readPayloadStore(next);
    if (remaining.status !== 'ready' || removedPayloadIds.some((id) => remaining.entries.has(id)))
      throw new Error('The removed block payloads remain in the native store.');
    return next;
  };
  const ops: TreeDocOp[] = [];
  let preview = part;
  for (const { control } of targets) {
    // Delete only registered native atom controls; explicit user locks on other blocks still
    // refuse.
    if (
      request.mode === 'delete' &&
      ownsInline.get(control.id) &&
      contentControlPropertiesOf(control).lock === 'sdtContentLocked'
    ) {
      const unlock: TreeDocOp = {
        op: 'setContentControlProperties',
        controlId: control.id,
        lock: 'unlocked',
      };
      const result = applyTreeOp(preview, unlock);
      if (!result.ok) return fail('apply-failed', `Native atom unlock refused: ${result.reason}`);
      preview = result.part;
      ops.push(unlock);
    }
    const op: TreeDocOp = {
      op: 'removeContentControl',
      controlId: control.id,
      keepContent: request.mode === 'unwrap',
    };
    const removed = applyTreeOp(preview, op);
    if (!removed.ok) return fail('apply-failed', `Native block removal refused: ${removed.reason}`);
    preview = removed.part;
    ops.push(op);
  }
  let prepared: { ok: true; part: typeof part } | { ok: false; reason: string } = {
    ok: true,
    part: preview,
  };
  const body = bodyStoryRoot(prepared.part);
  const needsParagraph = body !== null && storyParagraphs(body).length === 0;
  if (needsParagraph) {
    const original = findNode(part, surface.session.paragraphIds()[0]!);
    if (!original || original.kind !== 'paragraph')
      return fail('invalid-document', 'The removed block has no native paragraph template.');
    const insertion: TreeDocOp = {
      op: 'insertParagraphs',
      parentId: body!.id,
      beforeId: null,
      paragraphs: [
        {
          ...original,
          children: original.children.filter((child) => child.kind === 'paragraphProperties'),
        },
      ],
    };
    prepared = applyTreeOp(prepared.part, insertion);
    if (!prepared.ok)
      return fail('apply-failed', `Native empty paragraph refused: ${prepared.reason}`);
    ops.push(insertion);
  }
  try {
    cleanup(withPart(pkg, prepared.part));
    const result = surface.applyAutomationOps(() => ops, { kind: 'body' }, [cleanup], false, true);
    if (result.rejected || !result.committed)
      return fail('apply-failed', String(result.reason ?? 'The block removal did not commit.'));
    if (needsParagraph) {
      const paragraphId = surface.session.paragraphIds()[0]!;
      const position = { paragraphId, offset: 0 };
      surface.setSelection({ anchor: position, head: position });
    }
    return { success: true, revision: surface.session.packageRevision(), removedPayloadIds };
  } catch (error) {
    return fail(
      'apply-failed',
      error instanceof Error ? error.message : 'The block removal failed.'
    );
  }
}
