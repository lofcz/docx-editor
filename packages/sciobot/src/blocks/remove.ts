import { parentNodeOf } from '@docx-editor.dev/core/store';
import { decodeMathPayload } from './math-codec';
import { decodeMentionPayload } from './mention-codec';
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
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { buildBlockIndex } from './block-index';
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

/** Remove a registered wrapper and its own metadata in one native undo unit. */
export function removeBlock(editor: DocxEditorInstance, request: BlockRemoval): BlockRemoveResult {
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
  if (surface.editingMode() === 'suggest')
    return fail('unsupported-operation', 'Tracked block removal requires a review transaction.');
  const pkg = surface.session.currentPackage(),
    part = surface.session.part();
  const index = buildBlockIndex(pkg);
  const block = [...index.byId.values()].find((entry) => entry.controlId === request.controlId);
  const control = findNode(part, request.controlId);
  if (!block || block.partName !== part.name || !control || control.kind !== 'contentControl')
    return fail('invalid-input', 'The registered block is missing from the active story.');
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready')
    return fail('invalid-document', 'The block payload store is unavailable.');
  const ids = new Set([block.id]);
  if (request.mode === 'delete')
    for (const { node } of contentControlsIn(control)) {
      const id = blockIdFromTag(contentControlPropertiesOf(node).tag ?? '');
      if (id !== null) ids.add(id);
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
      if (!ownedInline || editor.snapshot().collaborationStatus === 'inactive') {
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
  const op = {
    op: 'removeContentControl' as const,
    controlId: control.id,
    keepContent: request.mode === 'unwrap',
  };
  const ops: TreeDocOp[] = [];
  let preview = part;
  // The host owns these immutable inline atoms. Ordinary text editing retains their lock.
  // Delete only registered native atom controls; explicit user locks on other blocks still refuse.
  const encoded = encodeBlockPayload(block.payload);
  const math = decodeMathPayload(encoded);
  const mention = decodeMentionPayload(encoded);
  const ownedInline =
    parentNodeOf(part, control.id)?.kind === 'paragraph' &&
    ((math.status === 'supported' && math.payload.type === 'inline_equation') ||
      mention.status === 'supported' ||
      block.payload.type === 'source');
  if (
    request.mode === 'delete' &&
    ownedInline &&
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
  let prepared = applyTreeOp(preview, op);
  if (!prepared.ok) return fail('apply-failed', `Native block removal refused: ${prepared.reason}`);
  ops.push(op);
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
