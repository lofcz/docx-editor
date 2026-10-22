import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { contentControlPropertiesOf, findNode, type TreeDocOp } from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import type { BlockPayload } from './payload-schema';
import { withPayloadRecord } from './payload-opc-codec';

export interface BlockPayloadWrite {
  readonly controlId: string;
  readonly payload: BlockPayload;
  readonly label: string;
  /** Native operations that update the readable Word fallback. */
  readonly fallbackOps: readonly TreeDocOp[];
  readonly expectedRevision?: number;
}

export type BlockPayloadWriteResult =
  | { readonly success: true; readonly revision: number; readonly tag: string }
  | NativeEditFailure;

/** Update an existing block's metadata and native fallback in one gated transaction. */
export function writeBlockPayload(
  editor: DocxEditorInstance,
  write: BlockPayloadWrite
): BlockPayloadWriteResult {
  const surface = editor.surface;
  if (!surface)
    return { success: false, code: 'document-not-ready', error: 'The document is not attached.' };
  surface.flushPendingInput();
  const revision = surface.session.packageRevision();
  if (write.expectedRevision !== undefined && write.expectedRevision !== revision)
    return {
      success: false,
      code: 'stale-revision',
      error: 'Read the current document before updating this block.',
    };
  // Property review needs coordinated old/new payloads, implemented by the review layer.
  if (surface.editingMode() === 'suggest')
    return {
      success: false,
      code: 'unsupported-operation',
      error: 'Tracked block properties require a review transaction.',
    };
  const prepared = withPayloadRecord(surface.session.currentPackage(), write.payload, write.label);
  if (!prepared.ok) return { success: false, code: 'invalid-document', error: prepared.reason };
  const control = findNode(surface.session.part(), write.controlId);
  if (!control || control.kind !== 'contentControl')
    return { success: false, code: 'invalid-document', error: 'The block control is missing.' };
  // Unchanged properties stay out of the transaction, so concurrent writers share no property
  // write. A missing alias stays missing: two sessions adding one at once would each insert a
  // `w:alias`, and the merged control would hold two. Created blocks carry their alias.
  const properties = contentControlPropertiesOf(control);
  const tag = properties.tag === prepared.tag ? undefined : prepared.tag;
  const alias =
    properties.alias === undefined || properties.alias === write.label ? undefined : write.label;
  const ops: TreeDocOp[] = [
    ...(tag === undefined && alias === undefined
      ? []
      : [
          {
            op: 'setContentControlProperties' as const,
            controlId: write.controlId,
            ...(tag === undefined ? {} : { tag }),
            ...(alias === undefined ? {} : { alias }),
          },
        ]),
    ...write.fallbackOps,
  ];
  try {
    // Each submitted widget action is one undo unit, including rapid consecutive actions.
    const result = surface.applyAutomationOps(
      () => ops,
      { kind: 'body' },
      [
        (working) => {
          const next = withPayloadRecord(working, write.payload, write.label);
          if (!next.ok) throw new Error(`Payload transaction refused: ${next.reason}`);
          return next.pkg;
        },
      ],
      false,
      true
    );
    if (result.rejected || !result.committed)
      return {
        success: false,
        code: 'apply-failed',
        error: String(result.reason ?? 'The block update did not commit.'),
      };
    return { success: true, revision: surface.session.packageRevision(), tag: prepared.tag };
  } catch (error) {
    return {
      success: false,
      code: 'apply-failed',
      error: error instanceof Error ? error.message : 'The block update failed.',
    };
  }
}
