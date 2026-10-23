import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  contentControlPropertiesOf,
  isStoryPart,
  type OoxmlNode,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { decodeMentionPayload, mentionWordFallback, type MentionPayload } from './mention-codec';
import { blockTag, readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import { encodeBlockPayload } from './payload-schema';

export interface MentionInsertion {
  readonly payload: MentionPayload;
  readonly paragraphId: string;
  readonly offset: number;
  readonly expectedRevision?: number;
}
export type MentionInsertResult =
  | { readonly success: true; readonly revision: number; readonly controlId: string }
  | NativeEditFailure;
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});
function control(root: OoxmlNode, tag: string): OoxmlNode | null {
  if (root.kind === 'textValue') return null;
  if (root.kind === 'contentControl' && contentControlPropertiesOf(root).tag === tag) return root;
  for (const child of root.children) {
    const found = control(child, tag);
    if (found) return found;
  }
  return null;
}

/** Persist original mention identity with a literal Word fallback in one native history unit.
 * Tracked insertion belongs to the review executor.
 */
export function insertMention(
  editor: DocxEditorInstance,
  request: MentionInsertion
): MentionInsertResult {
  let fallback: ReturnType<typeof mentionWordFallback>;
  let tag: string;
  try {
    const read = decodeMentionPayload(encodeBlockPayload(request.payload));
    if (read.status !== 'supported')
      return fail('invalid-input', 'The mention payload is invalid.');
    fallback = mentionWordFallback(read.payload);
    tag = blockTag(read.payload.id);
  } catch (error) {
    return fail(
      'invalid-input',
      error instanceof Error ? error.message : 'Invalid mention payload.'
    );
  }
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (
    request.expectedRevision !== undefined &&
    request.expectedRevision !== surface.session.packageRevision()
  )
    return fail('stale-revision', 'Read the current document before inserting this mention.');
  if (surface.editingMode() === 'suggest')
    return fail(
      'unsupported-operation',
      'Tracked mention insertion requires a review transaction.'
    );
  const pkg = surface.session.currentPackage();
  const store = readPayloadStore(pkg);
  if (store.status === 'invalid') return fail('invalid-document', store.reason);
  if (
    (store.status === 'ready' && store.entries.has(request.payload.id)) ||
    [...pkg.parts.values()].some((part) => isStoryPart(part) && control(part.root, tag))
  )
    return fail('invalid-input', 'An application block with this ID already exists.');
  if (editor.snapshot().collaborationStatus !== 'inactive' && store.status !== 'ready')
    return fail(
      'unsupported-operation',
      'Prepare native payload storage before sharing application mentions.'
    );
  const op: TreeDocOp = {
    op: 'insertInlineContentControl',
    paragraphId: request.paragraphId,
    offset: request.offset,
    tag,
    text: fallback.text,
    alias: fallback.label,
    lock: 'sdtContentLocked',
  };
  const preview = applyTreeOp(surface.session.part(), op);
  if (!preview.ok)
    return fail('invalid-input', `Native mention insertion refused: ${preview.reason}.`);
  const prepared = withPayloadRecord(pkg, request.payload, fallback.label);
  if (!prepared.ok) return fail('invalid-document', prepared.reason);
  try {
    const result = surface.applyAutomationOps(
      () => [op],
      { kind: 'body' },
      [
        (working) => {
          const next = withPayloadRecord(working, request.payload, fallback.label);
          if (!next.ok) throw new Error(`Mention storage refused: ${next.reason}`);
          return next.pkg;
        },
      ],
      false,
      true
    );
    if (result.rejected || !result.committed)
      return fail('apply-failed', String(result.reason ?? 'The mention insertion did not commit.'));
    const inserted = control(surface.session.part().root, tag);
    if (!inserted) return fail('apply-failed', 'The inserted mention cannot be found.');
    return { success: true, revision: surface.session.packageRevision(), controlId: inserted.id };
  } catch (error) {
    return fail(
      'apply-failed',
      error instanceof Error ? error.message : 'The mention insertion failed.'
    );
  }
}
