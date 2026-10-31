import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  contentControlPropertiesOf,
  isStoryPart,
  type OoxmlNode,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { activeMdxRegistry } from '../registry/active';
import { isInlineComponentPayload } from '../registry/mdx-read';
import { blockTag, readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import { encodeBlockPayload, type BlockPayload } from './payload-schema';

export interface InlineComponentInsertion {
  readonly payload: BlockPayload;
  readonly paragraphId: string;
  readonly offset: number;
  readonly expectedRevision?: number;
}
export type InlineComponentInsertResult =
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

/**
 * Insert one registered inline component at a text position, with its payload record and the
 * registry's readable Word fallback, in one native history unit. Tracked insertion belongs to
 * the review executor.
 */
export function insertInlineComponent(
  editor: DocxEditorInstance,
  request: InlineComponentInsertion
): InlineComponentInsertResult {
  const registry = activeMdxRegistry();
  const binding = registry.components.get(request.payload.type)?.inline;
  let text: string;
  let tag: string;
  try {
    encodeBlockPayload(request.payload);
    if (!binding || !isInlineComponentPayload(request.payload, registry))
      return fail('invalid-input', 'The payload is not a valid inline component.');
    text = binding.fallback(request.payload).text;
    tag = blockTag(request.payload.id);
  } catch (error) {
    return fail(
      'invalid-input',
      error instanceof Error ? error.message : 'Invalid inline component payload.'
    );
  }
  const label = request.payload.type;
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (
    request.expectedRevision !== undefined &&
    request.expectedRevision !== surface.session.packageRevision()
  )
    return fail('stale-revision', 'Read the current document before inserting this component.');
  if (surface.editingMode() !== 'edit')
    return fail('unsupported-operation', 'Inline component insertion requires edit mode.');
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
      'Prepare native payload storage before sharing application components.'
    );
  const op: TreeDocOp = {
    op: 'insertInlineContentControl',
    paragraphId: request.paragraphId,
    offset: request.offset,
    tag,
    text,
    alias: label,
    lock: 'sdtContentLocked',
  };
  const preview = applyTreeOp(surface.session.part(), op);
  if (!preview.ok) return fail('invalid-input', `Native insertion refused: ${preview.reason}.`);
  const prepared = withPayloadRecord(pkg, request.payload, label);
  if (!prepared.ok) return fail('invalid-document', prepared.reason);
  try {
    const result = surface.applyAutomationOps(
      () => [op],
      { kind: 'body' },
      [
        (working) => {
          const next = withPayloadRecord(working, request.payload, label);
          if (!next.ok) throw new Error(`Component storage refused: ${next.reason}`);
          return next.pkg;
        },
      ],
      false,
      true
    );
    if (result.rejected || !result.committed)
      return fail('apply-failed', String(result.reason ?? 'The insertion did not commit.'));
    const inserted = control(surface.session.part().root, tag);
    if (!inserted) return fail('apply-failed', 'The inserted component cannot be found.');
    return { success: true, revision: surface.session.packageRevision(), controlId: inserted.id };
  } catch (error) {
    return fail('apply-failed', error instanceof Error ? error.message : 'The insertion failed.');
  }
}
