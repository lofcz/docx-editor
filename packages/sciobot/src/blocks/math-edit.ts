import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { SemanticSelection } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  contentControlPropertiesOf,
  findNode,
  isAtomicContentControl,
  withPart,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { buildBlockIndex } from './block-index';
import { ancestorContentWriteRefusal } from './ancestor-content-write';
import { convertMathControl } from './math-convert';
import { insertBlockAtSelection } from './insert';
import { withAtomicMathFallbacks } from './math-ownership';
import { updateAtomicBlock } from './update-atomic';
import {
  BLOCK_MATH,
  INLINE_MATH,
  createMathPayload,
  mathWordFallback,
  withMathSource,
  type MathKind,
  type MathPayload,
} from './math-codec';
import { blockTag, readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import { encodeBlockPayload, type BlockPayload } from './payload-schema';

export type MathEditResult =
  { readonly success: true; readonly revision: number } | NativeEditFailure;
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});
function gate(editor: DocxEditorInstance): NativeEditFailure | null {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (surface.editingMode() !== 'edit')
    return fail(
      'unsupported-operation',
      surface.editingMode() === 'suggest'
        ? 'Tracked equation changes require a review transaction.'
        : 'The document is read-only.'
    );
  return null;
}

/** Publication can refuse and restore the shared package during the local transaction. */
function published(editor: DocxEditorInstance, payload: BlockPayload): MathEditResult {
  const surface = editor.surface!;
  const live = buildBlockIndex(surface.session.currentPackage()).byId.get(payload.id)?.payload;
  if (!live || encodeBlockPayload(live) !== encodeBlockPayload(payload))
    return fail(
      'apply-failed',
      'The equation change did not reach the shared document. Your draft remains open.'
    );
  return { success: true, revision: surface.session.packageRevision() };
}

/**
 * Commit one equation draft. Its payload is the optimistic ownership token: unrelated peer
 * edits can proceed, but an edited or deleted equation cannot be overwritten by an old draft.
 * Display conversion preserves the payload ID and canonical control identity.
 */
export function updateMath(
  editor: DocxEditorInstance,
  request: {
    readonly original: BlockPayload;
    readonly source: string;
    readonly kind: MathKind;
  }
): MathEditResult {
  const refused = gate(editor);
  if (refused) return refused;
  const surface = editor.surface!;
  const live = buildBlockIndex(surface.session.currentPackage()).byId.get(request.original.id);
  if (!live) return fail('stale-revision', 'The equation was deleted. Your draft remains open.');
  if (encodeBlockPayload(live.payload) !== encodeBlockPayload(request.original))
    return fail('stale-revision', 'The equation changed. Your draft remains open.');
  const control = findNode(
    surface.session.currentPackage().parts.get(live.partName)!,
    live.controlId
  );
  if (control) {
    const protectedAncestor = ancestorContentWriteRefusal(
      surface.session.currentPackage().parts.get(live.partName)!,
      control.id
    );
    if (protectedAncestor || contentControlPropertiesOf(control).dataBinding)
      return fail(
        'unsupported-operation',
        protectedAncestor ?? 'The equation control is bound to custom XML.'
      );
  }
  if (
    editor.snapshot().collaborationStatus !== 'inactive' &&
    (!control || !isAtomicContentControl(control))
  )
    return fail(
      'unsupported-operation',
      'Prepare equation ownership before creating the shared document. Your draft remains open.'
    );
  let payload: MathPayload;
  try {
    payload = withMathSource(live.payload, request.source);
  } catch (error) {
    return fail('invalid-input', String(error));
  }
  if (payload.type === (request.kind === 'inline' ? INLINE_MATH : BLOCK_MATH)) {
    const updated = updateAtomicBlock(editor, {
      blockId: payload.id,
      attributes: payload.attributes,
    });
    return updated.success ? published(editor, payload) : updated;
  }
  payload = {
    ...payload,
    type: request.kind === 'inline' ? INLINE_MATH : BLOCK_MATH,
  };
  const replacement = payload;
  const convert = (pkg: OoxmlPackage) => {
    const converted = convertMathControl(pkg, live.partName, replacement);
    const saved = withPayloadRecord(converted, replacement, mathWordFallback(replacement).label);
    if (!saved.ok) throw Error(saved.reason);
    return saved.pkg;
  };
  try {
    convert(surface.session.currentPackage());
    const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [convert], false, true);
    if (result.rejected || !result.committed) return fail('apply-failed', String(result.reason));
    return published(editor, payload);
  } catch (error) {
    return fail('unsupported-operation', error instanceof Error ? error.message : String(error));
  }
}

/** Replace a captured text range and insert its equation in one native history transaction. */
export function insertMathAtSelection(
  editor: DocxEditorInstance,
  request: {
    readonly id: string;
    readonly source: string;
    readonly kind: MathKind;
    readonly selection: SemanticSelection;
    readonly expectedRevision: number;
  }
): MathEditResult {
  const refused = gate(editor);
  if (refused) return refused;
  const surface = editor.surface!;
  if (surface.session.packageRevision() !== request.expectedRevision)
    return fail('stale-revision', 'The insertion position changed. Your draft remains open.');
  const pkg = surface.session.currentPackage();
  const store = readPayloadStore(pkg);
  if (store.status === 'invalid') return fail('invalid-document', store.reason);
  if (editor.snapshot().collaborationStatus !== 'inactive' && store.status !== 'ready')
    return fail(
      'unsupported-operation',
      'Prepare native payload storage before sharing application equations.'
    );
  if (
    buildBlockIndex(pkg).byId.has(request.id) ||
    (store.status === 'ready' && store.entries.has(request.id))
  )
    return fail('invalid-input', 'An equation with this ID already exists.');
  let payload;
  try {
    payload = createMathPayload(request.id, request.kind, request.source);
  } catch (error) {
    return fail('invalid-input', String(error));
  }
  const fallback = mathWordFallback(payload);
  if (request.kind === 'block') {
    const inserted = insertBlockAtSelection(editor, {
      payload,
      label: fallback.label,
      paragraphs: fallback.paragraphs,
      atomicContent: true,
      selection: request.selection,
      expectedRevision: request.expectedRevision,
    });
    return inserted.success ? published(editor, payload) : inserted;
  }
  const { anchor, head } = request.selection;
  const plan = surface.planRangeDeletion(request.selection);
  const expanded = anchor.paragraphId !== head.paragraphId || anchor.offset !== head.offset;
  if (expanded && !plan.ops.length)
    return fail('unsupported-operation', 'This text selection cannot be replaced.');
  const paragraphId = plan.replaceAt.paragraphId;
  const start = plan.replaceAt.offset;
  const ops: TreeDocOp[] = [...plan.ops];
  ops.push({
    op: 'insertInlineContentControl',
    paragraphId,
    offset: start,
    tag: blockTag(payload.id),
    text: fallback.inlineText,
    alias: fallback.label,
    lock: 'contentLocked',
  });
  let part = surface.session.part();
  for (const op of ops) {
    const applied = applyTreeOp(part, op);
    if (!applied.ok) return fail('unsupported-operation', applied.reason);
    part = applied.part;
  }
  const write = (working: OoxmlPackage) => {
    const stored = withPayloadRecord(working, payload, fallback.label);
    if (!stored.ok) throw Error(stored.reason);
    return withAtomicMathFallbacks(stored.pkg, new Set([payload.id]));
  };
  try {
    write(withPart(pkg, part));
    const result = surface.applyAutomationOps(
      () => ops,
      { kind: 'body' },
      [write],
      false,
      true,
      'after'
    );
    if (result.rejected || !result.committed) return fail('apply-failed', String(result.reason));
    const point = { paragraphId, offset: start + fallback.inlineText.length };
    surface.setSelection({ anchor: point, head: point });
    return published(editor, payload);
  } catch (error) {
    return fail('unsupported-operation', error instanceof Error ? error.message : String(error));
  }
}
