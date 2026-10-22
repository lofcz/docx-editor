import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { NativeEditFailure } from '../contracts/document';
import {
  DEFAULT_TARGET_PARSE_LIMITS,
  parseTargetMarkdown,
  targetSourceOffsetAt,
  type TargetAstNode,
} from '../projection/parse-target';
import { readEditorRevision } from '../projection/revision-cache';
import type { DirectPlanContext } from './apply-direct';
import { prepareGeneratedMarkdown } from './generated-markdown';
import { beginStreamInsertion, type StreamAppendResult } from './stream-insert';

export type MarkdownStreamResult =
  | StreamAppendResult
  | { readonly success: true; readonly committed: false };
export interface MarkdownInsertionStream {
  readonly state: 'active' | 'ended';
  /** Incomplete syntax stays request-local, available to the application's stream preview. */
  readonly pendingMarkdown: string;
  /** Append a decoded text delta. Completed blocks appear in the native document progressively. */
  push(delta: string): MarkdownStreamResult;
  /** Flush the final block before releasing the anchor, including a response without a final LF. */
  finish(): MarkdownStreamResult;
  /** Stop receiving data, flush valid received content, and release the anchor. */
  cancel(): MarkdownStreamResult;
}
const failure = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});
const buffered = (): MarkdownStreamResult => ({ success: true, committed: false });

// A future definition can reinterpret an unresolved bracket. Keep that syntax until EOF.
const needsFullContext = (node: TargetAstNode): boolean =>
  node.type === 'definition' ||
  node.type === 'linkReference' ||
  node.type === 'imageReference' ||
  (node.type === 'text' && typeof node.value === 'string' && node.value.includes('[')) ||
  (Array.isArray(node.children) && (node.children as TargetAstNode[]).some(needsFullContext));

/** Stream parser and native insertion share no serialized placeholder or hidden document node. */
export function beginMarkdownInsertion(
  editor: DocxEditorInstance,
  context: DirectPlanContext,
  options: {
    readonly selectedIds?: readonly string[];
    readonly allocateBlockId?: () => string;
  } = {}
): { readonly success: true; readonly stream: MarkdownInsertionStream } | NativeEditFailure {
  const started = beginStreamInsertion(editor, context, options.selectedIds);
  if (!started.success) return started;
  const insertion = started.insertion;
  let pending = '';
  let bytes = 0;
  let last = 0;
  let final: MarkdownStreamResult | undefined;
  function flush(finish: boolean): MarkdownStreamResult {
    if (!pending.trim()) {
      if (finish) pending = '';
      return buffered();
    }
    const parsed = parseTargetMarkdown(pending);
    // An unfinished MDX tag can temporarily make the entire pending parse invalid.
    if (!parsed.success) return finish ? parsed : buffered();
    let cut = pending.length;
    if (!finish) {
      const children = parsed.target.ast.children as TargetAstNode[];
      if (children.length < 2 || children.some(needsFullContext)) return buffered();
      const position = children.at(-1)!.position as unknown as { start: { offset: number } };
      const boundary = targetSourceOffsetAt(parsed.target, position.start.offset);
      if (boundary === null || boundary <= 0) return buffered();
      cut = boundary;
    }
    const revision = readEditorRevision(editor, context.documentId, context.epoch);
    const surface = editor.surface;
    if (!revision || !surface)
      return failure('document-not-ready', 'The native document is unavailable.');
    const prepared = prepareGeneratedMarkdown(
      { pkg: surface.session.currentPackage(), revision, partName: surface.session.part().name },
      pending.slice(0, cut),
      options.allocateBlockId
    );
    if (!prepared.success) return prepared;
    const result = insertion.append(prepared.fragment);
    if (result.success) pending = pending.slice(cut);
    return result;
  }
  const stream: MarkdownInsertionStream = {
    get state() {
      return insertion.state;
    },
    get pendingMarkdown() {
      return pending;
    },
    push(delta) {
      if (insertion.state !== 'active')
        return failure('stale-revision', 'The generation request ended.');
      if (typeof delta !== 'string')
        return failure('invalid-input', 'A stream delta must be text.');
      const first = delta.charCodeAt(0);
      const joinedPair = last >= 0xd800 && last <= 0xdbff && first >= 0xdc00 && first <= 0xdfff;
      const nextBytes = bytes + new TextEncoder().encode(delta).length - (joinedPair ? 2 : 0);
      if (nextBytes > DEFAULT_TARGET_PARSE_LIMITS.maxBytes)
        return failure('invalid-input', 'Generated Markdown exceeds its byte limit.');
      bytes = nextBytes;
      if (delta.length) last = delta.charCodeAt(delta.length - 1);
      pending += delta;
      return delta.includes('\n') ? flush(false) : buffered();
    },
    finish() {
      if (final) return final;
      if (insertion.state !== 'active')
        return failure('stale-revision', 'The generation request ended.');
      try {
        return (final = flush(true));
      } finally {
        insertion.finish();
      }
    },
    cancel: () => stream.finish(),
  };
  return { success: true, stream };
}
