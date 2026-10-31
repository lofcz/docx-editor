import { adapterTagPrefix } from '../registry/adapter-names';
import { beginHistoryGroupFrames, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { HistoryGroupOutcome } from '@docx-editor.dev/core/contracts/editor';
import {
  contentControlPropertiesOf,
  findNode,
  paraIdOf,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { createEditPlan, type CanonicalEdit } from '../diff/plan';
import { readEditorRevision } from '../projection/revision-cache';
import { applyDirectPlan, type DirectPlanContext, type DirectPlanResult } from './apply-direct';
import { createStreamAnchor } from './stream-anchor';

/** A completed generated fragment, with the existing D05 rich-content bindings. */
export type GeneratedFragment = Pick<
  Extract<CanonicalEdit, { kind: 'siblings' }>,
  'inserted' | 'lists' | 'blocks' | 'styles'
>;

export type StreamAppendResult =
  | (Extract<DirectPlanResult, { success: true }> & {
      readonly history: HistoryGroupOutcome | null;
    })
  | NativeEditFailure;

export interface StreamInsertion {
  readonly state: 'active' | 'ended';
  /**
   * Append a completed fragment. Incomplete network tokens remain with the stream parser.
   * An intervening local gesture can split native history; the result reports that split
   * and the request ends. Peers' edits do not split this editor's gesture.
   */
  append(fragment: GeneratedFragment): StreamAppendResult;
  /** End generation and release the ephemeral anchor. Committed content stays. */
  finish(): void;
  /** Stop generation and release the anchor. Partial committed content stays undoable. */
  cancel(): void;
}

const stale = (error: string): NativeEditFailure => ({
  success: false,
  code: 'stale-revision',
  error,
});
// Content edits do not change ownership. Never use text equality to identify generated roots.
const rootKeys = (nodes: readonly OoxmlNode[]) =>
  nodes.map((node) => {
    if (node.kind === 'paragraph') return `paragraph:${paraIdOf(node) ?? node.id}`;
    if (node.kind === 'contentControl') {
      const tag = contentControlPropertiesOf(node).tag;
      if (tag?.startsWith(adapterTagPrefix())) return `control:${tag}`;
    }
    return `${node.kind}:${node.id}`;
  });

/**
 * Append completed generated fragments after one captured selection scope. Each frame uses the
 * native direct executor and explicit history group. A reported native split ends the request.
 * Without a split all frames share one undo gesture. The anchor never enters
 * serialized content. Concurrent edits outside the insertion range remain canonical and intact.
 */
export function beginStreamInsertion(
  editor: DocxEditorInstance,
  context: DirectPlanContext,
  selectedIds: readonly string[] = []
): { readonly success: true; readonly insertion: StreamInsertion } | NativeEditFailure {
  const started = createStreamAnchor(editor, selectedIds);
  if (!started.success) return started;
  const anchor = started.anchor;
  const group = editor.beginHistoryGroup();
  let active = true;
  let owned: readonly string[] = [];
  function end(): void {
    if (!active) return;
    active = false;
    anchor.release();
    group.end();
  }
  const insertion: StreamInsertion = {
    get state() {
      if (anchor.state === 'released' || group.state === 'closed') end();
      return active ? 'active' : 'ended';
    },
    finish: end,
    cancel: end,
    append(fragment) {
      if (insertion.state !== 'active') return stale('The generated insertion has ended.');
      const revision = readEditorRevision(editor, context.documentId, context.epoch);
      const resolved = anchor.resolve();
      if (!revision || !resolved.success) {
        end();
        return resolved.success ? stale('The native document is unavailable.') : resolved;
      }
      const surface = editor.surface!;
      const parent = findNode(surface.session.part(), resolved.point.parentId);
      if (!parent || parent.kind === 'textValue') {
        end();
        return stale('The generation parent disappeared.');
      }
      const start = parent.children.findIndex((node) => node.id === resolved.point.afterId) + 1;
      const retained = parent.children.slice(start, start + owned.length);
      if (JSON.stringify(rootKeys(retained)) !== JSON.stringify(owned))
        return stale('The generated range changed. Finish or cancel before generating again.');
      const next = parent.children[start + owned.length];
      const plan = createEditPlan(revision, [
        {
          kind: 'siblings',
          partName: surface.session.part().name,
          parentId: parent.id,
          beforeId: retained.at(-1)?.id ?? resolved.point.afterId,
          afterId:
            !next || (next.kind !== 'textValue' && next.localName === 'sectPr') ? null : next.id,
          removedIds: [],
          inserted: fragment.inserted,
          ...(fragment.lists ? { lists: fragment.lists } : {}),
          ...(fragment.blocks ? { blocks: fragment.blocks } : {}),
          ...(fragment.styles ? { styles: fragment.styles } : {}),
        },
      ]);
      if (!plan.success) return plan;
      const frames = beginHistoryGroupFrames(editor, group);
      if (!frames) {
        end();
        return stale('The generated insertion history group ended.');
      }
      const result = frames.run(() => applyDirectPlan(editor, plan.plan, context));
      const outcome = frames.finish();
      if (!result.success) return result;
      const current = anchor.resolve();
      const updated = current.success && findNode(surface.session.part(), current.point.parentId);
      if (updated && updated.kind !== 'textValue') {
        const first =
          updated.children.findIndex(
            (node) => node.id === (current.success ? current.point.afterId : '')
          ) + 1;
        const added = updated.children.length - parent.children.length;
        owned = rootKeys(updated.children.slice(first, first + owned.length + added));
      }
      // A local intervening edit or undo can split a gesture. Never silently extend it again.
      if (!current.success || outcome?.kind === 'split') end();
      return { ...result, history: outcome };
    },
  };
  return { success: true, insertion };
}
