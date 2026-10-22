import { findNode, type OoxmlPart } from '@docx-editor.dev/core/store';
import { diffArrays } from 'diff';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { jsonDataProblem } from '../data/json';
import type { TargetAstNode } from '../projection/parse-target';
import { createEditPlan, type CanonicalEdit, type CanonicalEditPlan } from './plan';
import { isTargetFragment } from './target-fragment';
import { targetAstKey as siblingAstKey } from './target-key';

export { targetAstKey as siblingAstKey } from './target-key';

export interface MappedSiblingBlock {
  readonly id: string;
  /** Both sides use the same normalized target parser. Native trees never enter this AST. */
  readonly ast: TargetAstNode;
}
export interface SiblingBlockSource {
  readonly revision: RevisionToken;
  readonly part: OoxmlPart;
  readonly parentId: string;
  /** A contiguous window of actual canonical children, not a flattened projection. */
  readonly startIndex: number;
  readonly blocks: readonly MappedSiblingBlock[];
}
export interface SiblingBlockPair {
  readonly id: string;
  readonly beforeIndex: number;
  readonly afterIndex: number;
  /** D15 must route changed pairs to content/property planners. Structural plans do not apply them. */
  readonly changed: boolean;
}
export interface SiblingBlockGap {
  readonly operationIndex: number;
  /** Indices in the supplied sibling arrays, before JSON plan detachment. */
  readonly beforeStart: number;
  readonly beforeEnd: number;
  readonly afterStart: number;
  readonly afterEnd: number;
}
/** Trusted application policy. Callbacks never enter serialized plans or model input. */
export interface SiblingPairingPolicy {
  readonly key?: (node: TargetAstNode) => string;
  readonly pairable?: (before: TargetAstNode, after: TargetAstNode) => boolean;
}
export type SiblingBlockResult =
  | {
      readonly success: true;
      readonly plan: CanonicalEditPlan;
      readonly pairs: readonly SiblingBlockPair[];
      readonly gaps: readonly SiblingBlockGap[];
    }
  | NativeEditFailure;

export const SIBLING_BLOCK_LIMITS = Object.freeze({
  maxBytes: 16 * 1024 * 1024,
  maxNodes: 200_000,
  maxDepth: 64,
  maxSiblings: 10_000,
  diffTimeoutMs: 1000,
});
const parents = new Set([
  'body',
  'tc',
  'sdtContent',
  'hdr',
  'ftr',
  'footnote',
  'endnote',
  'txbxContent',
]);
const containers = new Set(['blockquote', 'list', 'listItem', 'table', 'tableRow', 'tableCell']);
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });

function textOf(node: TargetAstNode): string {
  if (['text', 'inlineCode', 'code'].includes(node.type) && typeof node.value === 'string')
    return node.value;
  return Array.isArray(node.children)
    ? (node.children as TargetAstNode[]).map(textOf).join('')
    : '';
}
function similarText(before: string, after: string): boolean {
  if (before === after) return true;
  if (!before || !after) return false;
  let leading = 0,
    trailing = 0;
  while (leading < Math.min(before.length, after.length) && before[leading] === after[leading])
    leading++;
  while (
    trailing < Math.min(before.length, after.length) - leading &&
    before[before.length - trailing - 1] === after[after.length - trailing - 1]
  )
    trailing++;
  if ((leading + trailing) / Math.max(before.length, after.length) >= 0.3) return true;
  const a = before.match(/\S+/g) ?? [],
    b = after.match(/\S+/g) ?? [];
  const counts = new Map<string, number>();
  for (const word of a) counts.set(word, (counts.get(word) ?? 0) + 1);
  let shared = 0;
  for (const word of b) {
    const count = counts.get(word) ?? 0;
    if (count > 0) {
      shared++;
      counts.set(word, count - 1);
    }
  }
  return (2 * shared) / Math.max(1, a.length + b.length) >= 0.3;
}
export function siblingContentPairable(before: TargetAstNode, after: TargetAstNode): boolean {
  if (before.type !== after.type) return false;
  if (containers.has(before.type)) return true;
  return (
    ['paragraph', 'heading', 'code'].includes(before.type) &&
    similarText(textOf(before), textOf(after))
  );
}

/**
 * Plan bounded structural gaps while retaining exact siblings and compatible content pairs.
 * The native source is read-only. New fragments contain only parsed target data.
 * Operations use original canonical anchors and run from the last gap to the first.
 * D12 owns registered MDX identity pairing; call this planner on each retained container's children.
 * Executors must validate current revisions, anchors, protection, and native structural invariants.
 */
export function planSiblingBlocks(
  source: SiblingBlockSource,
  target: readonly TargetAstNode[],
  policy: SiblingPairingPolicy = {}
): SiblingBlockResult {
  if (!source || typeof source !== 'object' || Array.isArray(source))
    return fail('Missing sibling source.');
  const descriptors = Object.getOwnPropertyDescriptors(source);
  for (const field of ['revision', 'part', 'parentId', 'startIndex', 'blocks']) {
    const descriptor = descriptors[field];
    if (!descriptor || !('value' in descriptor) || !descriptor.enumerable)
      return fail('Sibling source fields must be plain values.');
  }
  const data = {
    revision: source.revision,
    parentId: source.parentId,
    startIndex: source.startIndex,
    blocks: source.blocks,
    target,
  };
  const problem = jsonDataProblem(data, SIBLING_BLOCK_LIMITS);
  if (problem) return fail(problem);
  if (
    !Array.isArray(source.blocks) ||
    !Array.isArray(target) ||
    source.blocks.length > SIBLING_BLOCK_LIMITS.maxSiblings ||
    target.length > SIBLING_BLOCK_LIMITS.maxSiblings ||
    !Number.isSafeInteger(source.startIndex) ||
    source.startIndex < 0 ||
    typeof source.parentId !== 'string'
  )
    return fail('Invalid or excessive sibling window.');
  const raw = JSON.stringify(data);
  if (
    raw.length > SIBLING_BLOCK_LIMITS.maxBytes ||
    new TextEncoder().encode(raw).byteLength > SIBLING_BLOCK_LIMITS.maxBytes
  )
    return fail('Sibling input exceeds its byte limit.');
  if (
    source.blocks.some(
      (block) => !block || typeof block.id !== 'string' || !block.id || !isTargetFragment(block.ast)
    ) ||
    !target.every(isTargetFragment)
  )
    return fail('Sibling blocks need parsed target AST fragments.');
  if (new Set(source.blocks.map((block) => block.id)).size !== source.blocks.length)
    return fail('Sibling IDs are not unique.', 'ambiguous-target');
  const parent = findNode(source.part, source.parentId);
  if (
    !parent ||
    parent.kind === 'textValue' ||
    parent.namespaceUri !== 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' ||
    !parents.has(parent.localName)
  )
    return fail(
      'The native block parent does not exist or cannot contain these blocks.',
      'invalid-document'
    );
  if (
    source.startIndex + source.blocks.length > parent.children.length ||
    source.blocks.some(
      (block, index) => parent.children[source.startIndex + index]?.id !== block.id
    )
  )
    return fail(
      'Sibling addresses do not describe contiguous native children.',
      'ambiguous-target'
    );
  if (
    source.blocks.some((_, index) => {
      const child = parent.children[source.startIndex + index]!;
      return (
        child.kind !== 'paragraph' &&
        child.kind !== 'table' &&
        !(
          child.kind !== 'textValue' &&
          child.localName === 'sdt' &&
          child.namespaceUri === parent.namespaceUri
        )
      );
    })
  )
    return fail('The mapped window contains unprojected native content.', 'unsupported-operation');
  const before = source.blocks.map((block) => block.ast);
  const key = policy.key ?? siblingAstKey,
    canPair = policy.pairable ?? siblingContentPairable;
  const oldKeys = before.map(key),
    newKeys = target.map(key);
  const changes = diffArrays(oldKeys, newKeys, { timeout: SIBLING_BLOCK_LIMITS.diffTimeoutMs });
  if (!changes)
    return fail('Sibling matching exceeded its planning budget.', 'unsupported-operation');
  const pairs: SiblingBlockPair[] = [],
    ops: CanonicalEdit[] = [],
    gaps: Omit<SiblingBlockGap, 'operationIndex'>[] = [];
  const pair = (a: number, b: number) =>
    pairs.push(
      Object.freeze({
        id: source.blocks[a]!.id,
        beforeIndex: a,
        afterIndex: b,
        changed: oldKeys[a] !== newKeys[b],
      })
    );
  const gap = (oldStart: number, oldEnd: number, newStart: number, newEnd: number) => {
    while (oldStart < oldEnd && newStart < newEnd && canPair(before[oldStart]!, target[newStart]!))
      pair(oldStart++, newStart++);
    const tail: [number, number][] = [];
    while (
      oldStart < oldEnd &&
      newStart < newEnd &&
      canPair(before[oldEnd - 1]!, target[newEnd - 1]!)
    )
      tail.push([--oldEnd, --newEnd]);
    if (oldStart < oldEnd || newStart < newEnd) {
      ops.push({
        kind: 'siblings',
        partName: source.part.name,
        parentId: source.parentId,
        beforeId: parent.children[source.startIndex + oldStart - 1]?.id ?? null,
        afterId: parent.children[source.startIndex + oldEnd]?.id ?? null,
        removedIds: source.blocks.slice(oldStart, oldEnd).map((block) => block.id),
        inserted: target.slice(newStart, newEnd),
      });
      gaps.push({
        beforeStart: oldStart,
        beforeEnd: oldEnd,
        afterStart: newStart,
        afterEnd: newEnd,
      });
    }
    for (const [a, b] of tail.reverse()) pair(a, b);
  };
  let oldIndex = 0,
    newIndex = 0,
    oldGap = 0,
    newGap = 0;
  for (const change of changes) {
    if (change.added) newIndex += change.count;
    else if (change.removed) oldIndex += change.count;
    else {
      gap(oldGap, oldIndex, newGap, newIndex);
      for (let i = 0; i < change.count; i++) pair(oldIndex++, newIndex++);
      oldGap = oldIndex;
      newGap = newIndex;
    }
  }
  gap(oldGap, oldIndex, newGap, newIndex);
  const result = createEditPlan(source.revision, ops.reverse());
  if (!result.success) return result;
  pairs.sort((a, b) => a.beforeIndex - b.beforeIndex);
  return {
    success: true,
    plan: result.plan,
    pairs: Object.freeze(pairs),
    gaps: Object.freeze(
      gaps.reverse().map((gap, operationIndex) => Object.freeze({ ...gap, operationIndex }))
    ),
  };
}
