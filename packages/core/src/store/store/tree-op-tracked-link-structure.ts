import { ownsInlineRevision } from './inline-transaction-state.ts';
import {
  createNodeIdAllocator,
  findNode,
  replaceChildren,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import { runWithoutJournalCapture } from '../package/canonical-primitive-capture.ts';
import { WML_NAMESPACE_URI, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { build, revisionAttributes } from './tree-op-tracked-builders.ts';
import { cloneWithNewIds, fromEdit, parentOf, TEXT_DEPS } from './tree-op-nodes.ts';
import { paragraphOffsetIndex, type Segment } from './tree-op-segments.ts';
import { invalidRevisionAttribution } from './tree-op-revision-attribution.ts';
import { nextRevisionId } from './tree-op-revision-ids.ts';
import { revisionKey } from './tree-op-tracked-adjacency.ts';
import type { TreeDocOp, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';

type LinkOp = Extract<
  TreeDocOp,
  { op: 'insertHyperlink' | 'removeHyperlink' | 'setHyperlinkTarget' }
>;
type TrackedLinkOp = LinkOp & { readonly revision: NonNullable<LinkOp['revision']> };
type Apply = (part: OoxmlPart, op: TreeDocOp) => TreeOpResult;
type Divide = (
  child: OoxmlNode,
  offsets: readonly number[],
  pieces: number,
  segments: readonly Segment[],
  mint: () => string
) => OoxmlNode[][];
const kids = (node: OoxmlNode) => (node.kind === 'textValue' ? [] : node.children);
const word = (node: OoxmlNode, name: string) =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && node.localName === name;
const proof = (node: OoxmlNode) => word(node, 'proofErr');
const COPYABLE = new Set([
  'rPr',
  't',
  'tab',
  'br',
  'cr',
  'noBreakHyphen',
  'softHyphen',
  'sym',
  'lastRenderedPageBreak',
]);
function copyableRun(node: OoxmlNode, deleted = false): boolean {
  if (node.kind !== 'run') return false;
  const revision = (node: OoxmlNode): boolean =>
    node.kind !== 'textValue' &&
    node.namespaceUri === WML_NAMESPACE_URI &&
    (/PrChange$/.test(node.localName) ||
      ['ins', 'del', 'moveFrom', 'moveTo'].includes(node.localName));
  const pending: OoxmlNode[] = [node];
  while (pending.length) {
    const next = pending.pop()!;
    if (revision(next)) return false;
    for (const child of kids(next)) pending.push(child);
  }
  return node.children.every(
    (child) =>
      child.namespaceUri === WML_NAMESPACE_URI &&
      (COPYABLE.has(child.localName) || (deleted && child.kind === 'deletedText'))
  );
}
function copyableSource(node: OoxmlNode, options?: EditOptions, deleted = false): boolean {
  if (proof(node) || copyableRun(node, deleted) || deletedCurrentLink(node, options)) return true;
  return (
    ownsInlineRevision(node, options?.trackedRevisionIds) &&
    kids(node).every((child) =>
      copyableSource(child, options, deleted || node.kind === 'revisionDelete')
    )
  );
}
/** A link whose entire visible source is already deleted by this exact transaction. */
function deletedCurrentLink(node: OoxmlNode, options?: EditOptions): boolean {
  return (
    node.kind === 'hyperlink' &&
    node.children.some((child) => child.kind === 'revisionDelete') &&
    node.children.every(
      (child) =>
        proof(child) ||
        (child.kind === 'revisionDelete' &&
          ownsInlineRevision(child, options?.trackedRevisionIds) &&
          copyableSource(child, options))
    )
  );
}
/** Keep deletion-only source links outside a new wrapper; they remain visible in Original. */
function insertionPreviewSource(
  part: OoxmlPart,
  op: Extract<TrackedLinkOp, { op: 'insertHyperlink' }>,
  options?: EditOptions
): { part: OoxmlPart; op: TreeDocOp } {
  const paragraph = findNode(part, op.paragraphId);
  const { revision: _revision, ...plain } = op;
  if (paragraph?.kind !== 'paragraph') return { part, op: plain };
  const index = paragraphOffsetIndex(paragraph);
  let removed = 0;
  const children = paragraph.children.filter((child) => {
    if (!deletedCurrentLink(child, options)) return true;
    const span = index.spanOf(child);
    if (!span || span.start < op.start || span.end > op.end) return true;
    removed += span.end - span.start;
    return false;
  });
  if (!removed) return { part, op: plain };
  const made = replaceChildren(part, paragraph.id, children);
  if (!made.ok) throw Error('Validated hyperlink preview source cannot be projected.');
  return { part: made.part, op: { ...plain, end: plain.end - removed } };
}
/** Project only validated source or its native split preview; copied wrapper IDs can differ. */
export function proposedHyperlinkContent(nodes: readonly OoxmlNode[]): readonly OoxmlNode[] {
  const result: OoxmlNode[] = [];
  for (const node of nodes) {
    if (node.kind === 'revisionInsert' || node.kind === 'revisionDelete') {
      if (node.kind === 'revisionInsert')
        for (const child of proposedHyperlinkContent(kids(node))) result.push(child);
    } else result.push(node);
  }
  return result;
}
export function isTrackedHyperlinkStructure(op: TreeDocOp, part?: OoxmlPart): op is TrackedLinkOp {
  if (!('revision' in op) || op.revision === undefined) return false;
  if (op.op === 'insertHyperlink' || op.op === 'removeHyperlink') return true;
  if (op.op !== 'setHyperlinkTarget' || !part) return false;
  const link = findNode(part, op.linkId);
  return (
    !!link &&
    kids(link).some((child) => child.kind === 'revisionInsert' || child.kind === 'revisionDelete')
  );
}
/** Plain validation runs first. Source revisions and non-copyable atoms never become fresh ownership. */
export function trackedHyperlinkStructureRefusal(
  part: OoxmlPart,
  op: TrackedLinkOp,
  options?: EditOptions
): TreeOpRejection | null {
  if (invalidRevisionAttribution(op.revision)) return 'invalidArgs';
  if (op.op !== 'insertHyperlink') {
    if (op.range !== undefined) return 'unsupported';
    const link = findNode(part, op.linkId),
      paragraph = link && parentOf(part, link.id);
    if (
      link?.kind !== 'hyperlink' ||
      paragraph?.kind !== 'paragraph' ||
      !link.children.some(
        (child) => child.kind === 'run' || ownsInlineRevision(child, options?.trackedRevisionIds)
      )
    )
      return 'unsupported';
    return link.children.every((child) => copyableSource(child, options)) ? null : 'unsupported';
  }
  const paragraph = findNode(part, op.paragraphId);
  if (paragraph?.kind !== 'paragraph') return 'unknown-paragraph';
  const index = paragraphOffsetIndex(paragraph);
  let cursor = 0;
  for (const child of paragraph.children) {
    if (word(child, 'pPr')) continue;
    const span = index.spanOf(child);
    if (!span || span.start === span.end) {
      if (cursor >= op.start && cursor < op.end && !proof(child)) return 'unsupported';
      continue;
    }
    cursor = span.end;
    if (span.end <= op.start || span.start >= op.end) continue;
    if (!copyableSource(child, options)) return 'unsupported';
  }
  return null;
}

/** Dispatch the tracked branch without expanding the common native validator. */
export function validateTrackedHyperlinkStructure(
  part: OoxmlPart,
  op: TreeDocOp,
  options: EditOptions | undefined,
  validate: (part: OoxmlPart, op: TreeDocOp, options?: EditOptions) => TreeOpRejection | null
): TreeOpRejection | null | undefined {
  if (!isTrackedHyperlinkStructure(op, part)) return undefined;
  const { revision: _revision, ...plain } = op;
  const ordinary = validate(part, plain, options);
  if (
    ordinary === 'invalid-property-value' &&
    op.op === 'insertHyperlink' &&
    !trackedHyperlinkStructureRefusal(part, op, options)
  ) {
    return runWithoutJournalCapture(() => {
      const projected = insertionPreviewSource(part, op, options);
      return projected.part === part ? ordinary : validate(projected.part, projected.op, options);
    });
  }
  return ordinary ?? trackedHyperlinkStructureRefusal(part, op, options);
}

/** Native replacement markup: old runs stay struck; the linked/unlinked insertion follows. */
export function applyTrackedHyperlinkStructure(
  part: OoxmlPart,
  op: TrackedLinkOp,
  apply: Apply,
  divide: Divide,
  options?: EditOptions
): TreeOpResult {
  const refusal = trackedHyperlinkStructureRefusal(part, op, options);
  if (refusal) return { ok: false, reason: refusal };
  const mint = createNodeIdAllocator(part),
    ids = options?.trackedRevisionIds,
    revisionId = ids ? () => ids.mint() : nextRevisionId(part);
  let insertionId: string | undefined, deletionId: string | undefined;
  const insertion = () => (insertionId ??= revisionId());
  const deletion = () => (deletionId ??= revisionId());
  const inserted = (children: readonly OoxmlNode[]) =>
    build(
      mint(),
      'revisionInsert',
      'ins',
      revisionAttributes(insertion(), op.revision),
      children.map((child) => cloneWithNewIds(child, mint))
    );
  const strike = (nodes: readonly OoxmlNode[]): OoxmlNode[] => {
    const result: OoxmlNode[] = [];
    let runs: OoxmlNode[] = [];
    const flush = () => {
      if (!runs.length) return;
      result.push(
        build(mint(), 'revisionDelete', 'del', revisionAttributes(deletion(), op.revision), runs)
      );
      runs = [];
    };
    for (const node of nodes) {
      if (node.kind === 'hyperlink' || node.kind === 'revisionDelete') {
        flush();
        result.push(node);
        continue;
      }
      if (node.kind === 'revisionInsert') continue;
      const copied = cloneWithNewIds(node, mint);
      if (proof(copied)) {
        runs.push(copied);
        continue;
      }
      if (copied.kind !== 'run') throw new Error('Validated hyperlink source is not a run.');
      runs.push({
        ...copied,
        children: copied.children.map((child) =>
          child.kind === 'text'
            ? ({ ...child, kind: 'deletedText', localName: 'delText' } as OoxmlNode)
            : child
        ),
      } as OoxmlNode);
    }
    flush();
    return result;
  };
  let paragraph: Extract<OoxmlNode, { kind: 'paragraph' }>,
    children: readonly OoxmlNode[],
    proposal: OoxmlNode;
  if (op.op !== 'insertHyperlink') {
    const link = findNode(part, op.linkId)!;
    if (link.kind !== 'hyperlink') return { ok: false, reason: 'tree-invariant' };
    const parent = parentOf(part, link.id);
    if (parent?.kind !== 'paragraph') return { ok: false, reason: 'tree-invariant' };
    paragraph = parent;
    const content = proposedHyperlinkContent(link.children);
    if (!content.some((child) => !proof(child)))
      return {
        ok: true,
        part,
        effect: {
          dirty: [],
          created: [],
          deleted: [],
          dependencyKeys: [],
          impact: 'paragraph-local',
        },
      };
    if (op.op === 'setHyperlinkTarget') {
      const { revision: _revision, ...plain } = op;
      const preview = runWithoutJournalCapture(() => apply(part, plain));
      if (!preview.ok) return preview;
      const target = findNode(preview.part, link.id);
      if (target?.kind !== 'hyperlink') return { ok: false, reason: 'tree-invariant' };
      const same =
        target.attributes.length === link.attributes.length &&
        target.attributes.every((attribute) =>
          link.attributes.some(
            (old) =>
              old.namespaceUri === attribute.namespaceUri &&
              old.localName === attribute.localName &&
              old.value === attribute.value
          )
        );
      if (same) return { ok: true, part, effect: { ...preview.effect, dirty: [] } };
      proposal = { ...target, id: mint(), children: [inserted(content)] } as OoxmlNode;
    } else proposal = inserted(content);
    const old = { ...link, children: strike(link.children) } as OoxmlNode;
    children = paragraph.children.flatMap((child) =>
      child.id === link.id ? (kids(old).length ? [old, proposal] : [proposal]) : [child]
    );
  } else {
    const preview = runWithoutJournalCapture(() => {
      const projected = insertionPreviewSource(part, op, options);
      return apply(projected.part, projected.op);
    });
    if (!preview.ok) return preview;
    const current = findNode(part, op.paragraphId),
      next = findNode(preview.part, op.paragraphId);
    if (current?.kind !== 'paragraph' || next?.kind !== 'paragraph')
      return { ok: false, reason: 'tree-invariant' };
    paragraph = current;
    const link = next.children.find(
      (child) => child.kind === 'hyperlink' && !findNode(part, child.id)
    );
    if (link?.kind !== 'hyperlink') return { ok: false, reason: 'tree-invariant' };
    const content = proposedHyperlinkContent(link.children);
    if (!content.some((child) => !proof(child))) return { ok: false, reason: 'unsupported' };
    proposal = {
      ...link,
      id: mint(),
      children: [inserted(content)],
    } as OoxmlNode;
    const index = paragraphOffsetIndex(paragraph),
      before: OoxmlNode[] = [],
      selected: OoxmlNode[] = [],
      after: OoxmlNode[] = [];
    let cursor = 0;
    for (const child of paragraph.children) {
      if (word(child, 'pPr')) {
        before.push(child);
        continue;
      }
      const span = index.spanOf(child);
      if (!span || span.start === span.end) {
        (cursor < op.start ? before : cursor < op.end ? selected : after).push(child);
        continue;
      }
      cursor = span.end;
      if (span.end <= op.start) {
        before.push(child);
        continue;
      }
      if (span.start >= op.end) {
        after.push(child);
        continue;
      }
      if (span.start >= op.start && span.end <= op.end) {
        selected.push(child);
        continue;
      }
      const pieces = divide(child, [op.start, op.end], 3, index.segments, mint);
      for (const piece of pieces[0]!) before.push(piece);
      for (const piece of pieces[1]!) selected.push(piece);
      for (const piece of pieces[2]!) after.push(piece);
    }
    children = [...before, ...strike(selected), proposal, ...after];
  }
  const result = fromEdit(replaceChildren(part, paragraph.id, children, options), {
    dirty: [paragraph.id],
    created: [proposal.id],
    deleted: [],
    dependencyKeys: TEXT_DEPS,
    impact: 'paragraph-local',
  });
  if (result.ok) {
    if (deletionId !== undefined)
      ids?.wrote(revisionKey(deletionId, op.revision.author, op.revision.date));
    if (insertionId !== undefined)
      ids?.wrote(revisionKey(insertionId, op.revision.author, op.revision.date));
  }
  return result;
}
