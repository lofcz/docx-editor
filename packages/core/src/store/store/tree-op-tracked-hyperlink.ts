import { copyTrackedText } from './tracked-text-copy.ts';
// Tracked hyperlink RETARGETING, in Word's content revision markup.
//
// Word has no property-change record for a link's target or tooltip: re-aiming a link while
// tracking replaces it. So does this writer. The link keeps its old target and its content is
// struck (`w:del` inside the link); a copy of the link with the new target follows it, holding
// the same runs as a tracked insertion. The struck link keeps the original characters'
// collaboration identity, so a peer deleting that text concurrently deletes the struck half,
// and the copy remains as the pending proposal. Accepting drops the struck link, which the resolution
// removes once it holds no runs, and keeps the copy; rejecting drops the copy and restores the
// original link and runs. Relationships are never touched: another link sharing the old
// target keeps it, and the caller mints the new one first, as the untracked command requires.
//
// The run content is copied exactly, properties and unknown children included, so the text and
// its marks read the same in both views. Content a copy cannot duplicate safely — fields, note
// references, drawings, nested revisions, and range markers such as bookmarks, which would
// collapse beside the accepted copy — refuses before any write. Proofing marks stay with the
// struck link and are hoisted out of it when accepting empties it.
import {
  createNodeIdAllocator,
  findNode,
  replaceChildren,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import { runWithoutJournalCapture } from '../package/canonical-primitive-capture.ts';
import { WML_NAMESPACE_URI, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { build, revisionAttributes } from './tree-op-tracked-builders.ts';
import { fromEdit, ok, parentOf, TEXT_DEPS } from './tree-op-nodes.ts';
import { invalidRevisionAttribution } from './tree-op-revision-attribution.ts';
import { nextRevisionId } from './tree-op-revision-ids.ts';
import { paragraphOffsetIndex } from './tree-op-segments.ts';
import type { RevisionAttributionInput, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';
import type { TreeDocOp } from './tree-ops.ts';

type TargetOp = Extract<TreeDocOp, { op: 'setHyperlinkTarget' }>;
type TrackedTargetOp = TargetOp & { readonly revision: RevisionAttributionInput };
type Apply = (part: OoxmlPart, op: TreeDocOp) => TreeOpResult;

/** Run children a copy duplicates without minting a second field, note, or drawing. */
const COPYABLE_RUN_CHILDREN = new Set([
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
const copyableRun = (node: OoxmlNode): boolean =>
  node.kind === 'run' &&
  (node.children as readonly OoxmlNode[]).every(
    (child) =>
      child.kind === 'textValue' ||
      (child.namespaceUri === WML_NAMESPACE_URI && COPYABLE_RUN_CHILDREN.has(child.localName))
  );

const proofingMark = (node: OoxmlNode): boolean =>
  node.kind !== 'textValue' &&
  node.namespaceUri === WML_NAMESPACE_URI &&
  node.localName === 'proofErr';

/** Refusals specific to the tracked form; the untracked validator runs first. */
export function trackedHyperlinkRefusal(
  part: OoxmlPart,
  op: TrackedTargetOp
): TreeOpRejection | null {
  if (invalidRevisionAttribution(op.revision)) return 'invalidArgs';
  // A subrange retarget splits the link first; its tracked form is not modelled here.
  if (op.range !== undefined) return 'unsupported';
  const link = findNode(part, op.linkId);
  if (link?.kind !== 'hyperlink') return 'tree-invariant';
  const paragraph = parentOf(part, link.id);
  if (paragraph?.kind !== 'paragraph') return 'unsupported';
  if (!link.children.some((child) => child.kind === 'run')) return 'unsupported';
  // Proofing marks travel with the struck link; a range marker would collapse on accept.
  if (!link.children.every((child) => proofingMark(child) || copyableRun(child)))
    return 'unsupported';
  const span = paragraphOffsetIndex(paragraph).spanOf(link);
  return span && span.end > span.start ? null : 'unsupported';
}

/**
 * Apply a tracked retarget. `apply` is the native applier: the untracked command computes the
 * new link attributes, and a tracked `deleteText` strikes the old link's content, so the
 * deletion follows every native tracked-deletion rule.
 */
export function applyTrackedHyperlinkTarget(
  part: OoxmlPart,
  op: TrackedTargetOp,
  apply: Apply,
  options?: EditOptions
): TreeOpResult {
  const refusal = trackedHyperlinkRefusal(part, op);
  if (refusal) return { ok: false, reason: refusal };
  const { revision, ...plain } = op;
  const link = findNode(part, op.linkId)!;
  if (link.kind !== 'hyperlink') return { ok: false, reason: 'tree-invariant' };
  const paragraph = parentOf(part, link.id)!;
  // Only the new attributes are read from the untracked command; its splices must not reach the
  // shared journal, or peers would see the old link re-aimed outright.
  const retargeted = runWithoutJournalCapture(() => apply(part, plain));
  if (!retargeted.ok) return retargeted;
  const next = findNode(retargeted.part, link.id);
  if (next?.kind !== 'hyperlink') return { ok: false, reason: 'tree-invariant' };
  const same =
    next.attributes.length === link.attributes.length &&
    next.attributes.every((attribute) =>
      link.attributes.some(
        (old) =>
          old.namespaceUri === attribute.namespaceUri &&
          old.localName === attribute.localName &&
          old.value === attribute.value
      )
    );
  // A retarget that names the current target and tooltip proposes nothing.
  if (same) return ok(part, { ...retargeted.effect, dirty: [] });
  const mint = createNodeIdAllocator(part);
  const insertionId = options?.trackedRevisionIds?.mint() ?? nextRevisionId(part)();
  const copied = link.children.filter((child) => !proofingMark(child));
  const proposal = Object.freeze({
    ...next,
    id: mint(),
    children: [
      build(
        mint(),
        'revisionInsert',
        'ins',
        revisionAttributes(insertionId, revision),
        copyTrackedText(copied, copied, mint, false)
      ),
    ],
  }) as OoxmlNode;
  const placed = replaceChildren(
    part,
    paragraph.id,
    (paragraph.children as readonly OoxmlNode[]).flatMap((child) =>
      child.id === link.id ? [child, proposal] : [child]
    ),
    options
  );
  if (!placed.ok)
    return fromEdit(placed, {
      dirty: [paragraph.id],
      created: [],
      deleted: [],
      dependencyKeys: TEXT_DEPS,
      impact: 'paragraph-local',
    });
  const placedParagraph = findNode(placed.part, paragraph.id) as Extract<
    OoxmlNode,
    { kind: 'paragraph' }
  >;
  const span = paragraphOffsetIndex(placedParagraph).spanOf(link.id)!;
  const struck = apply(placed.part, {
    op: 'deleteText',
    paragraphId: paragraph.id,
    start: span.start,
    end: span.end,
    revision,
  });
  if (!struck.ok) return struck;
  // The old link must keep its target and hold only its struck content.
  const old = findNode(struck.part, link.id);
  if (old?.kind !== 'hyperlink' || old.children.some((child) => child.kind === 'run'))
    return { ok: false, reason: 'unsupported' };
  return {
    ...struck,
    effect: {
      ...struck.effect,
      dirty: [...new Set([...struck.effect.dirty, paragraph.id])],
      created: [...new Set([...struck.effect.created, proposal.id])],
    },
  };
}
