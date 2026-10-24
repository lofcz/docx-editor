import { inlineContainersOf } from './tree-op-nodes.ts';
import { findNode, type EditOptions } from '../package/ooxml-edit.ts';
import type { OoxmlNode, OoxmlParagraphNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { segmentAncestryNodeId, type InsertionSite } from './tree-op-segments.ts';
import { applyInsertCommittedPayload } from './tree-op-tracked.ts';
import type { TreeOpResult } from './tree-op-types.ts';

/** Untracked content does not join an existing insertion's review ownership. */
export function committedInsertionWithinRevision(
  part: OoxmlPart,
  paragraph: OoxmlParagraphNode,
  offset: number,
  builders: readonly ((mint: () => string) => OoxmlNode)[],
  length: number,
  site: InsertionSite,
  options: EditOptions | undefined,
  owner: OoxmlNode | null,
  bias: 'left' | 'right'
): TreeOpResult | null {
  if (length === 0) return null;
  const landing =
    site.kind === 'withinValue' || site.kind === 'atBoundary'
      ? segmentAncestryNodeId(site.segment)
      : site.kind === 'newRun'
        ? site.holder.id
        : site.run.id;
  const path = inlineContainersOf(paragraph, landing);
  if (
    !path.some((node) => node.kind === 'revisionInsert') &&
    findNode(part, landing)?.kind !== 'revisionInsert'
  )
    return null;
  const nodes = (mint: () => string) => builders.map((builder) => builder(mint));
  return applyInsertCommittedPayload(
    part,
    paragraph,
    offset,
    { length, nodes },
    options,
    bias,
    owner
  );
}
