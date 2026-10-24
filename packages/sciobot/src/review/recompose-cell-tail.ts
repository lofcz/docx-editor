import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  paragraphTextOf,
  type OoxmlNode,
  type OoxmlParagraphNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { ReviewGroupReceipt } from './groups';
import { paragraphOwner, nativePropertyAddress, propertyAddressKey } from './property-fallback';
import { ownedAtomGapInventory } from './recompose-atom-text-continuity';
import { refuseWrite } from '../diff/fragment-seeds';

interface Tail {
  readonly paragraphId: string;
  readonly targetId: string;
  readonly length: number;
  readonly decisions: readonly string[];
  readonly cancelled: readonly string[];
  readonly expected: OoxmlParagraphNode;
}
function plainRuns(paragraph: OoxmlParagraphNode): readonly OoxmlNode[] | null {
  const runs = paragraph.children.filter((n) => n.kind !== 'paragraphProperties');
  return runs.length &&
    runs.every(
      (n) =>
        n.kind === 'run' &&
        n.children.every(
          (c) =>
            (c.kind === 'runProperties' && !c.attributes.length && !c.children.length) ||
            (c.kind === 'text' && c.children.length === 1 && c.children[0]!.kind === 'textValue')
        )
    )
    ? runs
    : null;
}
const textOf = (node: OoxmlNode): string =>
  node.kind === 'textValue' ? node.value : node.children.map(textOf).join('');

/** Physical atom slots establish continuity. Equal text only verifies an unchanged tail. */
export function planCellTailContinuation(
  part: OoxmlPart,
  source: OoxmlParagraphNode,
  baseline: OoxmlParagraphNode,
  target: OoxmlParagraphNode,
  receipts: readonly ReviewGroupReceipt[],
  selected: ReadonlySet<string>,
  sourceType: (id: string) => string | undefined,
  targetType: (id: string) => string | undefined
): Tail | null {
  const runs = plainRuns(baseline);
  if (!runs || !runs.every((run, i) => source.children.at(i - runs.length)?.id === run.id))
    return null;
  const before = ownedAtomGapInventory(source, sourceType, selected),
    after = ownedAtomGapInventory(target, targetType);
  if (!before || !after || JSON.stringify(before.atoms) !== JSON.stringify(after.atoms))
    return null;
  const sourceGap = before.gaps.at(-1)!,
    targetGap = after.gaps.at(-1)!;
  const leaves = runs.flatMap((run) =>
    run.kind === 'run' ? run.children.filter((n) => n.kind === 'text') : []
  );
  if (sourceGap.length !== leaves.length || sourceGap.some((n, i) => n.id !== leaves[i]!.id))
    return null;
  const text = runs.map(textOf).join('');
  if (!text || targetGap.map(textOf).join('') !== text) return null;
  // The complete target suffix must remain ordinary text, with no requested restyling.
  for (const node of targetGap) {
    const run = node.kind === 'text' && findRun(target, node.id);
    if (
      !run ||
      run.kind !== 'run' ||
      !run.children.every(
        (n) =>
          n.kind === 'text' ||
          (n.kind === 'runProperties' && !n.attributes.length && !n.children.length)
      )
    )
      return null;
  }
  const addresses = new Set(
    receipts
      .filter((r) => r.partName === part.name)
      .flatMap((r) => r.addresses.map((a) => propertyAddressKey(part.name, a)))
  );
  const all = collectRevisionSites(part);
  const sites = all.filter((site) => paragraphOwner(part, site.node) === baseline.id);
  const cancelled = new Set<string>();
  if (!sites.length) return null;
  for (const site of sites) {
    const address = nativePropertyAddress(site.node);
    if (
      site.refused ||
      site.node.localName !== 'del' ||
      !address ||
      !addresses.has(propertyAddressKey(part.name, address))
    )
      return null;
    cancelled.add(propertyAddressKey(part.name, address));
  }
  if (
    all.some((site) => {
      const address = nativePropertyAddress(site.node);
      return (
        address &&
        cancelled.has(propertyAddressKey(part.name, address)) &&
        paragraphOwner(part, site.node) !== baseline.id
      );
    })
  )
    return null;
  return {
    paragraphId: baseline.id,
    targetId: target.id,
    length: text.length,
    decisions: sites.map((site) => site.node.id),
    cancelled: [...cancelled],
    expected: baseline,
  };
}
function findRun(root: OoxmlNode, textId: string): OoxmlNode | null {
  if (root.kind === 'textValue') return null;
  if (root.kind === 'run' && root.children.some((n) => n.id === textId)) return root;
  for (const child of root.children) {
    const found = findRun(child, textId);
    if (found) return found;
  }
  return null;
}

/** Withdraw only this stage's redundant deletion and duplicate insertion of committed text. */
export function continueCellTail(part: OoxmlPart, plan: Tail, cancelled: Set<string>): OoxmlPart {
  const restored = applyTreeOp(part, { op: 'rejectAllRevisions', siteNodeIds: plan.decisions });
  if (!restored.ok) refuseWrite('Cannot restore the committed cell tail.', 'apply-failed');
  const node = findNode(restored.part, plan.paragraphId);
  if (!node || canonicalOoxmlFingerprint(node) !== canonicalOoxmlFingerprint(plan.expected))
    refuseWrite('The committed cell tail no longer matches its baseline.', 'ambiguous-target');
  const length = paragraphTextOf(restored.part, plan.targetId)?.length;
  if (length === undefined || length < plan.length)
    refuseWrite('The replacement cell tail disappeared.', 'apply-failed');
  const removed = applyTreeOp(restored.part, {
    op: 'deleteText',
    paragraphId: plan.targetId,
    start: length - plan.length,
    end: length,
  });
  if (!removed.ok) refuseWrite('Cannot remove the duplicate cell tail.', 'apply-failed');
  plan.cancelled.forEach((key) => cancelled.add(key));
  return removed.part;
}
