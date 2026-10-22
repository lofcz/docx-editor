import {
  collectRevisionSites,
  ensureHyperlinkRelationship,
  findNode,
  parentNodeOf,
  type OoxmlNode,
  type OoxmlPart,
  type RevisionAddress,
} from '@docx-editor.dev/core/store';
import { refuseWrite } from '../diff/fragment-seeds';
import type { CanonicalEdit } from '../diff/plan';
import type { ReviewGroupReceipt } from './groups';
import {
  nativePropertyAddress,
  propertyAddressKey,
  propertyControlRefusal,
} from './property-fallback';
import type { StagedTextCoordinates } from './tracked-stage-coordinates';
import { currentStagePart } from './tracked-stage-text';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

type LinkEdit = Extract<CanonicalEdit, { kind: 'link' | 'unlink' }>;
export const stageParagraphKey = (partName: string, paragraphId: string) =>
  JSON.stringify([partName, paragraphId]);
function paragraphOwner(part: OoxmlPart, node: OoxmlNode): string | null {
  let current: OoxmlNode | null = node;
  for (let depth = 0; current && depth < 256; depth++) {
    if (current.kind === 'paragraph') return current.id;
    current = parentNodeOf(part, current.id);
  }
  return null;
}
export function existingLinkStageParagraph(part: OoxmlPart, linkId: string): string {
  const link = findNode(part, linkId);
  const owner = link?.kind === 'hyperlink' && paragraphOwner(part, link);
  return owner || refuseWrite('The planned hyperlink no longer exists.', 'stale-revision');
}
export function linkStageParagraph(part: OoxmlPart, edit: LinkEdit): string {
  return edit.kind === 'link'
    ? edit.paragraphId
    : existingLinkStageParagraph(part, edit.hyperlinkId);
}
/** Resources and source ranges resolve inside the actual native stage. */
export function stageTrackedLink(
  writer: TrackedStageWriter,
  edit: LinkEdit,
  revision: TrackedStageAttribution,
  coordinates: StagedTextCoordinates
): void {
  const part = currentStagePart(writer, edit.partName),
    id = linkStageParagraph(part, edit);
  const protectedContent = propertyControlRefusal(part, id);
  if (protectedContent) refuseWrite(protectedContent, 'unsupported-operation');
  if (edit.kind === 'unlink') {
    writer.apply(edit.partName, { op: 'removeHyperlink', linkId: edit.hyperlinkId, revision });
    return;
  }
  const slices = coordinates.slices(edit.start, edit.end);
  if (!slices.length) refuseWrite('A proposed hyperlink needs a nonempty native range.');
  let relationshipId: string | undefined;
  if (edit.target.kind === 'external') {
    const url = edit.target.url;
    writer.edit((pkg) => {
      const made = ensureHyperlinkRelationship(pkg, url, edit.partName);
      if (!made) refuseWrite('The native engine refuses this URL.', 'unsupported-operation');
      relationshipId = made.relationshipId;
      return made.pkg;
    });
  }
  writer.apply(edit.partName, {
    op: 'insertHyperlink',
    paragraphId: edit.paragraphId,
    start: slices[0]!.start,
    end: slices.at(-1)!.end,
    ...(edit.target.kind === 'external' ? { relationshipId } : { anchor: edit.target.anchor }),
    ...(edit.title !== null ? { tooltip: edit.title } : {}),
    revision,
  });
}
/** Final receipts exclude source insertion sites retracted by later native link replacements. */
export function coupleLinkStageReceipts(
  writer: TrackedStageWriter,
  receipts: readonly ReviewGroupReceipt[],
  paragraphs: ReadonlySet<string>,
  initial: ReadonlySet<string>
): readonly ReviewGroupReceipt[] {
  const grouped = new Map<string, { partName: string; addresses: Map<string, RevisionAddress> }>();
  const claimed = new Set<string>();
  for (const [partName, part] of writer.currentPackage().parts)
    for (const site of collectRevisionSites(part)) {
      if (initial.has(site.node.id)) continue;
      if (
        site.node.kind !== 'revisionInsert' &&
        site.node.kind !== 'revisionDelete' &&
        site.node.localName !== 'rPrChange'
      )
        continue;
      const owner = paragraphOwner(part, site.node),
        key = stageParagraphKey(partName, owner ?? '');
      if (!owner || !paragraphs.has(key)) continue;
      const address = nativePropertyAddress(site.node);
      if (!address || site.refused)
        refuseWrite('A hyperlink dependency has no native review address.', 'apply-failed');
      const entry = grouped.get(key) ?? { partName, addresses: new Map<string, RevisionAddress>() };
      const identity = propertyAddressKey(partName, address);
      entry.addresses.set(identity, address);
      claimed.add(identity);
      grouped.set(key, entry);
    }
  return [
    ...receipts.flatMap((receipt) => {
      const addresses = receipt.addresses.filter(
        (address) => !claimed.has(propertyAddressKey(receipt.partName, address))
      );
      return addresses.length ? [{ ...receipt, addresses }] : [];
    }),
    ...Array.from(grouped.values(), (entry) => ({
      partName: entry.partName,
      addresses: [...entry.addresses.values()],
    })),
  ];
}
