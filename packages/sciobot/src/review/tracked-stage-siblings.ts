import {
  collectRevisionSites,
  contentControlsIn,
  directParagraphProperties,
  findNode,
  paragraphTextOf,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
  type RevisionAddress,
} from '@docx-editor.dev/core/store';
import { applyDirectListAssignments, type DirectListContext } from '../ai/direct-lists';
import { newCellListAssignments } from '../diff/new-cell-list-assignments';
import { nodeIds, refuseWrite } from '../diff/fragment-seeds';
import type { CanonicalEdit } from '../diff/plan';
import type { NativeSiblingProposal } from './track-blocks';
import { withBornChecklistOwnership } from './tracked-born-checklists';
import { newParagraphSeeds } from './new-paragraphs';
import { stageParagraphResources } from './tracked-stage-paragraphs';
import {
  nativePropertyAddress,
  paragraphOwner,
  propertyAddressKey,
  propertyControlRefusal,
} from './property-fallback';
import { currentStagePart } from './tracked-stage-text';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

type Siblings = Extract<CanonicalEdit, { kind: 'siblings' }>;
const section = (node: OoxmlNode): boolean =>
  node.kind !== 'textValue' &&
  ((node.namespaceUri === WML_NAMESPACE_URI && node.localName === 'sectPr') ||
    node.children.some(section));
function gap(part: OoxmlPart, edit: Siblings): readonly string[] | null {
  const parent = findNode(part, edit.parentId);
  if (
    !parent ||
    parent.kind === 'textValue' ||
    parent.namespaceUri !== WML_NAMESPACE_URI ||
    !['body', 'tc', 'sdtContent'].includes(parent.localName)
  )
    return null;
  const before =
    edit.beforeId === null ? -1 : parent.children.findIndex((node) => node.id === edit.beforeId);
  const terminal = parent.children.findIndex(
    (node) =>
      node.kind !== 'textValue' &&
      node.namespaceUri === WML_NAMESPACE_URI &&
      node.localName === 'sectPr'
  );
  const after =
    edit.afterId === null
      ? terminal < 0
        ? parent.children.length
        : terminal
      : parent.children.findIndex((node) => node.id === edit.afterId);
  if ((edit.beforeId !== null && before < 0) || after < 0 || before >= after) return null;
  return parent.children.slice(before + 1, after).map((node) => node.id);
}

/** Stage one exact paragraph/heading gap; only this step's fresh sites enter its receipt. */
export function stageTrackedSiblings(
  writer: TrackedStageWriter,
  edit: Siblings,
  revision: TrackedStageAttribution,
  listContext: DirectListContext
): readonly NativeSiblingProposal[] {
  const before = currentStagePart(writer, edit.partName);
  if (edit.blocks?.length)
    refuseWrite(
      'Tracked application fragments need their own staged executor.',
      'unsupported-operation'
    );
  const ids = gap(before, edit);
  if (
    !ids ||
    ids.length !== edit.removedIds.length ||
    ids.some((id, index) => id !== edit.removedIds[index])
  )
    refuseWrite('The original native sibling gap no longer matches the plan.', 'ambiguous-target');
  const protection = propertyControlRefusal(before, edit.parentId);
  if (protection) refuseWrite(protection, 'unsupported-operation');
  const originalSites = new Map(
    collectRevisionSites(before).map((site) => [site.node.id, JSON.stringify(site.node)])
  );
  const pendingParagraphs = new Set(
    collectRevisionSites(before).map((site) => paragraphOwner(before, site.node))
  );
  for (const id of edit.removedIds) {
    const paragraph = findNode(before, id);
    if (paragraph?.kind !== 'paragraph' || section(paragraph))
      refuseWrite(
        'Only complete native paragraphs without section marks can be removed here.',
        'unsupported-operation'
      );
    if (
      propertyControlRefusal(before, id) ||
      contentControlsIn(paragraph).some((entry) => propertyControlRefusal(before, entry.node.id))
    )
      refuseWrite(
        'Protected paragraph content cannot enter this proposal.',
        'unsupported-operation'
      );
    if (pendingParagraphs.has(id))
      refuseWrite(
        'Deleting an earlier pending paragraph needs explicit ownership reconciliation.',
        'unsupported-operation'
      );
  }
  const seeds = newParagraphSeeds(edit.inserted, { code: true });
  if (!seeds.success) refuseWrite(seeds.error, seeds.code);
  const oldIds = nodeIds(before);
  for (const id of edit.removedIds) {
    const text = paragraphTextOf(before, id)!;
    if (text)
      writer.apply(before.name, {
        op: 'deleteText',
        paragraphId: id,
        start: 0,
        end: text.length,
        revision,
      });
    writer.apply(before.name, {
      op: 'setParagraphMarkRevision',
      paragraphId: id,
      kind: 'del',
      revision,
    });
  }
  if (seeds.paragraphs.length)
    writer.apply(before.name, {
      op: 'insertParagraphs',
      parentId: edit.parentId,
      beforeId: edit.afterId,
      paragraphs: seeds.paragraphs,
      revision,
    });
  let current = currentStagePart(writer, before.name);
  const allocated = gap(current, edit);
  if (!allocated) refuseWrite('The native sibling gap disappeared.', 'apply-failed');
  const inserted = allocated.filter(
    (id) => !oldIds.has(id) && findNode(current, id)?.kind === 'paragraph'
  );
  if (inserted.length !== seeds.paragraphs.length)
    refuseWrite('Native sibling allocation did not match the planned gap.', 'apply-failed');
  stageParagraphResources(writer, before.name, inserted, seeds);
  const lists = newCellListAssignments(inserted, edit.lists);
  if (lists.length)
    // The paragraph insertion owns the whole new list item. Its properties need no
    // separate formatting proposal; rejection removes the complete born paragraph.
    writer.edit((pkg) => applyDirectListAssignments(pkg, before.name, lists, listContext));
  if (lists.some((assignment) => assignment.target.kind === 'checklist'))
    writer.edit((pkg) => withBornChecklistOwnership(pkg, before.name, inserted));
  for (const binding of edit.styles ?? []) {
    if (binding.path.length !== 1 || !inserted[binding.path[0]!])
      refuseWrite('A native sibling style has no allocated paragraph.', 'ambiguous-target');
    const paragraphId = inserted[binding.path[0]!]!;
    const properties = directParagraphProperties(
      currentStagePart(writer, before.name),
      paragraphId
    ).filter((property) => property.localName !== 'pStyle');
    writer.apply(before.name, {
      op: 'setParagraphProperties',
      paragraphId,
      properties: [...properties, { localName: 'pStyle', attributes: { val: binding.styleId } }],
    });
  }
  current = currentStagePart(writer, before.name);
  const owners = new Set([...edit.removedIds, ...inserted]);
  const sites = collectRevisionSites(current).filter((site) => !originalSites.has(site.node.id));
  const attributed = sites.map((site) => {
    if (
      !owners.has(paragraphOwner(current, site.node) ?? '') ||
      !['ins', 'del'].includes(site.node.localName)
    )
      refuseWrite('A new native revision escaped its structural gap.', 'apply-failed');
    return { nodeId: site.node.id, renumber: `sibling:${site.node.localName}` };
  });
  if (attributed.length)
    writer.apply(before.name, { op: 'setRevisionAttribution', revision, sites: attributed });
  current = currentStagePart(writer, before.name);
  for (const [id, old] of originalSites)
    if (JSON.stringify(findNode(current, id)) !== old)
      refuseWrite('A sibling proposal changed an unrelated pending revision.', 'ambiguous-target');
  const addresses = new Map<string, RevisionAddress>();
  for (const site of collectRevisionSites(current)) {
    if (originalSites.has(site.node.id)) continue;
    const address = nativePropertyAddress(site.node);
    if (!owners.has(paragraphOwner(current, site.node) ?? '') || !address || site.refused)
      refuseWrite('Missing native sibling review receipt.', 'apply-failed');
    addresses.set(propertyAddressKey(before.name, address), Object.freeze({ ...address }));
  }
  if (addresses.size !== Number(Boolean(inserted.length)) + Number(Boolean(edit.removedIds.length)))
    refuseWrite('A structural gap lost its distinct insertion/deletion addresses.', 'apply-failed');
  if (!addresses.size) return [];
  const receipt = [...addresses.values()];
  return [
    {
      identity: JSON.stringify([
        before.name,
        receipt.map((address) => propertyAddressKey(before.name, address)).sort(),
      ]),
      partName: before.name,
      addresses: receipt,
      insertedParagraphIds: inserted,
      removedParagraphIds: [...edit.removedIds],
    },
  ];
}
