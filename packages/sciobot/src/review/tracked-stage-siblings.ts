import type { InlineImageRuntime } from '../native/inline-image-runtime';
import { stageBornAtoms } from './tracked-born-atoms';
import {
  collectRevisionSites,
  contentControlsIn,
  directParagraphProperties,
  findNode,
  paragraphTextOf,
  parentNodeOf,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
  type RevisionAddress,
} from '@docx-editor.dev/core/store';
import { applyDirectListAssignments, type DirectListContext } from '../ai/direct-lists';
import { newCellListAssignments } from '../diff/new-cell-list-assignments';
import { nodeIds, refuseWrite, siblingFragmentSeeds } from '../diff/fragment-seeds';
import type { CanonicalEdit } from '../diff/plan';
import type { NativeSiblingProposal } from './track-blocks';
import { withBornChecklistOwnership } from './tracked-born-checklists';
import { stageSiblingTables } from './tracked-sibling-tables';
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
  listContext: DirectListContext,
  images?: InlineImageRuntime
): readonly NativeSiblingProposal[] {
  const before = currentStagePart(writer, edit.partName);
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
  const seeded = siblingFragmentSeeds(
    edit,
    !!edit.afterId && findNode(before, edit.afterId)?.kind === 'paragraph',
    images
  );
  if (!seeded.success) refuseWrite(seeded.error, seeded.code);
  const seeds = seeded.seeds;
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
  const listBindings = (edit.lists ?? [])
    .filter((binding) => binding.path.length === 1)
    .map((binding) => {
      const index = binding.path.length === 1 ? seeds.paragraphOf.get(binding.path[0]!) : undefined;
      if (index === undefined)
        refuseWrite('A native list binding has no new paragraph.', 'ambiguous-target');
      return { ...binding, path: [index] };
    });
  const lists = newCellListAssignments(inserted, listBindings);
  if (lists.length)
    // The paragraph insertion owns the whole new list item. Its properties need no
    // separate formatting proposal; rejection removes the complete born paragraph.
    writer.edit((pkg) => applyDirectListAssignments(pkg, before.name, lists, listContext));
  if (lists.some((assignment) => assignment.target.kind === 'checklist'))
    writer.edit((pkg) => withBornChecklistOwnership(pkg, before.name, inserted));
  for (const [index, styleId] of seeds.styles) {
    const paragraphId = inserted[index];
    if (!paragraphId)
      refuseWrite('A native sibling style has no allocated paragraph.', 'ambiguous-target');
    const properties = directParagraphProperties(
      currentStagePart(writer, before.name),
      paragraphId
    ).filter((property) => property.localName !== 'pStyle');
    writer.apply(before.name, {
      op: 'setParagraphProperties',
      paragraphId,
      properties: [...properties, { localName: 'pStyle', attributes: { val: styleId } }],
    });
  }
  stageSiblingTables(writer, edit, seeds, inserted, revision, listContext, images);
  stageBornAtoms(
    writer,
    before.name,
    inserted,
    (edit.atoms ?? []).filter(
      (atom) =>
        atom.location.kind === 'new-paragraph' &&
        edit.inserted[atom.location.path[0]!]?.type !== 'table'
    ),
    images
  );
  current = currentStagePart(writer, before.name);
  const finalGap = gap(current, edit);
  if (!finalGap) refuseWrite('The completed rich sibling gap disappeared.', 'apply-failed');
  const bornRoots = finalGap.filter((id) => !oldIds.has(id));
  const owners = new Set([...edit.removedIds, ...bornRoots]);
  const belongs = (node: OoxmlNode): boolean => {
    let owner: OoxmlNode | null = node;
    while (owner) {
      if (owners.has(owner.id)) return true;
      owner = parentNodeOf(current, owner.id);
    }
    return false;
  };
  const sites = collectRevisionSites(current).filter((site) => !originalSites.has(site.node.id));
  const attributed = sites.map((site) => {
    if (!belongs(site.node) || !['ins', 'del', 'cellIns'].includes(site.node.localName))
      refuseWrite('A new native revision escaped its structural gap.', 'apply-failed');
    return {
      nodeId: site.node.id,
      renumber: `sibling:${site.node.localName === 'cellIns' ? 'ins' : site.node.localName}`,
    };
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
    if (!belongs(site.node) || !address || site.refused)
      refuseWrite('Missing native sibling review receipt.', 'apply-failed');
    addresses.set(propertyAddressKey(before.name, address), Object.freeze({ ...address }));
  }
  if (
    addresses.size !==
    Number(bornRoots.some((id) => findNode(current, id)?.kind !== 'contentControl')) +
      Number(Boolean(edit.removedIds.length))
  )
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
