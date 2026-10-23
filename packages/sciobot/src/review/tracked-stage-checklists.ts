import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  contentControlPropertiesOf,
  createNodeIdAllocator,
  directParagraphProperties,
  findNode,
  parentNodeOf,
  replaceChildren,
  withPart,
  WML_NAMESPACE_URI,
  TreePackageStore,
  type OoxmlElement,
  type OoxmlNode,
  type RevisionAddress,
} from '@docx-editor.dev/core/store';
import {
  applyDirectListAssignments,
  createDirectListContext,
  type DirectListContext,
} from '../ai/direct-lists';
import { planListMembership } from '../diff/lists';
import type { CanonicalEdit, CanonicalEditPlan } from '../diff/plan';
import { refuseWrite } from '../diff/fragment-seeds';
import type { ReviewGroupReceipt } from './groups';
import {
  nativePropertyAddress,
  propertyControlRefusal,
  propertyAddressKey,
} from './property-fallback';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

type Edit = Extract<CanonicalEdit, { kind: 'list' }>;
/** Native SDT snapshots retain old checkbox properties beside the proposed marker. */
export function stageTrackedChecklist(
  writer: TrackedStageWriter,
  plan: CanonicalEditPlan,
  edit: Edit,
  revision: TrackedStageAttribution,
  context: DirectListContext = createDirectListContext()
): readonly ReviewGroupReceipt[] {
  const pkg = writer.currentPackage(),
    part = pkg.parts.get(edit.partName)!;
  if (edit.after.kind !== 'checklist' && !edit.before.checkbox)
    refuseWrite('This tracked list executor needs a checklist intention.', 'unsupported-operation');
  const protection = propertyControlRefusal(part, edit.paragraphId);
  if (protection) refuseWrite(protection, 'unsupported-operation');
  const checked = planListMembership({ pkg, partName: edit.partName, revision: plan.revision }, [
    { paragraphId: edit.paragraphId, target: edit.after },
  ]);
  if (!checked.success) refuseWrite(checked.error, checked.code);
  const current = checked.plan.operations[0];
  if (
    !current ||
    current.kind !== 'list' ||
    JSON.stringify(current.before) !== JSON.stringify(edit.before)
  )
    refuseWrite('The checklist source changed after planning.', 'stale-revision');
  const originalSites = new Set(collectRevisionSites(part).map((site) => site.node.id));
  if (edit.after.kind === 'ordered' || edit.after.kind === 'bullet')
    writer.edit((working) =>
      applyDirectListAssignments(
        working,
        edit.partName,
        [{ paragraphId: edit.paragraphId, before: null, target: edit.after }],
        context,
        revision
      )
    );
  // Detached transactions isolate preview primitives from the observed live journal.
  const previewStore = new TreePackageStore(
    writer.currentPackage(),
    writer.currentPackage().parts.get(pkg.mainDocumentPart)!
  );
  const previewResult = previewStore.transact(
    { kind: 'body' },
    (ctx) => {
      ctx.applyPackage((value) =>
        applyDirectListAssignments(
          value,
          edit.partName,
          [{ paragraphId: edit.paragraphId, before: edit.before, target: edit.after }],
          { instances: new Map(context.instances), levels: new Set(context.levels) }
        )
      );
    },
    { recordsHistory: false }
  );
  if (!previewResult.ok)
    refuseWrite('The direct checklist preview refused.', 'unsupported-operation');
  const preview = previewStore.currentPackage();
  const proposedPart = preview.parts.get(edit.partName)!;
  const old = edit.before.checkbox ? findNode(part, edit.before.checkbox.controlId) : null;
  if (old) {
    const controlProtection = propertyControlRefusal(part, old.id);
    if (controlProtection) refuseWrite(controlProtection, 'unsupported-operation');
  }
  const paragraph = findNode(proposedPart, edit.paragraphId)!;
  if (paragraph.kind !== 'paragraph')
    refuseWrite('The checklist paragraph disappeared.', 'stale-revision');
  const next = old
    ? findNode(proposedPart, old.id)
    : paragraph.children.find(
        (child) =>
          child.kind === 'contentControl' &&
          contentControlPropertiesOf(child).tag === 'scio:list-checkbox'
      );
  const changed =
    !old || !next || canonicalOoxmlFingerprint(old) !== canonicalOoxmlFingerprint(next);
  const currentProperties = directParagraphProperties(part, edit.paragraphId);
  const proposedProperties = directParagraphProperties(proposedPart, edit.paragraphId);
  if (JSON.stringify(currentProperties) !== JSON.stringify(proposedProperties)) {
    if (
      edit.after.kind !== 'ordered' &&
      edit.after.kind !== 'bullet' &&
      proposedProperties.some((property) => property.localName === 'numPr')
    )
      writer.apply(edit.partName, {
        op: 'setListNumbering',
        paragraphId: edit.paragraphId,
        numId: '0',
        revision,
      });
    writer.apply(edit.partName, {
      op: 'setParagraphProperties',
      paragraphId: edit.paragraphId,
      properties: proposedProperties,
      revision,
    });
  }
  if (changed)
    writer.edit((working) => {
      const currentPart = working.parts.get(edit.partName)!;
      const source = old && findNode(currentPart, old.id);
      const parent = source
        ? parentNodeOf(currentPart, source.id)
        : findNode(currentPart, edit.paragraphId);
      if (!parent || parent.kind !== 'paragraph')
        refuseWrite(
          'The checklist marker needs its direct paragraph owner.',
          'unsupported-operation'
        );
      const mint = createNodeIdAllocator(currentPart);
      const copy = (node: OoxmlNode, deleted: boolean, preserveIds = false): OoxmlNode =>
        node.kind === 'textValue'
          ? { ...node, id: preserveIds ? node.id : mint() }
          : ({
              ...node,
              id: preserveIds ? node.id : mint(),
              ...(deleted && node.kind === 'text'
                ? { kind: 'deletedText', localName: 'delText' }
                : {}),
              children: node.children.map((child) => copy(child, deleted, preserveIds)),
            } as OoxmlNode);
      const wrap = (node: OoxmlNode, kind: 'revisionInsert' | 'revisionDelete'): OoxmlElement =>
        ({
          id: mint(),
          kind,
          namespaceUri: WML_NAMESPACE_URI,
          localName: kind === 'revisionInsert' ? 'ins' : 'del',
          prefix: 'w',
          namespaceBindings: [],
          attributes: Object.entries({ id: '0', author: revision.author, date: revision.date }).map(
            ([localName, value]) => ({
              kind: 'genericExtension',
              namespaceUri: WML_NAMESPACE_URI,
              localName,
              prefix: 'w',
              value,
            })
          ),
          children: [node],
        }) as OoxmlElement;
      const nodes: OoxmlNode[] = [];
      if (source) nodes.push(wrap(copy(source, true, true), 'revisionDelete'));
      if (next) nodes.push(wrap(copy(next, false), 'revisionInsert'));
      const children = source
        ? parent.children.flatMap((child) => (child.id === source.id ? nodes : [child]))
        : [
            ...parent.children.slice(0, parent.children[0]?.kind === 'paragraphProperties' ? 1 : 0),
            ...nodes,
            ...parent.children.slice(parent.children[0]?.kind === 'paragraphProperties' ? 1 : 0),
          ];
      const result = replaceChildren(currentPart, parent.id, children);
      if (!result.ok)
        refuseWrite('Tracked checklist marker failed native validation.', 'unsupported-operation');
      return withPart(working, result.part);
    });
  const fresh = collectRevisionSites(writer.currentPackage().parts.get(edit.partName)!).filter(
    (site) => !originalSites.has(site.node.id)
  );
  if (!fresh.length) return [];
  writer.apply(edit.partName, {
    op: 'setRevisionAttribution',
    revision,
    sites: fresh.map((site) => ({
      nodeId: site.node.id,
      renumber: `checklist:${site.node.localName}`,
    })),
  });
  const addresses = new Map<string, RevisionAddress>();
  for (const site of collectRevisionSites(writer.currentPackage().parts.get(edit.partName)!)) {
    if (originalSites.has(site.node.id)) continue;
    const address = nativePropertyAddress(site.node);
    if (!address || site.refused)
      refuseWrite('The checklist revision cannot be reviewed.', 'unsupported-operation');
    addresses.set(propertyAddressKey(edit.partName, address), address);
  }
  return [{ partName: edit.partName, addresses: [...addresses.values()] }];
}
