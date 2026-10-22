import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  collectRevisionSites,
  findNode,
  parentNodeOf,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
  type RevisionAddress,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import type { CanonicalEditPlan } from '../diff/plan';
import {
  nativePropertyAddress,
  propertyAddressKey,
  propertyControlRefusal,
} from './property-fallback';
import type { TrackTextContext } from './tracked-inline';
import {
  commitTrackedStructure,
  fail,
  preflightTrackedStructure,
  readTrackedStructureInput,
} from './tracked-structure';

/** Word cell, merge, property, and paragraph addresses survive save/reopen. */
export interface NativeMergeProposal {
  readonly identity: string;
  readonly partName: string;
  readonly kind: 'cellMerge' | 'cellSplit';
  readonly tableId: string;
  readonly addresses: readonly RevisionAddress[];
}
export type TrackTableMergesResult =
  | {
      readonly success: true;
      readonly revision: RevisionToken;
      readonly proposals: readonly NativeMergeProposal[];
    }
  | NativeEditFailure;

type MergeOp = Extract<TreeDocOp, { op: 'mergeTableCells' | 'splitTableCell' }>;

/** The nearest table, counting a tracked paragraph inside its cell. */
const ownerTable = (part: OoxmlPart, node: OoxmlNode): string | null => {
  let current: OoxmlNode | null = node;
  for (let depth = 0; current && depth < 12; depth++) {
    if (current.kind === 'table') return current.id;
    current = parentNodeOf(part, current.id);
  }
  return null;
};
const address = (node: OoxmlNode): RevisionAddress | null => {
  if (node.kind === 'textValue') return null;
  if (node.localName !== 'tblGridChange') return nativePropertyAddress(node);
  const id = node.attributes.find(
    (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'id'
  )?.value;
  return id === undefined ? null : { id, author: '' };
};

/**
 * Propose D10 merges and splits as Word cell revisions in one gated, separately undoable
 * transaction. Accepting lands the direct command; rejecting restores the original cells.
 */
export function trackTableMergePlan(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  context: TrackTextContext
): TrackTableMergesResult {
  const read = readTrackedStructureInput(editor, input, context, 'merge');
  if (!read.success) return read;
  const { plan, part, revision } = read.input;
  const steps: MergeOp[] = [];
  for (const edit of plan.operations) {
    if (edit.partName !== part.name)
      return fail('Tracked merges must belong to the body part.', 'unsupported-operation');
    if (
      edit.kind !== 'table' ||
      (edit.operation.op !== 'mergeTableCells' && edit.operation.op !== 'splitTableCell')
    )
      return fail(
        'This executor accepts table merge and split plans only.',
        'unsupported-operation'
      );
    const operation = edit.operation;
    if (findNode(part, operation.tableId)?.kind !== 'table')
      return fail('A planned merge names no native table.', 'ambiguous-target');
    if (steps.some((step) => step.tableId === operation.tableId))
      return fail('Plan one tracked merge or split per table.', 'unsupported-operation');
    const protection = propertyControlRefusal(part, operation.tableId);
    if (protection) return fail(protection, 'unsupported-operation');
    steps.push({ ...operation, revision });
  }
  if (!steps.length)
    return { success: true, revision: Object.freeze(read.input.actual), proposals: [] };
  const originalSites = new Map(
    collectRevisionSites(part).map((site) => [site.node.id, JSON.stringify(site.node)])
  );
  let proposals: readonly NativeMergeProposal[] = [];
  const complete = (pkg: OoxmlPackage): OoxmlPackage => {
    const current = pkg.parts.get(part.name)!;
    const addresses = steps.map(() => new Map<string, RevisionAddress>());
    for (const site of collectRevisionSites(current)) {
      const before = originalSites.get(site.node.id);
      if (before !== undefined) {
        if (before !== JSON.stringify(site.node))
          throw new Error('A merge proposal changed an unrelated pending revision.');
        continue;
      }
      const index = steps.findIndex((step) => step.tableId === ownerTable(current, site.node)),
        found = address(site.node);
      if (index < 0 || !found) throw new Error('A new native revision escaped its planned table.');
      addresses[index]!.set(propertyAddressKey(part.name, found), Object.freeze({ ...found }));
    }
    proposals = Object.freeze(
      steps.map((step, index) => {
        const group = [...addresses[index]!.values()];
        if (!group.length) throw new Error('A planned merge did not receive a native revision.');
        return Object.freeze({
          identity: JSON.stringify([
            part.name,
            group.map((entry) => propertyAddressKey(part.name, entry)).sort(),
          ]),
          partName: part.name,
          kind: step.op === 'mergeTableCells' ? 'cellMerge' : 'cellSplit',
          tableId: step.tableId,
          addresses: Object.freeze(group),
        });
      })
    );
    return pkg;
  };
  const refused = preflightTrackedStructure(read.input, steps, complete, 'merge', () => proposals);
  if (refused) return refused;
  proposals = [];
  const committed = commitTrackedStructure(read.input, steps, complete, 'merge', () => proposals);
  if (!committed.success) return committed;
  return { success: true, revision: committed.revision, proposals };
}
