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
import type { PlannedTablePropertyOperation } from '../diff/table-format';
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

/** Word table, row, cell, and grid property-change addresses survive save/reopen. */
export interface NativeTableFormatProposal {
  readonly identity: string;
  readonly partName: string;
  readonly kind: 'tableFormat';
  readonly tableId: string;
  readonly addresses: readonly RevisionAddress[];
}
export type TrackTableFormatResult =
  | {
      readonly success: true;
      readonly revision: RevisionToken;
      readonly proposals: readonly NativeTableFormatProposal[];
    }
  | NativeEditFailure;

const PROPERTY_OPS: ReadonlySet<string> = new Set<PlannedTablePropertyOperation['op']>([
  'setTableProperties',
  'setTableColumnWidths',
  'setTableRightEdgeWidth',
  'setTableRowHeight',
  'setTableCellBorders',
  'setTableCellFill',
  'setTableCellVerticalAlignment',
]);
const RECORDS: ReadonlySet<string> = new Set([
  'tblPrChange',
  'trPrChange',
  'tcPrChange',
  'tblGridChange',
]);

/** The nearest table, without crossing into an enclosing one. */
const ownerTable = (part: OoxmlPart, node: OoxmlNode): string | null => {
  let current: OoxmlNode | null = node;
  for (let depth = 0; current && depth < 8; depth++) {
    if (current.kind === 'table') return current.id;
    current = parentNodeOf(part, current.id);
  }
  return null;
};
const address = (node: OoxmlNode): RevisionAddress | null => {
  if (node.kind === 'textValue') return null;
  if (node.localName !== 'tblGridChange') return nativePropertyAddress(node);
  // CT_TblGridChange carries an id and no author; native review reads it that way.
  const id = node.attributes.find(
    (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'id'
  )?.value;
  return id === undefined ? null : { id, author: '' };
};

/**
 * Propose D11 table and cell property changes as Word property-change records in one gated,
 * separately undoable transaction, independently from cell text. Accepting lands the direct
 * command; rejecting restores widths, fills, borders, alignment, heights, and header flags.
 */
export function trackTableFormatPlan(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  context: TrackTextContext
): TrackTableFormatResult {
  const read = readTrackedStructureInput(editor, input, context, 'table format');
  if (!read.success) return read;
  const { plan, part, revision, surface } = read.input;
  const steps: (PlannedTablePropertyOperation & { readonly revision: typeof revision })[] = [];
  for (const edit of plan.operations) {
    if (edit.partName !== part.name)
      return fail(
        'Tracked table formatting must belong to the body part.',
        'unsupported-operation'
      );
    if (edit.kind !== 'table' || !PROPERTY_OPS.has(edit.operation.op))
      return fail('This executor accepts table property plans only.', 'unsupported-operation');
    const operation = edit.operation as PlannedTablePropertyOperation;
    if (findNode(part, operation.tableId)?.kind !== 'table')
      return fail('A planned table property names no native table.', 'ambiguous-target');
    const protection = propertyControlRefusal(part, operation.tableId);
    if (protection) return fail(protection, 'unsupported-operation');
    steps.push({ ...operation, revision });
  }
  if (!steps.length)
    return { success: true, revision: Object.freeze(read.input.actual), proposals: [] };
  // Property records are formatting revisions; a disabled policy never applies bare changes.
  if (
    surface.session.trackingSettings().doNotTrackFormatting ||
    !editor.snapshot().revisionMarkup.trackFormatting
  )
    return fail(
      'Formatting tracking is disabled for this document or editor.',
      'unsupported-operation'
    );
  const tables = [...new Set(steps.map((step) => step.tableId))];
  // Records inside a planned table may be re-attributed or dropped; every other one stays.
  const outside = new Map(
    collectRevisionSites(part)
      .filter((site) => !tables.includes(ownerTable(part, site.node) ?? ''))
      .map((site) => [site.node.id, JSON.stringify(site.node)])
  );
  const before = new Set(
    collectRevisionSites(part)
      .filter((site) => tables.includes(ownerTable(part, site.node) ?? ''))
      .map((site) => JSON.stringify(site.node))
  );
  let proposals: readonly NativeTableFormatProposal[] = [];
  const complete = (pkg: OoxmlPackage): OoxmlPackage => {
    const current = pkg.parts.get(part.name)!;
    const addresses = tables.map(() => new Map<string, RevisionAddress>());
    const seen = new Set<string>();
    for (const site of collectRevisionSites(current)) {
      const table = ownerTable(current, site.node);
      const index = tables.indexOf(table ?? '');
      if (index < 0) {
        if (outside.get(site.node.id) !== JSON.stringify(site.node))
          throw new Error('A table format proposal changed an unrelated pending revision.');
        seen.add(site.node.id);
        continue;
      }
      if (!RECORDS.has(site.node.localName)) {
        if (!before.has(JSON.stringify(site.node)))
          throw new Error('A table format proposal created a non-property revision.');
        continue;
      }
      if (before.has(JSON.stringify(site.node))) continue;
      const found = address(site.node);
      if (!found) throw new Error('A table property record has no native address.');
      addresses[index]!.set(propertyAddressKey(part.name, found), Object.freeze({ ...found }));
    }
    if ([...outside.keys()].some((id) => !seen.has(id)))
      throw new Error('A table format proposal resolved an unrelated pending revision.');
    proposals = Object.freeze(
      tables.flatMap((tableId, index) => {
        const group = [...addresses[index]!.values()];
        // A command that lands back on the original leaves no record and no proposal.
        if (!group.length) return [];
        return [
          Object.freeze({
            identity: JSON.stringify([
              part.name,
              group.map((entry) => propertyAddressKey(part.name, entry)).sort(),
            ]),
            partName: part.name,
            kind: 'tableFormat' as const,
            tableId,
            addresses: Object.freeze(group),
          }),
        ];
      })
    );
    return pkg;
  };
  const ops: TreeDocOp[] = steps;
  const refused = preflightTrackedStructure(
    read.input,
    ops,
    complete,
    'table format',
    () => proposals
  );
  if (refused) return refused;
  proposals = [];
  const committed = commitTrackedStructure(
    read.input,
    ops,
    complete,
    'table format',
    () => proposals
  );
  if (!committed.success) return committed;
  return { success: true, revision: committed.revision, proposals };
}
