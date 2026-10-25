import { cellPropertiesAt } from '../diff/new-cell-properties';
import { cellAtomBindings } from '../diff/inserted-cell-atoms';
import { createInlineImageRuntime, type InlineImageRuntime } from '../native/inline-image-runtime';
import { createDirectListContext } from '../ai/direct-lists';
import { cellListBindings } from '../diff/inserted-cell-lists';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  collectRevisionSites,
  findNode,
  parentNodeOf,
  withPart,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
  type RevisionAddress,
  type TreeDocOp,
  WML_NAMESPACE_URI,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { insertedColumnIndex, insertedColumnOperation } from '../diff/inserted-column-native';
import type { InsertedTableColumnIntent } from '../diff/inserted-column-types';
import type { CanonicalEditPlan } from '../diff/plan';
import { materializeNewCells, newCellSeeds, type NewCellSeeds } from './inserted-row-content';
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

/** Word cell and table-property addresses survive save/reopen. Cell IDs address this mount. */
export interface NativeColumnProposal {
  readonly identity: string;
  readonly partName: string;
  readonly kind: 'columnInsert' | 'columnDelete';
  readonly tableId: string;
  /** New cells for an insertion; cells marked for deletion otherwise. */
  readonly cellIds: readonly string[];
  readonly addresses: readonly RevisionAddress[];
}
export type TrackTableColumnsResult =
  | {
      readonly success: true;
      readonly revision: RevisionToken;
      readonly proposals: readonly NativeColumnProposal[];
    }
  | NativeEditFailure;

interface ColumnStep {
  readonly op: Extract<TreeDocOp, { op: 'insertTableColumn' | 'deleteTableColumn' }>;
  readonly tableId: string;
  readonly kind: 'columnInsert' | 'columnDelete';
  readonly intent?: InsertedTableColumnIntent;
  readonly seeds?: ReadonlyMap<string, NewCellSeeds>;
}

const cellIdsOf = (part: OoxmlPart, tableId: string): string[] => {
  const table = findNode(part, tableId);
  if (!table || table.kind === 'textValue') return [];
  return table.children.flatMap((row) =>
    row.kind === 'tableRow'
      ? row.children.filter((cell) => cell.kind === 'tableCell').map((cell) => cell.id)
      : []
  );
};
const gridAddress = (node: OoxmlNode): RevisionAddress | null => {
  if (node.kind === 'textValue') return null;
  const id = node.attributes.find(
    (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'id'
  )?.value;
  return id === undefined ? null : { id, author: '' };
};
const rowCells = (part: OoxmlPart, rowId: string): string[] => {
  const row = findNode(part, rowId);
  return row && row.kind !== 'textValue'
    ? row.children.filter((cell) => cell.kind === 'tableCell').map((cell) => cell.id)
    : [];
};
/** The nearest table, without crossing into an enclosing table. */
const ownerTable = (part: OoxmlPart, node: OoxmlNode): string | null => {
  let current: OoxmlNode | null = node;
  for (let depth = 0; current && depth < 6; depth++) {
    if (current.kind === 'table') return current.id;
    current = parentNodeOf(part, current.id);
  }
  return null;
};

function columnStep(
  edit: CanonicalEditPlan['operations'][number],
  revision: { author: string; date: string }
): ColumnStep | NativeEditFailure {
  if (edit.kind === 'insert-column') {
    const intent = edit.intent;
    const seeds = new Map<string, NewCellSeeds>();
    for (const [index, entry] of intent.cells.entries()) {
      const payloads = intent.payloads
        .filter((binding) => binding.path[0] === index)
        .map((binding) => ({ ...binding, path: binding.path.slice(1) }));
      const atoms = cellAtomBindings(intent.atoms, index);
      const rich = entry.blocks.some(
        (node) => node.type !== 'paragraph' && node.type !== 'heading'
      );
      if (rich && (intent.lists ?? []).some((binding) => binding.path[0] === index))
        return fail('Lists inside rich new cells need nested fragment list placement.');
      const seeded = newCellSeeds(
        entry.blocks,
        cellListBindings(intent.lists, index),
        cellPropertiesAt(intent.cellProperties, index),
        rich
          ? {
              inserted: entry.blocks,
              blocks: payloads,
              atoms,
            }
          : undefined
      );
      if (!seeded.success) return seeded;
      seeds.set(entry.rowId, seeded.seeds);
    }
    return {
      op: {
        ...insertedColumnOperation(intent),
        revision,
        ...(intent.properties.widthTwips === undefined
          ? {}
          : { widthTwips: intent.properties.widthTwips }),
      },
      tableId: intent.tableId,
      kind: 'columnInsert',
      intent,
      seeds,
    };
  }
  if (edit.kind === 'table') {
    const operation = edit.operation;
    if (operation.op === 'insertTableColumn' || operation.op === 'deleteTableColumn')
      return {
        op: { ...operation, revision },
        tableId: operation.tableId,
        kind: operation.op === 'insertTableColumn' ? 'columnInsert' : 'columnDelete',
      };
  }
  return fail('This executor accepts table column plans only.', 'unsupported-operation');
}

/**
 * Propose native column insertions and deletions as Word cell and table-property revisions in
 * one gated, separately undoable transaction. A new column's content belongs to its cell
 * insertions; explicit widths use the same native sizing as the direct executor.
 */
export function trackTableColumnPlan(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  context: TrackTextContext
): TrackTableColumnsResult {
  const read = readTrackedStructureInput(editor, input, context, 'column');
  if (!read.success) return read;
  const { plan, part, revision } = read.input;
  const steps: ColumnStep[] = [];
  for (const edit of plan.operations) {
    if (edit.partName !== part.name)
      return fail('Tracked columns must belong to the body part.', 'unsupported-operation');
    const step = columnStep(edit, revision);
    if ('success' in step) return step;
    if (findNode(part, step.tableId)?.kind !== 'table')
      return fail('A planned column names no native table.', 'ambiguous-target');
    // One structure decision set per table; a second column needs the first decided.
    if (steps.some((entry) => entry.tableId === step.tableId))
      return fail('Plan one tracked column change per table.', 'unsupported-operation');
    const protection = propertyControlRefusal(part, step.tableId);
    if (protection) return fail(protection, 'unsupported-operation');
    steps.push(step);
  }
  if (!steps.length)
    return { success: true, revision: Object.freeze(read.input.actual), proposals: [] };
  const originalSites = new Map(
    collectRevisionSites(part).map((site) => [site.node.id, JSON.stringify(site.node)])
  );
  let proposals: readonly NativeColumnProposal[] = [];
  const imagePreparation = createInlineImageRuntime(plan, context.images);
  if (!imagePreparation.success) return imagePreparation;
  const imageRuntime: InlineImageRuntime | undefined = imagePreparation.runtime;
  const complete = (pkg: OoxmlPackage): OoxmlPackage => {
    const listContext = createDirectListContext();
    let working = pkg;
    const cells = steps.map((step) => {
      const current = working.parts.get(part.name)!;
      // Deletion receipts name the cells that received `w:cellDel`, after the site scan.
      if (step.kind === 'columnDelete') return [] as string[];
      const intent = step.intent;
      if (!intent) {
        const before = new Set(cellIdsOf(part, step.tableId));
        return cellIdsOf(current, step.tableId).filter((id) => !before.has(id));
      }
      // Core sized the new column inside the tracked insertion, before recording its history.
      if (insertedColumnIndex(part, current, intent) === null)
        throw new Error('Native column allocation is ambiguous.');
      const next = current;
      const created = intent.cells.flatMap((entry) => {
        const old = new Set(rowCells(part, entry.rowId)),
          fresh = rowCells(next, entry.rowId).filter((id) => !old.has(id));
        if (fresh.length > 1 || (!fresh.length && entry.blocks.length))
          throw new Error('The new native column cell is ambiguous.');
        return fresh.length ? [{ cellId: fresh[0]!, seeds: step.seeds!.get(entry.rowId)! }] : [];
      });
      working = materializeNewCells(
        withPart(working, next),
        part.name,
        created,
        listContext,
        imageRuntime,
        [...step.seeds!.values()].some((seed) => seed.fragments) ? revision : undefined
      );
      return created.map((entry) => entry.cellId);
    });
    if (imageRuntime) working = imageRuntime.complete(working);
    const current = working.parts.get(part.name)!;
    const addresses = steps.map(() => new Map<string, RevisionAddress>());
    for (const site of collectRevisionSites(current)) {
      const before = originalSites.get(site.node.id);
      if (before !== undefined) {
        if (before !== JSON.stringify(site.node))
          throw new Error('A column proposal changed an unrelated pending revision.');
        continue;
      }
      const table = ownerTable(current, site.node),
        index = steps.findIndex((step) => step.tableId === table),
        // CT_TblGridChange carries an id and no author; native review reads it that way.
        address =
          site.node.localName === 'tblGridChange'
            ? gridAddress(site.node)
            : nativePropertyAddress(site.node);
      if (index < 0 || !address)
        throw new Error('A new native revision escaped its planned table.');
      if (site.node.localName === 'cellDel') {
        const cell = site.parent ? parentNodeOf(current, site.parent.id) : null;
        if (cell?.kind === 'tableCell') cells[index]!.push(cell.id);
      }
      addresses[index]!.set(propertyAddressKey(part.name, address), Object.freeze({ ...address }));
    }
    proposals = Object.freeze(
      steps.map((step, index) => {
        const group = [...addresses[index]!.values()];
        if (!group.length) throw new Error('A planned column did not receive a native revision.');
        return Object.freeze({
          identity: JSON.stringify([
            part.name,
            group.map((address) => propertyAddressKey(part.name, address)).sort(),
          ]),
          partName: part.name,
          kind: step.kind,
          tableId: step.tableId,
          cellIds: Object.freeze(cells[index]!),
          addresses: Object.freeze(group),
        });
      })
    );
    return working;
  };
  const ops = steps.map((step) => step.op);
  const refused = preflightTrackedStructure(read.input, ops, complete, 'column', () => proposals);
  if (refused) return refused;
  proposals = [];
  const committed = commitTrackedStructure(read.input, ops, complete, 'column', () => proposals);
  if (!committed.success) return committed;
  return { success: true, revision: committed.revision, proposals };
}
