// Tracked table COLUMN insertion and deletion, in Word's cell and table-property revision markup.
//
// A column has no element of its own, so Word records its insertion as `w:cellIns` on every new
// cell, `w:tcPrChange` on every existing cell (a row formatting decision restores unrecorded
// cells to defaults, so every survivor must record its old properties), `w:tblGridChange` for
// the old grid, and `w:tblPrChange` when the table width changes. Rejecting that set restores the
// original table exactly. A deletion marks each deleted cell with `w:cellDel`; the grid and the
// other cells stay until the decision. Accepting it follows Word: the neighbour takes the width.
import {
  createMarkContainerIdAllocator,
  createNodeIdAllocator,
  replaceNode,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import { ooxmlTreesEqual } from '../package/ooxml-serialize.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import {
  wmlFreshNamespaceContextAt,
  type WmlFreshNamespaceContext,
} from '../package/wml-namespace.ts';
import { findNode } from '../package/ooxml-edit.ts';
import { readMergeableTable } from './tree-op-table-cell-shape.ts';
import { patchTblPrChild, patchTcPrChild, patchTrPrChild } from './tree-op-table-properties.ts';
import { wmlChildNamed } from './tree-op-table-shared.ts';
import { cloneWithNewIds, fromEdit, TEXT_DEPS } from './tree-op-nodes.ts';
import { invalidRevisionAttribution } from './tree-op-revision-attribution.ts';
import { nextRevisionId } from './tree-op-revision-ids.ts';
import { freshWmlElement } from './table-property-attributes.ts';
import { tableWithDeclaredWmlBinding } from './tree-op-tables.ts';
import { applyTableProperties } from './tree-op-table-authoring.ts';
import { readTwipsMeasure } from '../units.ts';
import type { RevisionAttributionInput, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';
import type { TreeDocOp } from './tree-ops.ts';

type ColumnOp = Extract<TreeDocOp, { op: 'insertTableColumn' | 'deleteTableColumn' }>;
type TrackedColumnOp = ColumnOp & { readonly revision: RevisionAttributionInput };

const PENDING_TABLE_MARKERS = new Set([
  'ins',
  'del',
  'trPrChange',
  'cellIns',
  'cellDel',
  'cellMerge',
  'tcPrChange',
  'tblPrChange',
  'tblPrExChange',
  'tblGridChange',
]);
const isW = (node: OoxmlNode, name: string): node is OoxmlElement =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && node.localName === name;
const elements = (node: OoxmlNode | undefined): readonly OoxmlElement[] =>
  node && node.kind !== 'textValue'
    ? (node.children.filter((child) => child.kind !== 'textValue') as OoxmlElement[])
    : [];

/** Pending structure or property history on THIS table, not on tables nested in its cells. */
export function hasPendingTableRevision(table: OoxmlElement): boolean {
  const holds = (container: OoxmlElement | undefined) =>
    elements(container).some(
      (child) =>
        child.namespaceUri === WML_NAMESPACE_URI && PENDING_TABLE_MARKERS.has(child.localName)
    );
  if (holds(wmlChildNamed(table, 'tblPr')) || holds(wmlChildNamed(table, 'tblGrid'))) return true;
  return elements(table).some(
    (row) =>
      row.kind === 'tableRow' &&
      (holds(wmlChildNamed(row, 'trPr')) ||
        elements(row).some(
          (cell) => cell.kind === 'tableCell' && holds(wmlChildNamed(cell, 'tcPr'))
        ))
  );
}

/** Refusals specific to the tracked form; the untracked validator runs first. */
export function trackedTableColumnRefusal(
  part: OoxmlPart,
  op: TrackedColumnOp
): TreeOpRejection | null {
  if (invalidRevisionAttribution(op.revision)) return 'invalidArgs';
  const width = (op as { widthTwips?: unknown }).widthTwips;
  if (width !== undefined && (op.op !== 'insertTableColumn' || !Number.isInteger(width)))
    return 'invalidArgs';
  const table = findNode(part, op.tableId);
  if (table?.kind !== 'table') return 'unknown-table';
  // One decision set per table: stacked structure histories cannot restore one another.
  if (hasPendingTableRevision(table)) return 'unsupported';
  if (op.op === 'deleteTableColumn') {
    const read = readMergeableTable(part, op.tableId);
    if (!read.ok) return read.reason;
    const column = read.table.topology.gridColumns.findIndex(
      (entry) => entry.id === op.gridColumnId
    );
    if (column < 0) return 'unknown-grid-column';
    // A horizontal span across the column would need a narrowing that `w:cellDel` cannot record.
    for (const cells of read.table.index.byRow.values())
      if (cells.some((cell) => cell.startCol <= column && cell.endCol >= column && cell.span > 1))
        return 'unsupported';
  }
  return null;
}

export interface Stamp {
  readonly wml: WmlFreshNamespaceContext;
  readonly nextId: () => string;
  readonly holderId: () => string;
  readonly revisionId: () => string;
  readonly revision: RevisionAttributionInput;
}
export function attributes(stamp: Stamp, withAuthor = true) {
  const attribute = (localName: string, value: string) => ({
    kind: 'genericExtension' as const,
    namespaceUri: WML_NAMESPACE_URI,
    localName,
    prefix: stamp.wml.attributePrefix,
    value,
  });
  return [
    attribute('id', stamp.revisionId()),
    ...(withAuthor
      ? [
          attribute('author', stamp.revision.author),
          ...(stamp.revision.date === undefined ? [] : [attribute('date', stamp.revision.date)]),
        ]
      : []),
  ];
}
/** `w:<name>Change` holding a copy of the container as it was. */
export function history(
  stamp: Stamp,
  container: string,
  prior: readonly OoxmlNode[],
  withAuthor = true
) {
  const snapshot = freshWmlElement(
    container,
    stamp.nextId,
    stamp.wml,
    [],
    prior.map((node) => cloneWithNewIds(node, stamp.nextId))
  );
  return freshWmlElement(
    `${container}Change`,
    stamp.nextId,
    stamp.wml,
    attributes(stamp, withAuthor),
    [snapshot]
  );
}
export function withCellMarker(
  cell: OoxmlElement,
  marker: OoxmlElement,
  stamp: Stamp
): OoxmlElement {
  const existing = wmlChildNamed(cell, 'tcPr');
  const holder = existing ?? freshWmlElement('tcPr', stamp.holderId, stamp.wml, []);
  const patched = patchTcPrChild(holder, marker);
  if (!patched.ok) throw new Error('cell marker refused');
  const children = existing
    ? cell.children.map((child) => (child === existing ? patched.container : child))
    : [patched.container, ...cell.children];
  return Object.freeze({ ...cell, children }) as OoxmlElement;
}

/** Marker allocation for one tracked table proposal. */
export function newStamp(
  part: OoxmlPart,
  table: OoxmlElement,
  revision: RevisionAttributionInput,
  revisionIds?: () => string
): Stamp {
  return {
    wml: wmlFreshNamespaceContextAt(part, table),
    nextId: createNodeIdAllocator(part),
    holderId: createMarkContainerIdAllocator(part),
    revisionId: revisionIds ?? nextRevisionId(part),
    revision,
  };
}
function withRowMarker(row: OoxmlElement, stamp: Stamp): OoxmlElement {
  const existing = wmlChildNamed(row, 'trPr');
  const holder = existing ?? freshWmlElement('trPr', stamp.holderId, stamp.wml, []);
  const patched = patchTrPrChild(
    holder,
    freshWmlElement('ins', stamp.nextId, stamp.wml, attributes(stamp))
  );
  if (!patched.ok) throw new Error('row marker refused');
  const children = existing
    ? row.children.map((child) => (child === existing ? patched.container : child))
    : [patched.container, ...row.children];
  return Object.freeze({ ...row, children }) as OoxmlElement;
}

/**
 * Decorate a structural edit by comparing the table before and after the untracked command:
 * new rows and cells become insertions, every surviving cell records its old properties, and the
 * grid and table properties record their old values.
 */
export function insertionMarkup(
  before: OoxmlElement,
  after: OoxmlElement,
  stamp: Stamp
): OoxmlElement {
  const oldCells = new Map<string, OoxmlElement>();
  const oldRows = new Set(
    elements(before)
      .filter((row) => row.kind === 'tableRow')
      .map((row) => row.id)
  );
  for (const row of elements(before))
    if (row.kind === 'tableRow')
      for (const cell of elements(row)) if (cell.kind === 'tableCell') oldCells.set(cell.id, cell);
  const oldGrid = wmlChildNamed(before, 'tblGrid');
  const oldTblPr = wmlChildNamed(before, 'tblPr');
  const children = after.children.map((child): OoxmlNode => {
    if (isW(child, 'tblPr') && oldTblPr && !ooxmlTreesEqual(child, oldTblPr)) {
      const patched = patchTblPrChild(child, history(stamp, 'tblPr', oldTblPr.children));
      if (!patched.ok) throw new Error('table property history refused');
      return patched.container;
    }
    if (isW(child, 'tblGrid') && oldGrid && !ooxmlTreesEqual(child, oldGrid))
      // CT_TblGridChange is CT_Markup: an id and no author.
      return Object.freeze({
        ...child,
        children: [...child.children, history(stamp, 'tblGrid', oldGrid.children, false)],
      }) as OoxmlElement;
    if (child.kind !== 'tableRow') return child;
    const freshRow = !oldRows.has(child.id);
    const cells = child.children.map((cell): OoxmlNode => {
      if (cell.kind !== 'tableCell') return cell;
      const old = oldCells.get(cell.id);
      if (!old)
        return withCellMarker(
          cell,
          freshWmlElement('cellIns', stamp.nextId, stamp.wml, attributes(stamp)),
          stamp
        );
      return withCellMarker(
        cell,
        history(stamp, 'tcPr', wmlChildNamed(old, 'tcPr')?.children ?? []),
        stamp
      );
    });
    const row = Object.freeze({ ...child, children: cells }) as OoxmlElement;
    return freshRow ? withRowMarker(row, stamp) : row;
  });
  return Object.freeze({ ...after, children }) as OoxmlElement;
}

const gridOf = (part: OoxmlPart, tableId: string) =>
  elements(wmlChildNamed(findNode(part, tableId) as OoxmlElement, 'tblGrid')).filter((node) =>
    isW(node, 'gridCol')
  );
/** Native fixed-grid sizing of the column the insertion created; other widths stay as authored. */
function sizeNewColumn(
  before: OoxmlPart,
  after: OoxmlPart,
  tableId: string,
  widthTwips: number,
  options?: EditOptions
): TreeOpResult {
  const old = new Set(gridOf(before, tableId).map((column) => column.id));
  const widths: number[] = [];
  let created = 0;
  for (const column of gridOf(after, tableId)) {
    const fresh = !old.has(column.id);
    created += Number(fresh);
    const width = fresh
      ? widthTwips
      : readTwipsMeasure(
          column.attributes.find(
            (attribute) =>
              attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'w'
          )?.value
        );
    if (width === null) return { ok: false, reason: 'unsupported' };
    widths.push(width);
  }
  if (created !== 1) return { ok: false, reason: 'unsupported' };
  return applyTableProperties(
    after,
    { op: 'setTableProperties', tableId, columnWidthsTwips: widths },
    options
  );
}

/**
 * Apply a tracked column insertion or deletion. `untracked` is the native column applier, so the
 * tracked insertion lands exactly the cells, spans, and widths the direct command creates.
 */
export function applyTrackedTableColumnOp(
  part: OoxmlPart,
  op: TrackedColumnOp,
  untracked: (part: OoxmlPart, op: ColumnOp) => TreeOpResult,
  options?: EditOptions
): TreeOpResult {
  const refusal = trackedTableColumnRefusal(part, op);
  if (refusal) return { ok: false, reason: refusal };
  const { revision, ...plain } = op;
  const before = findNode(part, op.tableId) as OoxmlElement;
  const stamp = newStamp(part, before, revision, options?.trackedRevisionIds?.mint);
  const wml = stamp.wml;
  if (op.op === 'insertTableColumn') {
    const { widthTwips, ...column } = plain as ColumnOp & { widthTwips?: number };
    let inserted = untracked(part, column as ColumnOp);
    if (!inserted.ok) return inserted;
    if (widthTwips !== undefined) {
      const sized = sizeNewColumn(part, inserted.part, op.tableId, widthTwips, options);
      if (!sized.ok) return sized;
      inserted = { ...sized, effect: inserted.effect };
    }
    const after = findNode(inserted.part, op.tableId) as OoxmlElement;
    const marked = tableWithDeclaredWmlBinding(
      insertionMarkup(before, after, stamp),
      inserted.part,
      wml
    );
    return fromEdit(replaceNode(inserted.part, op.tableId, marked, options), {
      ...inserted.effect,
      dirty: [...new Set([...inserted.effect.dirty, op.tableId])],
    });
  }
  const read = readMergeableTable(part, op.tableId);
  if (!read.ok) return { ok: false, reason: read.reason };
  const column = read.table.topology.gridColumns.findIndex((entry) => entry.id === op.gridColumnId);
  const deleted = new Set<string>();
  for (const cells of read.table.index.byRow.values())
    for (const cell of cells) if (cell.startCol === column) deleted.add(cell.cell.id);
  const marked = Object.freeze({
    ...before,
    children: before.children.map((row): OoxmlNode => {
      if (row.kind !== 'tableRow') return row;
      return Object.freeze({
        ...row,
        children: row.children.map((cell) =>
          cell.kind === 'tableCell' && deleted.has(cell.id)
            ? withCellMarker(
                cell,
                freshWmlElement('cellDel', stamp.nextId, stamp.wml, attributes(stamp)),
                stamp
              )
            : cell
        ),
      }) as OoxmlElement;
    }),
  }) as OoxmlElement;
  return fromEdit(
    replaceNode(part, op.tableId, tableWithDeclaredWmlBinding(marked, part, wml), options),
    {
      dirty: [op.tableId, ...deleted],
      created: [],
      deleted: [],
      dependencyKeys: TEXT_DEPS,
      impact: 'flow-structural',
    }
  );
}
