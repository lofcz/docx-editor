import { trackedMergeParagraphs } from './tracked-merge-paragraphs.ts';
// Tracked table cell MERGE and SPLIT, in Word's cell revision markup.
//
// A split is decorated like a column insertion: new rows and cells become insertions and every
// surviving cell records its old properties, so rejecting restores the original table.
//
// A merge keeps the original cells pending. Cells right of each row's first cell carry `w:cellDel`,
// whose acceptance gives their span and width to that first cell. In a multi-row merge the top
// first cell carries `w:cellMerge` with `w:vMerge="rest"` and the lower ones `w:vMerge="cont"`,
// whose acceptance discards their content. The content the merge gathers into the owner arrives as
// a tracked paragraph insertion, and an empty owner paragraph the merge drops has its mark deleted.
// Accepting the set lands the direct merge; rejecting it restores the original cells and content.
import { runWithoutJournalCapture } from '../package/canonical-primitive-capture.ts';
import { findNode, replaceNode, type EditOptions } from '../package/ooxml-edit.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { readMergeableTable } from './tree-op-table-cell-shape.ts';
import { wmlChildNamed } from './tree-op-table-shared.ts';
import { fromEdit } from './tree-op-nodes.ts';
import { invalidRevisionAttribution } from './tree-op-revision-attribution.ts';
import { freshWmlElement } from './table-property-attributes.ts';
import {
  attributes,
  trackedTableColumnRefusal,
  hasPendingTableRevision,
  insertionMarkup,
  newStamp,
  withCellMarker,
  type Stamp,
} from './tree-op-tracked-table-columns.ts';
import { tableWithDeclaredWmlBinding } from './tree-op-tables.ts';
import type { RevisionAttributionInput, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';
import type { TreeDocOp } from './tree-ops.ts';

type MergeOp = Extract<TreeDocOp, { op: 'mergeTableCells' }>;
type SplitOp = Extract<TreeDocOp, { op: 'splitTableCell' }>;
type Tracked<T> = T & { readonly revision: RevisionAttributionInput };
type Apply = (part: OoxmlPart, op: TreeDocOp) => TreeOpResult;

/** Refusals specific to the tracked form; the untracked validator runs first. */
export function trackedTableMergeRefusal(
  part: OoxmlPart,
  op: Tracked<MergeOp | SplitOp>
): TreeOpRejection | null {
  if (invalidRevisionAttribution(op.revision)) return 'invalidArgs';
  const table = findNode(part, op.tableId);
  if (table?.kind !== 'table') return 'unknown-table';
  return hasPendingTableRevision(table) ? 'unsupported' : null;
}

const children = (node: OoxmlNode | null | undefined): readonly OoxmlNode[] =>
  node && node.kind !== 'textValue' ? node.children : [];

export function applyTrackedSplitTableCell(
  part: OoxmlPart,
  op: Tracked<SplitOp>,
  untracked: Apply,
  options?: EditOptions
): TreeOpResult {
  const refusal = trackedTableMergeRefusal(part, op);
  if (refusal) return { ok: false, reason: refusal };
  const { revision, ...plain } = op;
  const before = findNode(part, op.tableId) as OoxmlElement;
  const split = untracked(part, plain);
  if (!split.ok) return split;
  const stamp = newStamp(part, before, revision, options?.trackedRevisionIds?.mint);
  const after = findNode(split.part, op.tableId) as OoxmlElement;
  const marked = tableWithDeclaredWmlBinding(
    insertionMarkup(before, after, stamp),
    split.part,
    stamp.wml
  );
  return fromEdit(replaceNode(split.part, op.tableId, marked, options), {
    ...split.effect,
    dirty: [...new Set([...split.effect.dirty, op.tableId])],
  });
}

function cellMergeMarker(stamp: Stamp, cell: OoxmlElement, next: 'rest' | 'cont'): OoxmlElement {
  const vMerge = wmlChildNamed(wmlChildNamed(cell, 'tcPr') ?? cell, 'vMerge');
  const prior = vMerge
    ? vMerge.attributes.find((a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === 'val')
        ?.value === 'restart'
      ? 'rest'
      : 'cont'
    : undefined;
  const attribute = (localName: string, value: string) => ({
    kind: 'genericExtension' as const,
    namespaceUri: WML_NAMESPACE_URI,
    localName,
    prefix: stamp.wml.attributePrefix,
    value,
  });
  return freshWmlElement('cellMerge', stamp.nextId, stamp.wml, [
    ...attributes(stamp),
    attribute('vMerge', next),
    ...(prior === undefined ? [] : [attribute('vMergeOrig', prior)]),
  ]);
}

export function applyTrackedMergeTableCells(
  part: OoxmlPart,
  op: Tracked<MergeOp>,
  untracked: Apply,
  apply: Apply,
  options?: EditOptions
): TreeOpResult {
  const refusal = trackedTableMergeRefusal(part, op);
  if (refusal) return { ok: false, reason: refusal };
  const { revision, ...plain } = op;
  // A scratch direct merge shows where content lands; it never enters the shared journal.
  const direct = runWithoutJournalCapture(() => untracked(part, plain));
  if (!direct.ok) return direct;
  const read = readMergeableTable(part, op.tableId);
  if (!read.ok) return { ok: false, reason: read.reason };
  const selected = read.table.index.placed.filter((cell) => op.cellIds.includes(cell.cell.id));
  const top = Math.min(...selected.map((cell) => cell.rowIndex));
  const bottom = Math.max(...selected.map((cell) => cell.rowIndex));
  const left = Math.min(...selected.map((cell) => cell.startCol));
  const right = Math.max(...selected.map((cell) => cell.endCol));
  const owners = new Map<string, 'rest' | 'cont' | null>();
  const removed = new Set<string>();
  for (let row = top; row <= bottom; row += 1)
    for (const cell of read.table.index.byRow.get(row) ?? []) {
      if (cell.startCol < left || cell.endCol > right) continue;
      if (cell.startCol === left)
        owners.set(cell.cell.id, top === bottom ? null : row === top ? 'rest' : 'cont');
      else removed.add(cell.cell.id);
    }
  const ownerId = [...owners.keys()][0]!;
  const ownerBefore = findNode(part, ownerId) as OoxmlElement;
  const ownerAfter = findNode(direct.part, ownerId);
  if (!ownerAfter || ownerAfter.kind === 'textValue')
    return { ok: false, reason: 'tree-invariant' };
  const kept = new Set(children(ownerBefore).map((node) => node.id));
  const moved = children(ownerAfter).filter(
    (node) =>
      (node.kind === 'paragraph' || node.kind === 'table' || node.kind === 'contentControl') &&
      !kept.has(node.id)
  );
  const dropped = children(ownerBefore).filter(
    (node) =>
      node.kind === 'paragraph' && !children(ownerAfter).some((entry) => entry.id === node.id)
  );
  // Only plain paragraphs can travel as a tracked insertion; dropped owner paragraphs are empty.
  if (moved.some((node) => node.kind !== 'paragraph')) return { ok: false, reason: 'unsupported' };
  if (dropped.some((node) => children(node).some((child) => child.kind !== 'paragraphProperties')))
    return { ok: false, reason: 'unsupported' };

  const projectedMoved = trackedMergeParagraphs(part, moved, options?.trackedRevisionIds);
  if (!projectedMoved) return { ok: false, reason: 'unsupported' };
  const before = findNode(part, op.tableId) as OoxmlElement;
  const stamp = newStamp(part, before, revision, options?.trackedRevisionIds?.mint);
  const marked = Object.freeze({
    ...before,
    children: before.children.map((row): OoxmlNode => {
      if (row.kind !== 'tableRow') return row;
      return Object.freeze({
        ...row,
        children: row.children.map((cell): OoxmlNode => {
          if (cell.kind !== 'tableCell') return cell;
          if (removed.has(cell.id))
            return withCellMarker(
              cell,
              freshWmlElement('cellDel', stamp.nextId, stamp.wml, attributes(stamp)),
              stamp
            );
          const merge = owners.get(cell.id);
          return merge ? withCellMarker(cell, cellMergeMarker(stamp, cell, merge), stamp) : cell;
        }),
      }) as OoxmlElement;
    }),
  }) as OoxmlElement;
  const replaced = replaceNode(
    part,
    op.tableId,
    tableWithDeclaredWmlBinding(marked, part, stamp.wml),
    options
  );
  if (!replaced.ok)
    return { ok: false, reason: 'tree-invariant', detail: JSON.stringify(replaced.issues) };
  let current = replaced.part;
  const step = (next: TreeDocOp): TreeOpRejection | null => {
    const result = apply(current, next);
    if (!result.ok) return result.reason;
    current = result.part;
    return null;
  };
  const tracked = {
    author: revision.author,
    ...(revision.date === undefined ? {} : { date: revision.date }),
  };
  if (projectedMoved.length) {
    const failed = step({
      op: 'insertParagraphs',
      parentId: ownerId,
      beforeId: null,
      paragraphs: projectedMoved,
      revision: tracked,
    });
    if (failed)
      return { ok: false, reason: failed === 'invalidArgs' ? 'invalidArgs' : 'unsupported' };
  }
  for (const paragraph of dropped) {
    const failed = step({
      op: 'setParagraphMarkRevision',
      paragraphId: paragraph.id,
      kind: 'del',
      revision: tracked,
    });
    if (failed) return { ok: false, reason: 'unsupported' };
  }
  return {
    ok: true,
    part: current,
    effect: {
      ...direct.effect,
      created: [],
      deleted: [],
      dirty: [op.tableId, ...owners.keys(), ...removed],
    },
  };
}

/** Validation of the tracked table structure forms, after the untracked validator passes. */
export function trackedStructureRefusal(
  part: OoxmlPart,
  op: Extract<
    TreeDocOp,
    { op: 'insertTableColumn' | 'deleteTableColumn' | 'mergeTableCells' | 'splitTableCell' }
  >
): TreeOpRejection | null {
  if (!op.revision) return 'widthTwips' in op && op.widthTwips !== undefined ? 'invalidArgs' : null;
  return op.op === 'insertTableColumn' || op.op === 'deleteTableColumn'
    ? trackedTableColumnRefusal(part, { ...op, revision: op.revision })
    : trackedTableMergeRefusal(part, { ...op, revision: op.revision });
}
