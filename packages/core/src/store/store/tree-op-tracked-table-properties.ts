// Tracked table PROPERTY changes, in Word's property-change markup.
//
// The untracked command runs first, so the tracked form lands exactly the properties the direct
// command writes. The table before and after is then compared container by container, and each
// container the command changed keeps its new children and gains one change wrapper holding its
// old ones: `w:tblPrChange`, `w:trPrChange`, `w:tcPrChange`, and `w:tblGridChange` for the grid.
// Existing review decisions resolve them: accept drops the records, reject restores them.
//
// The tracked-format rules of `tree-op-tracked-properties.ts` apply. The record is the ORIGINAL:
// a container already holding a pending change keeps that snapshot and only re-attributes it. A
// write that lands back on the original leaves no record of this author's; another author's
// stays exactly as written. A changed row records every cell, and a grid change every row, as a
// tracked column does, because rejecting a row formatting decision restores unrecorded cells to
// defaults.
//
// A holder the command creates where the source had none takes a `mark` family id, so a reject
// that empties it removes it again and restores the table exactly.
import { findNode, replaceNode, type EditOptions } from '../package/ooxml-edit.ts';
import { ooxmlTreesEqual } from '../package/ooxml-serialize.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { equivalentNodes } from './ooxml-node-equality.ts';
import { attributeValueOf, fromEdit } from './tree-op-nodes.ts';
import { invalidRevisionAttribution } from './tree-op-revision-attribution.ts';
import { patchTblPrChild, patchTcPrChild, patchTrPrChild } from './tree-op-table-properties.ts';
import { wmlChildNamed } from './tree-op-table-shared.ts';
import { tableWithDeclaredWmlBinding } from './tree-op-tables.ts';
import { recordedProperties } from './tree-op-tracked-properties.ts';
import { history, newStamp, type Stamp } from './tree-op-tracked-table-columns.ts';
import type { RevisionAttributionInput, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';
import type { TreeDocOp } from './tree-ops.ts';

export const TRACKED_TABLE_PROPERTY_OPS = new Set([
  'setTableProperties',
  'setTableColumnWidths',
  'setTableRightEdgeWidth',
  'setTableRowHeight',
  'setTableCellBorders',
  'setTableCellFill',
  'setTableCellVerticalAlignment',
] as const);
export type TablePropertyOp = Extract<
  TreeDocOp,
  { op: typeof TRACKED_TABLE_PROPERTY_OPS extends Set<infer T> ? T : never }
>;
type Tracked<T> = T & { readonly revision: RevisionAttributionInput };
export const isTablePropertyOp = (op: TreeDocOp): op is TablePropertyOp =>
  (TRACKED_TABLE_PROPERTY_OPS as ReadonlySet<string>).has(op.op);
type Container = 'tblPr' | 'trPr' | 'tcPr';

/** Pending structure decisions a property history cannot stack on. */
const STRUCTURE_MARKERS: Readonly<Record<Container | 'tblGrid', ReadonlySet<string>>> = {
  tblPr: new Set(),
  tblGrid: new Set(['tblGridChange']),
  trPr: new Set(['ins', 'del']),
  tcPr: new Set(['cellIns', 'cellDel', 'cellMerge']),
};
/** What a record never holds: its own wrapper, and other pending decisions on the holder. */
const UNRECORDED: Readonly<Record<Container, ReadonlySet<string>>> = {
  tblPr: new Set(['tblPrChange']),
  trPr: new Set(['trPrChange', 'ins', 'del']),
  tcPr: new Set(['tcPrChange', 'cellIns', 'cellDel', 'cellMerge']),
};
const PATCH = { tblPr: patchTblPrChild, trPr: patchTrPrChild, tcPr: patchTcPrChild } as const;

const isW = (node: OoxmlNode, name: string): node is OoxmlElement =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && node.localName === name;
const elements = (node: OoxmlNode | undefined): readonly OoxmlElement[] =>
  node && node.kind !== 'textValue'
    ? (node.children.filter((child) => child.kind !== 'textValue') as OoxmlElement[])
    : [];
const holds = (container: OoxmlElement | undefined, names: ReadonlySet<string>) =>
  elements(container).some(
    (child) => child.namespaceUri === WML_NAMESPACE_URI && names.has(child.localName)
  );

/** Pending row/cell structure or grid history on THIS table, not on nested tables. */
function hasPendingStructure(table: OoxmlElement): boolean {
  return elements(table).some(
    (row) =>
      row.kind === 'tableRow' &&
      (holds(wmlChildNamed(row, 'trPr'), STRUCTURE_MARKERS.trPr) ||
        wmlChildNamed(row, 'tblPrEx') !== undefined ||
        elements(row).some(
          (cell) =>
            cell.kind === 'tableCell' && holds(wmlChildNamed(cell, 'tcPr'), STRUCTURE_MARKERS.tcPr)
        ))
  );
}

/** Refusals specific to the tracked form; the untracked validator runs first. */
export function trackedTablePropertyRefusal(
  part: OoxmlPart,
  op: Tracked<TablePropertyOp>
): TreeOpRejection | null {
  if (invalidRevisionAttribution(op.revision)) return 'invalidArgs';
  const table = findNode(part, op.tableId);
  if (table?.kind !== 'table') return 'unknown-table';
  // A property history over pending structure cannot restore it: one structure set per table.
  return hasPendingStructure(table) ? 'unsupported' : null;
}

const recordable = (children: readonly OoxmlNode[], container: Container) =>
  children.filter(
    (child) =>
      child.kind === 'textValue' ||
      child.namespaceUri !== WML_NAMESPACE_URI ||
      !UNRECORDED[container].has(child.localName)
  );
/** Property sets compared as elements; a schema reorder alone is not a change. */
function sameProperties(left: readonly OoxmlNode[], right: readonly OoxmlNode[]): boolean {
  const byName = (nodes: readonly OoxmlNode[]) =>
    [...nodes].sort((a, b) => {
      const first = a.kind === 'textValue' ? '' : a.localName;
      const second = b.kind === 'textValue' ? '' : b.localName;
      return first < second ? -1 : first > second ? 1 : 0;
    });
  return equivalentNodes(byName(left), byName(right));
}

interface ContainerState {
  readonly prior: OoxmlElement | undefined;
  readonly next: OoxmlElement | undefined;
  readonly existing: OoxmlElement | undefined;
  /** What a reject restores: a pending record's snapshot, else the state before the command. */
  readonly original: readonly OoxmlNode[];
  readonly changedByCommand: boolean;
  /** The live properties after the command differ from the original. */
  readonly differs: boolean;
}
function containerState(
  container: Container,
  prior: OoxmlElement | undefined,
  next: OoxmlElement | undefined
): ContainerState {
  const changeName = `${container}Change`;
  const existing = prior?.children.find((child) => isW(child, changeName)) as
    | OoxmlElement
    | undefined;
  const original = recordable(
    (existing ? recordedProperties(existing) : null) ?? prior?.children ?? [],
    container
  );
  const body = (next?.children ?? []).filter((child) => !isW(child, changeName));
  return {
    prior,
    next,
    existing,
    original,
    changedByCommand: !sameHolder(prior, next),
    differs: !sameProperties(recordable(body, container), original),
  };
}

/**
 * The container after the command with its change record folded in. `needed` says whether the
 * container must carry a record: its own properties differ from the original, or a sibling in
 * its row does (rejecting a row record resets unrecorded cells), or the grid changed.
 */
function recorded(
  container: Container,
  state: ContainerState,
  stamp: Stamp,
  needed: boolean
): OoxmlElement | undefined {
  const { prior, next, existing, original } = state;
  const changeName = `${container}Change`;
  if (!state.changedByCommand) {
    // Untouched: a standing record stays exactly as its author wrote it while the row needs it.
    if (needed ? existing : !existing) return next;
  }
  // A holder the command created takes a mark id, so emptying it on reject removes it.
  let holder =
    next === undefined
      ? ({
          id: stamp.holderId(),
          kind: 'generic',
          namespaceUri: WML_NAMESPACE_URI,
          ...(stamp.wml.elementPrefix === undefined ? {} : { prefix: stamp.wml.elementPrefix }),
          namespaceBindings: [],
          attributes: [],
          children: [],
          localName: container,
        } as OoxmlElement)
      : prior === undefined
        ? ({ ...next, id: stamp.holderId() } as OoxmlElement)
        : next;
  const body = holder.children.filter((child) => !isW(child, changeName));
  holder = Object.freeze({ ...holder, children: body }) as OoxmlElement;
  const foreign =
    existing !== undefined &&
    attributeValueOf(existing, 'author', WML_NAMESPACE_URI) !== stamp.revision.author;
  if (!state.differs && foreign) {
    // Back at another author's original: their record stays as written.
    const kept = PATCH[container](holder, existing);
    if (!kept.ok) throw new Error('table property record refused');
    return kept.container;
  }
  if (!needed)
    return body.length === 0 && (prior === undefined || createdHolder(holder)) ? undefined : holder;
  // The ORIGINAL, never an intermediate state; a standing record is re-attributed here.
  const patched = PATCH[container](holder, history(stamp, container, original));
  if (!patched.ok) throw new Error('table property record refused');
  return patched.container;
}

/** Replace, insert, or drop a property holder at the head of its owner's children. */
function withHolder(
  owner: OoxmlElement,
  name: Container,
  prior: OoxmlElement | undefined,
  holder: OoxmlElement | undefined
): OoxmlElement {
  if (holder === prior) return owner;
  const current = wmlChildNamed(owner, name);
  let children: readonly OoxmlNode[];
  if (current && holder)
    children = owner.children.map((child) => (child === current ? holder : child));
  else if (current) children = owner.children.filter((child) => child !== current);
  else if (holder) {
    // `w:tblPrEx` precedes `w:trPr` in a row; every other holder leads its owner.
    const after = name === 'trPr' ? owner.children.findIndex((child) => isW(child, 'tblPrEx')) : -1;
    children = [...owner.children.slice(0, after + 1), holder, ...owner.children.slice(after + 1)];
  } else return owner;
  return Object.freeze({ ...owner, children }) as OoxmlElement;
}

/** A holder a tracked proposal created (mark-family id), absent from the source. */
const createdHolder = (node: OoxmlElement) =>
  node.id.slice(node.id.lastIndexOf('#') + 1).startsWith('mark:');
const sameHolder = (prior: OoxmlElement | undefined, next: OoxmlElement | undefined) =>
  prior === next || (prior !== undefined && next !== undefined && ooxmlTreesEqual(prior, next));

/** Decorate the command's table by comparing every property container before and after. */
function propertyMarkup(before: OoxmlElement, after: OoxmlElement, stamp: Stamp): OoxmlElement {
  const oldRows = new Map<string, OoxmlElement>();
  const oldCells = new Map<string, OoxmlElement>();
  for (const row of elements(before))
    if (row.kind === 'tableRow') {
      oldRows.set(row.id, row);
      for (const cell of elements(row)) if (cell.kind === 'tableCell') oldCells.set(cell.id, cell);
    }
  const oldGrid = wmlChildNamed(before, 'tblGrid');
  const newGrid = wmlChildNamed(after, 'tblGrid');
  const gridChanged = !sameHolder(oldGrid, newGrid);
  if (gridChanged && (!oldGrid || !newGrid || holds(oldGrid, STRUCTURE_MARKERS.tblGrid)))
    // One grid history per table: a second grid change waits for the first decision.
    throw new Error('A pending grid history cannot record another grid change.');
  const tblPrState = containerState(
    'tblPr',
    wmlChildNamed(before, 'tblPr'),
    wmlChildNamed(after, 'tblPr')
  );
  const tblPr = recorded('tblPr', tblPrState, stamp, tblPrState.differs);
  let table = withHolder(after, 'tblPr', wmlChildNamed(after, 'tblPr'), tblPr);
  const children = table.children.map((child): OoxmlNode => {
    if (gridChanged && child === newGrid)
      // CT_TblGridChange is CT_Markup: an id and no author.
      return Object.freeze({
        ...child,
        children: [...child.children, history(stamp, 'tblGrid', oldGrid!.children, false)],
      }) as OoxmlElement;
    if (child.kind !== 'tableRow') return child;
    const oldRow = oldRows.get(child.id);
    if (!oldRow) throw new Error('A table property command changed the row set.');
    const rowState = containerState(
      'trPr',
      wmlChildNamed(oldRow, 'trPr'),
      wmlChildNamed(child, 'trPr')
    );
    const cellStates = new Map(
      elements(child)
        .filter((cell) => cell.kind === 'tableCell')
        .map((cell) => {
          const old = oldCells.get(cell.id);
          if (!old) throw new Error('A table property command changed the cell set.');
          return [
            cell.id,
            containerState('tcPr', wmlChildNamed(old, 'tcPr'), wmlChildNamed(cell, 'tcPr')),
          ];
        })
    );
    const states = [rowState, ...cellStates.values()];
    if (!gridChanged && !states.some((state) => state.changedByCommand)) return child;
    // Rejecting any record in a row resets the row's unrecorded cells, so a row that still
    // differs from its original records every cell, and one back at it records none of this
    // author's.
    const rowNeeds = gridChanged || states.some((state) => state.differs);
    let row = withHolder(
      child,
      'trPr',
      rowState.next,
      recorded(
        'trPr',
        rowState,
        stamp,
        rowState.differs || (gridChanged && rowState.existing !== undefined)
      )
    );
    row = Object.freeze({
      ...row,
      children: row.children.map((cell): OoxmlNode => {
        if (cell.kind !== 'tableCell') return cell;
        const state = cellStates.get(cell.id)!;
        return withHolder(cell, 'tcPr', state.next, recorded('tcPr', state, stamp, rowNeeds));
      }),
    }) as OoxmlElement;
    return row;
  });
  table = Object.freeze({ ...table, children }) as OoxmlElement;
  return table;
}

/**
 * Apply a tracked table property command. `untracked` is the native applier, so accepting the
 * proposal lands exactly the direct result and rejecting it restores the original properties.
 */
export function applyTrackedTablePropertyOp(
  part: OoxmlPart,
  op: Tracked<TablePropertyOp>,
  untracked: (part: OoxmlPart, op: TablePropertyOp) => TreeOpResult,
  options?: EditOptions
): TreeOpResult {
  const refusal = trackedTablePropertyRefusal(part, op);
  if (refusal) return { ok: false, reason: refusal };
  const { revision, ...plain } = op;
  const before = findNode(part, op.tableId) as OoxmlElement;
  const applied = untracked(part, plain as TablePropertyOp);
  if (!applied.ok) return applied;
  const after = findNode(applied.part, op.tableId) as OoxmlElement;
  if (after === before || ooxmlTreesEqual(after, before)) return applied;
  const stamp = newStamp(part, before, revision);
  let marked: OoxmlElement;
  try {
    marked = tableWithDeclaredWmlBinding(
      propertyMarkup(before, after, stamp),
      applied.part,
      stamp.wml
    );
  } catch {
    return { ok: false, reason: 'unsupported' };
  }
  return fromEdit(replaceNode(applied.part, op.tableId, marked, options), {
    ...applied.effect,
    dirty: [...new Set([...applied.effect.dirty, op.tableId])],
  });
}
