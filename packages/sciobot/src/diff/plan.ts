import { isRetainedAtomBindings } from './retained-atoms';
import { isInsertedAtomBindings } from './inserted-atom-bindings';
import type { InsertedInlineAtom } from './inserted-atoms';
import { isTableGeometryIntent, type TableGeometryIntent } from './table-geometry';
import type { TreeDocOp } from '@docx-editor.dev/core/store';
import { isBlockId } from '../blocks/payload-schema';
import { jsonDataProblem, type JsonObject, type JsonValue } from '../data/json';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import type { BlockPropertyPath } from '../projection/block-tags';
import type { TargetAstNode } from '../projection/parse-target';
import { isInsertedBlockBindings, type InsertedBlockBinding } from './inserted-blocks';
import { isTargetFragment } from './target-fragment';
import { isInsertedRowIntent, type InsertedTableRowIntent } from './inserted-row-types';
import { isInsertedColumnIntent, type InsertedTableColumnIntent } from './inserted-column-types';
import { isSplitCellIntent, type SplitTableCellIntent } from './split-cell-types';
import { isInsertedListBindings, type InsertedListBinding } from './inserted-list-types';
export type { InsertedListBinding } from './inserted-list-types';
export type { InsertedBlockBinding } from './inserted-blocks';

/** A paragraph style the document defines, applied to the new paragraphs of one inserted block. */
export interface InsertedStyleBinding {
  readonly path: readonly number[];
  readonly styleId: string;
}
import {
  isListSourceState,
  isListTarget,
  type ListSourceState,
  type ListTarget,
} from './list-types';

export type PlannedValue =
  | { readonly present: false }
  | { readonly present: true; readonly value: JsonValue };

export interface PlannedPropertyChange {
  readonly name: string;
  readonly before: PlannedValue;
  readonly after: PlannedValue;
}

export type PlannedFormatTarget =
  | {
      readonly kind: 'run';
      readonly paragraphId: string;
      readonly start: number;
      readonly end: number;
    }
  | { readonly kind: 'paragraph'; readonly paragraphId: string }
  | { readonly kind: 'hyperlink'; readonly hyperlinkId: string };

const TABLE_KEYS = {
  insertTable: ['beforeParagraphId', 'rows', 'cols', 'columnWidthTwips'],
  insertTableRow: ['tableId', 'rowId', 'where'],
  deleteTableRow: ['tableId', 'rowId', 'referenceCellId'],
  insertTableColumn: ['tableId', 'where', 'gridColumnId', 'referenceCellId'],
  deleteTableColumn: ['tableId', 'gridColumnId'],
  mergeTableCells: ['tableId', 'cellIds'],
  splitTableCell: ['tableId', 'cellId', 'rows', 'columns', 'columnDistribution'],
  setTableColumnWidths: [
    'tableId',
    'leftGridColumnId',
    'rightGridColumnId',
    'leftWidthTwips',
    'rightWidthTwips',
  ],
  setTableRightEdgeWidth: ['tableId', 'gridColumnId', 'columnWidthTwips', 'tableWidthTwips'],
  setTableProperties: ['tableId', 'styleId', 'headerRowCount', 'columnWidthsTwips'],
  setTableRowHeight: ['tableId', 'rowId', 'heightTwips'],
  setTableCellBorders: ['tableId', 'cellIds', 'scope', 'target', 'spec'],
  setTableCellFill: ['tableId', 'cellIds', 'color'],
  setTableCellVerticalAlignment: ['tableId', 'cellIds', 'alignment'],
  authorTable: ['action'],
} as const satisfies Partial<Record<TreeDocOp['op'], readonly string[]>>;
type WithoutRevision<T> = T extends unknown ? Omit<T, 'revision'> : never;
/** Reuse native table commands. Executors supply review attribution and native validation. */
export type PlannedTableOperation = WithoutRevision<
  Extract<TreeDocOp, { op: keyof typeof TABLE_KEYS }>
>;

interface LocatedEdit {
  readonly partName: string;
}

/** Semantic edits contain addresses and new fragments, never a reconstructed canonical document. */
export type CanonicalEdit = LocatedEdit &
  (
    | {
        readonly kind: 'text';
        readonly paragraphId: string;
        readonly start: number;
        readonly end: number;
        readonly before: string;
        readonly after: string;
        /** Formatting ownership when inserted text falls exactly on a run boundary. */
        readonly bias?: 'left' | 'right';
      }
    | {
        readonly kind: 'inline-atoms';
        readonly paragraphId: string;
        /** Exact post-text native paragraph before atom insertion. */
        readonly before: string;
        readonly offset: number;
        readonly ast: TargetAstNode;
        readonly atoms: readonly InsertedInlineAtom[];
      }
    | {
        readonly kind: 'format';
        readonly target: PlannedFormatTarget;
        readonly changes: readonly PlannedPropertyChange[];
      }
    | {
        readonly kind: 'siblings';
        readonly parentId: string;
        /** Unchanged neighbors locate the affected sibling range. Null means a container edge. */
        readonly beforeId: string | null;
        readonly afterId: string | null;
        readonly removedIds: readonly string[];
        /** Only newly inserted target blocks. Existing canonical nodes stay in their package. */
        readonly inserted: readonly TargetAstNode[];
        /** Resolve these paths after allocating new native paragraphs. IDs never come from a preview. */
        readonly lists?: readonly InsertedListBinding[];
        /** Resolved document paragraph styles for inserted block quotes, by inserted index path. */
        readonly styles?: readonly InsertedStyleBinding[];
        /** New registered blocks and their new payloads, by inserted index path. */
        readonly blocks?: readonly InsertedBlockBinding[];
        readonly atoms?: readonly InsertedInlineAtom[];
      }
    | {
        /** Wrap an exact post-text range in a new hyperlink; the executor mints any relationship. */
        readonly kind: 'link';
        readonly paragraphId: string;
        readonly start: number;
        readonly end: number;
        readonly target:
          | { readonly kind: 'external'; readonly url: string }
          | { readonly kind: 'bookmark'; readonly anchor: string };
        readonly title: string | null;
      }
    | {
        /** Remove one existing hyperlink wrapper and keep its runs. */
        readonly kind: 'unlink';
        readonly hyperlinkId: string;
      }
    | { readonly kind: 'table'; readonly operation: PlannedTableOperation }
    | { readonly kind: 'insert-row'; readonly intent: InsertedTableRowIntent }
    | { readonly kind: 'insert-column'; readonly intent: InsertedTableColumnIntent }
    | { readonly kind: 'split-cell'; readonly intent: SplitTableCellIntent }
    | { readonly kind: 'table-geometry'; readonly intent: TableGeometryIntent }
    | {
        readonly kind: 'list';
        readonly paragraphId: string;
        readonly before: ListSourceState;
        readonly after: ListTarget;
      }
    | {
        readonly kind: 'payload';
        readonly blockId: string;
        readonly controlId: string;
        /** An empty path replaces the complete atomic payload and requires whole-block review. */
        readonly path: BlockPropertyPath;
        readonly before: PlannedValue;
        readonly after: PlannedValue;
      }
  );

/** Direct and suggestion executors consume the same mode-independent plan. */
export interface CanonicalEditPlan {
  readonly schemaVersion: 1;
  readonly revision: RevisionToken;
  readonly operations: readonly CanonicalEdit[];
}
export type EditPlanResult =
  | { readonly success: true; readonly plan: CanonicalEditPlan }
  | NativeEditFailure;

export const EDIT_PLAN_LIMITS = Object.freeze({
  maxBytes: 16 * 1024 * 1024,
  maxNodes: 200_000,
  maxDepth: 64,
});
const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const id = (value: unknown): value is string => typeof value === 'string' && value.length > 0;
const integer = (value: unknown): value is number =>
  Number.isSafeInteger(value) && Number(value) >= 0;
const keysOnly = (value: object, keys: readonly string[]) =>
  Object.keys(value).every((key) => keys.includes(key));

/** Style bindings name top-level inserted block quotes once, with a bounded printable style ID. */
function insertedStyles(inserted: readonly TargetAstNode[], value: unknown): boolean {
  if (!Array.isArray(value) || value.length > inserted.length) return false;
  const seen = new Set<number>();
  return value.every((binding) => {
    if (!record(binding) || !keysOnly(binding, ['path', 'styleId'])) return false;
    const path = binding.path,
      styleId = binding.styleId;
    return (
      Array.isArray(path) &&
      path.length === 1 &&
      integer(path[0]) &&
      !seen.has(path[0] as number) &&
      (seen.add(path[0] as number), true) &&
      inserted[path[0] as number]?.type === 'blockquote' &&
      typeof styleId === 'string' &&
      /^[\x21-\x7e][\x20-\x7e]{0,127}$/.test(styleId)
    );
  });
}

function equalJson(before: JsonValue, after: JsonValue): boolean {
  if (before === after) return true;
  if (before === null || after === null || typeof before !== 'object' || typeof after !== 'object')
    return false;
  if (Array.isArray(before) !== Array.isArray(after)) return false;
  const a = before as JsonObject,
    b = after as JsonObject;
  const keys = Object.keys(a);
  return (
    keys.length === Object.keys(b).length &&
    keys.every((key) => Object.hasOwn(b, key) && equalJson(a[key]!, b[key]!))
  );
}

function valueState(value: unknown): value is PlannedValue {
  return (
    record(value) &&
    (value.present === false
      ? keysOnly(value, ['present'])
      : value.present === true &&
        Object.hasOwn(value, 'value') &&
        keysOnly(value, ['present', 'value']))
  );
}
function sameState(before: PlannedValue, after: PlannedValue): boolean {
  return (
    before.present === after.present &&
    (!before.present || (after.present && equalJson(before.value, after.value)))
  );
}
function freeze(value: unknown): void {
  if (value === null || typeof value !== 'object') return;
  for (const child of Object.values(value)) freeze(child);
  Object.freeze(value);
}
function formatTarget(value: unknown): boolean {
  if (!record(value)) return false;
  if (value.kind === 'hyperlink')
    return id(value.hyperlinkId) && keysOnly(value, ['kind', 'hyperlinkId']);
  if (!id(value.paragraphId)) return false;
  if (value.kind === 'paragraph') return keysOnly(value, ['kind', 'paragraphId']);
  return (
    value.kind === 'run' &&
    integer(value.start) &&
    integer(value.end) &&
    value.start <= value.end &&
    keysOnly(value, ['kind', 'paragraphId', 'start', 'end'])
  );
}

function tableData(value: unknown): boolean {
  if (value === null || typeof value !== 'object') return true;
  if (Object.hasOwn(value, 'children')) return false;
  return Object.values(value).every(tableData);
}

function tableAddress(operation: Record<string, unknown>): boolean {
  if (operation.op !== 'authorTable')
    return id(operation.op === 'insertTable' ? operation.beforeParagraphId : operation.tableId);
  const action = operation.action;
  if (!record(action)) return false;
  if (action.kind === 'insert')
    return (
      id(action.paragraphId) &&
      keysOnly(action, ['kind', 'paragraphId', 'offset', 'rowCount', 'columnCount', 'values'])
    );
  return (
    action.kind === 'existing' &&
    id(action.tableId) &&
    record(action.mutation) &&
    keysOnly(action, ['kind', 'tableId', 'mutation'])
  );
}

/** Structural schema checks do not replace canonical address, geometry, or mode validation during application. */
function editProblem(edit: unknown): string | null {
  if (
    !record(edit) ||
    !id(edit.partName) ||
    !edit.partName.startsWith('/') ||
    edit.partName.split('/').includes('..')
  )
    return 'invalid-part-name';
  const base = ['kind', 'partName'];
  switch (edit.kind) {
    case 'text':
      return id(edit.paragraphId) &&
        integer(edit.start) &&
        integer(edit.end) &&
        edit.start <= edit.end &&
        typeof edit.before === 'string' &&
        typeof edit.after === 'string' &&
        edit.before.length === edit.end - edit.start &&
        (!Object.hasOwn(edit, 'bias') || edit.bias === 'left' || edit.bias === 'right') &&
        keysOnly(edit, [...base, 'paragraphId', 'start', 'end', 'before', 'after', 'bias'])
        ? null
        : 'invalid-text-edit';
    case 'inline-atoms':
      return id(edit.paragraphId) &&
        integer(edit.offset) &&
        typeof edit.before === 'string' &&
        record(edit.ast) &&
        isTargetFragment(edit.ast) &&
        isRetainedAtomBindings(
          edit.paragraphId,
          edit.offset,
          edit.before,
          edit.ast as unknown as TargetAstNode,
          edit.atoms
        ) &&
        keysOnly(edit, [...base, 'paragraphId', 'before', 'offset', 'ast', 'atoms'])
        ? null
        : 'invalid-inline-atoms-edit';
    case 'format': {
      if (
        !formatTarget(edit.target) ||
        !Array.isArray(edit.changes) ||
        !keysOnly(edit, [...base, 'target', 'changes'])
      )
        return 'invalid-format-edit';
      const names = new Set<string>();
      for (const change of edit.changes) {
        if (
          !record(change) ||
          !id(change.name) ||
          names.has(change.name) ||
          !valueState(change.before) ||
          !valueState(change.after) ||
          !keysOnly(change, ['name', 'before', 'after'])
        )
          return 'invalid-property-change';
        names.add(change.name);
      }
      return null;
    }
    case 'siblings':
      return id(edit.parentId) &&
        (edit.beforeId === null || id(edit.beforeId)) &&
        (edit.afterId === null || id(edit.afterId)) &&
        (edit.beforeId === null || edit.beforeId !== edit.afterId) &&
        Array.isArray(edit.removedIds) &&
        edit.removedIds.every(id) &&
        new Set(edit.removedIds).size === edit.removedIds.length &&
        (edit.beforeId === null || !edit.removedIds.includes(edit.beforeId)) &&
        (edit.afterId === null || !edit.removedIds.includes(edit.afterId)) &&
        Array.isArray(edit.inserted) &&
        edit.inserted.every(isTargetFragment) &&
        (!Object.hasOwn(edit, 'lists') || isInsertedListBindings(edit.inserted, edit.lists)) &&
        (!Object.hasOwn(edit, 'styles') || insertedStyles(edit.inserted, edit.styles)) &&
        (!Object.hasOwn(edit, 'blocks') || isInsertedBlockBindings(edit.inserted, edit.blocks)) &&
        (!Object.hasOwn(edit, 'atoms') ||
          isInsertedAtomBindings(
            edit.inserted,
            (edit.blocks ?? []) as InsertedBlockBinding[],
            edit.atoms
          )) &&
        keysOnly(edit, [
          ...base,
          'parentId',
          'beforeId',
          'afterId',
          'removedIds',
          'inserted',
          'lists',
          'styles',
          'blocks',
          'atoms',
        ])
        ? null
        : 'invalid-sibling-edit';
    case 'link': {
      const target = edit.target as unknown;
      return id(edit.paragraphId) &&
        integer(edit.start) &&
        integer(edit.end) &&
        edit.start < edit.end &&
        record(target) &&
        ((target.kind === 'external' &&
          typeof target.url === 'string' &&
          target.url.length > 0 &&
          target.url.length <= 4096 &&
          keysOnly(target, ['kind', 'url'])) ||
          (target.kind === 'bookmark' &&
            typeof target.anchor === 'string' &&
            target.anchor.length > 0 &&
            target.anchor.length <= 4096 &&
            keysOnly(target, ['kind', 'anchor']))) &&
        (edit.title === null || (typeof edit.title === 'string' && edit.title.length <= 4096)) &&
        keysOnly(edit, [...base, 'paragraphId', 'start', 'end', 'target', 'title'])
        ? null
        : 'invalid-link-edit';
    }
    case 'unlink':
      return id(edit.hyperlinkId) && keysOnly(edit, [...base, 'hyperlinkId'])
        ? null
        : 'invalid-unlink-edit';
    case 'payload':
      return isBlockId(edit.blockId) &&
        id(edit.controlId) &&
        Array.isArray(edit.path) &&
        edit.path.every((step) => typeof step === 'string' || integer(step)) &&
        valueState(edit.before) &&
        valueState(edit.after) &&
        keysOnly(edit, [...base, 'blockId', 'controlId', 'path', 'before', 'after'])
        ? null
        : 'invalid-payload-edit';
    case 'list':
      return id(edit.paragraphId) &&
        isListSourceState(edit.before) &&
        isListTarget(edit.after) &&
        keysOnly(edit, [...base, 'paragraphId', 'before', 'after'])
        ? null
        : 'invalid-list-edit';
    case 'table':
      if (
        !record(edit.operation) ||
        typeof edit.operation.op !== 'string' ||
        !Object.hasOwn(TABLE_KEYS, edit.operation.op) ||
        !keysOnly(edit.operation, [
          'op',
          ...TABLE_KEYS[edit.operation.op as keyof typeof TABLE_KEYS],
        ]) ||
        !tableData(edit.operation) ||
        !keysOnly(edit, [...base, 'operation'])
      )
        return 'invalid-table-edit';
      return tableAddress(edit.operation) ? null : 'invalid-table-address';
    case 'insert-row':
      return isInsertedRowIntent(edit.intent) && keysOnly(edit, [...base, 'intent'])
        ? null
        : 'invalid-inserted-row';
    case 'insert-column':
      return isInsertedColumnIntent(edit.intent) && keysOnly(edit, [...base, 'intent'])
        ? null
        : 'invalid-inserted-column';
    case 'table-geometry':
      return isTableGeometryIntent(edit.intent) && keysOnly(edit, [...base, 'intent'])
        ? null
        : 'invalid-table-geometry';
    case 'split-cell':
      return isSplitCellIntent(edit.intent) && keysOnly(edit, [...base, 'intent'])
        ? null
        : 'invalid-split-cell';
    default:
      return 'unknown-edit-kind';
  }
}

function withoutNoop(edit: CanonicalEdit): CanonicalEdit | null {
  switch (edit.kind) {
    case 'text':
      return edit.before === edit.after ? null : edit;
    case 'payload':
      return sameState(edit.before, edit.after) ? null : edit;
    case 'format': {
      const changes = edit.changes.filter((change) => !sameState(change.before, change.after));
      return changes.length ? { ...edit, changes } : null;
    }
    case 'siblings':
      return edit.removedIds.length || edit.inserted.length ? edit : null;
    case 'inline-atoms':
    case 'table':
    case 'insert-row':
    case 'insert-column':
    case 'split-cell':
    case 'table-geometry':
    case 'link':
    case 'unlink':
      return edit;
    case 'list': {
      const { before, after } = edit;
      if (after.kind === 'none') return before.numbering || before.checkbox ? edit : null;
      if (after.kind === 'checklist')
        return !before.numbering &&
          before.checkbox?.checked === after.checked &&
          before.checkbox.level === after.level
          ? null
          : edit;
      return !before.checkbox &&
        after.instance.kind === 'existing' &&
        before.numbering?.numId === after.instance.numId &&
        before.numbering.level === after.level
        ? null
        : edit;
    }
  }
}

function jsonCopy(value: unknown): unknown {
  const problem = jsonDataProblem(value, EDIT_PLAN_LIMITS);
  if (problem) throw new TypeError(problem);
  const raw = JSON.stringify(value);
  if (
    raw.length > EDIT_PLAN_LIMITS.maxBytes ||
    new TextEncoder().encode(raw).byteLength > EDIT_PLAN_LIMITS.maxBytes
  )
    throw new RangeError('byte-limit');
  return JSON.parse(raw);
}

/** Validate before reading input fields or invoking JSON.stringify. Never retain caller-owned objects. */
export function createEditPlan(
  revision: RevisionToken,
  operations: readonly CanonicalEdit[]
): EditPlanResult {
  try {
    const copy = jsonCopy({ revision, operations }) as {
      revision: RevisionToken;
      operations: unknown[];
    };
    const token = copy.revision;
    if (
      !record(token) ||
      !id(token.documentId) ||
      !id(token.epoch) ||
      !integer(token.revision) ||
      !keysOnly(token, ['documentId', 'epoch', 'revision'])
    )
      throw new TypeError('invalid-revision');
    if (!Array.isArray(copy.operations)) throw new TypeError('expected-operation-array');
    for (const edit of copy.operations) {
      const problem = editProblem(edit);
      if (problem) throw new TypeError(problem);
    }
    const insertedIds = new Set<string>();
    for (const edit of copy.operations as CanonicalEdit[]) {
      const blocks =
        edit.kind === 'siblings'
          ? (edit.blocks ?? [])
          : edit.kind === 'insert-row' ||
              edit.kind === 'insert-column' ||
              edit.kind === 'split-cell'
            ? edit.intent.payloads
            : [];
      const atoms =
        edit.kind === 'inline-atoms'
          ? edit.atoms
          : edit.kind === 'siblings'
            ? (edit.atoms ?? [])
            : [];
      for (const id of [
        ...blocks.map((binding) => binding.payload.id),
        ...atoms.flatMap((atom) => (atom.kind === 'inline-control' ? [atom.payload.id] : [])),
      ]) {
        if (insertedIds.has(id)) throw new TypeError('duplicate-inserted-block-id');
        insertedIds.add(id);
      }
    }
    const instances = new Map<string, string>();
    for (const edit of copy.operations as CanonicalEdit[]) {
      const targets =
        edit.kind === 'list'
          ? [edit.after]
          : edit.kind === 'siblings'
            ? (edit.lists ?? []).map((binding) => binding.target)
            : edit.kind === 'insert-row' || edit.kind === 'insert-column'
              ? (edit.intent.lists ?? []).map((binding) => binding.target)
              : [];
      for (const target of targets) {
        if (
          (target.kind !== 'ordered' && target.kind !== 'bullet') ||
          target.instance.kind !== 'new'
        )
          continue;
        const levelKey = JSON.stringify([target.instance.key, target.level]),
          definition = JSON.stringify([
            target.kind,
            target.kind === 'ordered' ? (target.instance.start ?? 1) : null,
          ]),
          old = instances.get(levelKey);
        if (old !== undefined && old !== definition)
          throw new TypeError('conflicting-new-list-instance');
        instances.set(levelKey, definition);
      }
    }
    const edits = (copy.operations as CanonicalEdit[])
      .map(withoutNoop)
      .filter((edit): edit is CanonicalEdit => edit !== null);
    const plan: CanonicalEditPlan = { schemaVersion: 1, revision: token, operations: edits };
    freeze(plan);
    return { success: true, plan };
  } catch (cause) {
    return {
      success: false,
      code: 'invalid-input',
      error: cause instanceof Error ? cause.message : String(cause),
    };
  }
}

/** Revalidate serialized data. Unknown versions and extra root fields cannot become executable plans. */
export function decodeEditPlan(raw: string): EditPlanResult {
  try {
    if (
      typeof raw !== 'string' ||
      raw.length > EDIT_PLAN_LIMITS.maxBytes ||
      new TextEncoder().encode(raw).byteLength > EDIT_PLAN_LIMITS.maxBytes
    )
      throw new TypeError('byte-limit');
    const value: unknown = JSON.parse(raw);
    if (
      !record(value) ||
      value.schemaVersion !== 1 ||
      !keysOnly(value, ['schemaVersion', 'revision', 'operations'])
    )
      throw new TypeError('invalid-plan-schema');
    return createEditPlan(value.revision as RevisionToken, value.operations as CanonicalEdit[]);
  } catch (cause) {
    return {
      success: false,
      code: 'invalid-input',
      error: cause instanceof Error ? cause.message : String(cause),
    };
  }
}

export function serializeEditPlan(plan: CanonicalEditPlan): string {
  const copy = jsonCopy(plan) as CanonicalEditPlan;
  if (
    !record(copy) ||
    copy.schemaVersion !== 1 ||
    !keysOnly(copy, ['schemaVersion', 'revision', 'operations'])
  )
    throw new TypeError('invalid-plan-schema');
  const result = createEditPlan(copy.revision, copy.operations);
  if (!result.success) throw new TypeError(result.error);
  return JSON.stringify(result.plan);
}
