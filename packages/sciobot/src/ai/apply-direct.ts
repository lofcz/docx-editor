import { cellPropertiesAt } from '../diff/new-cell-properties';
import { cellListBindings } from '../diff/inserted-cell-lists';
import { resolveTableGeometry, geometrySplitCells } from '../diff/table-geometry';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  directParagraphProperties,
  ensureHyperlinkRelationship,
  findNode,
  hyperlinkTargetOf,
  paragraphTextOf,
  relationshipTargetIn,
  withPart,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
import type { BlockPayload } from '../blocks/payload-schema';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { jsonDataKey } from '../diff/target-key';
import { planListMembership } from '../diff/lists';
import {
  decodeEditPlan,
  serializeEditPlan,
  type CanonicalEdit,
  type CanonicalEditPlan,
  type PlannedValue,
} from '../diff/plan';
import { directParagraphTextOps } from '../diff/paragraph-text';
import {
  insertedColumnIndex,
  insertedColumnOperation,
  insertedColumnWidthOperation,
} from '../diff/inserted-column-native';
import { nativeVisualSize, splitResultCells } from '../diff/split-cell-native';
import { readEditorRevision } from '../projection/revision-cache';
import { trackedRunFormatOps } from '../review/tracked-run-format';
import {
  SCIOBOT_CODE_STYLE_ID,
  inlineCodeCharacterStyles,
  withInlineCodeStyle,
} from '../native/inline-code';
import {
  applyDirectListAssignments,
  createDirectListContext,
  type DirectListContext,
  type DirectListAssignment,
} from './direct-lists';
import {
  completeSiblingFragments,
  fillNewCells,
  NativeWriteRefusal,
  nodeIds,
  refuseWrite as refuse,
  rowCellIds,
  siblingFragmentSeeds,
  tableRowIds,
  wmlElement,
} from '../diff/fragment-seeds';

/** Mount identity for the editor that owns the plan's revision token. */
export interface DirectPlanContext {
  readonly documentId: string;
  readonly epoch: string;
}
export type DirectPlanResult =
  | {
      readonly success: true;
      /** Revision after the commit; equal to the plan revision when nothing changed. */
      readonly revision: RevisionToken;
      readonly committed: boolean;
    }
  | NativeEditFailure;

type Edit<K extends CanonicalEdit['kind']> = Extract<CanonicalEdit, { kind: K }>;
/** Work that needs identities allocated by earlier native operations in the same transaction. */
type Deferred = (
  pkg: OoxmlPackage,
  original: ReadonlySet<string>,
  listContext: DirectListContext
) => OoxmlPackage;
/** Hyperlink targets resolve through the surface's gated relationship mint at commit time. */
type Staged =
  | TreeDocOp
  | { readonly retarget: Edit<'format'>; readonly link: LinkTarget }
  | { readonly linkEdit: Edit<'link'> };
interface LinkTarget {
  readonly linkId: string;
  readonly url?: string;
  readonly relationshipId?: string;
  readonly anchor?: string;
  readonly tooltip?: string;
}

const failure = (error: unknown): NativeEditFailure =>
  error instanceof NativeWriteRefusal
    ? { success: false, code: error.code, error: error.message }
    : {
        success: false,
        code: 'apply-failed',
        error: error instanceof Error ? error.message : String(error),
      };
const LINK_SCRATCH = { author: 'direct', date: '2000-01-01T00:00:00Z' };
const SIBLING_PARENTS = new Set(['body', 'tc', 'sdtContent']);
const wml = wmlElement;
const hasSection = (node: OoxmlNode): boolean =>
  wml(node, 'sectPr') || (node.kind !== 'textValue' && node.children.some(hasSection));
const attributesOf = (value: PlannedValue): Record<string, string> => {
  if (
    !value.present ||
    !value.value ||
    typeof value.value !== 'object' ||
    Array.isArray(value.value) ||
    !Object.values(value.value).every((entry) => typeof entry === 'string')
  )
    refuse('Invalid native property value.');
  return value.value as Record<string, string>;
};
const sameState = (actual: Readonly<Record<string, string>> | undefined, expected: PlannedValue) =>
  expected.present
    ? actual !== undefined && jsonDataKey(actual) === jsonDataKey(attributesOf(expected))
    : actual === undefined;

function sibling(part: OoxmlPart, edit: Edit<'siblings'>): readonly string[] {
  const parent = findNode(part, edit.parentId);
  if (!parent || parent.kind === 'textValue' || !SIBLING_PARENTS.has(parent.localName))
    refuse('The planned sibling parent is not a native story container.', 'ambiguous-target');
  const children = (parent as OoxmlElement).children,
    before = edit.beforeId === null ? -1 : children.findIndex((node) => node.id === edit.beforeId),
    terminal = children.findIndex((node) => wml(node, 'sectPr')),
    after =
      edit.afterId === null
        ? terminal < 0
          ? children.length
          : terminal
        : children.findIndex((node) => node.id === edit.afterId);
  if ((edit.beforeId !== null && before < 0) || after < 0 || before >= after)
    refuse('The original native sibling gap no longer matches the plan.', 'stale-revision');
  const gap = children.slice(before + 1, after).map((node) => node.id);
  if (gap.length !== edit.removedIds.length || gap.some((id, i) => id !== edit.removedIds[i]))
    refuse('The original native sibling gap no longer matches the plan.', 'stale-revision');
  return gap;
}
/** New paragraphs between the sibling anchors that were absent before this transaction. */
function freshInGap(
  part: OoxmlPart,
  edit: Edit<'siblings'>,
  original: ReadonlySet<string>
): string[] {
  const parent = findNode(part, edit.parentId) as OoxmlElement,
    before =
      edit.beforeId === null ? -1 : parent.children.findIndex((node) => node.id === edit.beforeId),
    after =
      edit.afterId === null
        ? parent.children.length
        : parent.children.findIndex((node) => node.id === edit.afterId);
  return parent.children
    .slice(before + 1, after)
    .filter((node) => node.kind === 'paragraph' && !original.has(node.id))
    .map((node) => node.id);
}
interface Lowering {
  readonly staged: readonly Staged[];
  readonly deferred: readonly Deferred[];
}

/**
 * Validate every planned old value against the current canonical package and lower the plan into
 * native operations plus identity-dependent completions. Nothing here touches a live editor.
 */
/** Whether new target fragments in an operation contain an inline code node. */
function insertsInlineCode(value: unknown): boolean {
  if (Array.isArray(value)) return value.some(insertsInlineCode);
  if (value === null || typeof value !== 'object') return false;
  const node = value as Record<string, unknown>;
  return node.type === 'inlineCode' || Object.values(node).some(insertsInlineCode);
}

function lowerDirectPlan(pkg: OoxmlPackage, plan: CanonicalEditPlan): Lowering {
  const partName = pkg.mainDocumentPart,
    staged: Staged[] = [],
    deferred: Deferred[] = [],
    lists: DirectListAssignment[] = [],
    payloads = new Map<string, BlockPayload>();
  let preview = pkg.parts.get(partName)!;
  let deferredTopology = false;
  const apply = (op: TreeDocOp) => {
    const result = applyTreeOp(preview, op);
    if (!result.ok) refuse(`Native operation ${op.op} refused: ${result.reason}.`, 'apply-failed');
    preview = (result as { part: OoxmlPart }).part;
    staged.push(op);
  };
  for (const edit of plan.operations) {
    if (edit.partName !== partName)
      refuse('Direct AI plans address the main document story only.', 'unsupported-operation');
    switch (edit.kind) {
      case 'text': {
        const text = paragraphTextOf(preview, edit.paragraphId);
        if (
          text === null ||
          edit.end > text.length ||
          text.slice(edit.start, edit.end) !== edit.before
        )
          refuse('Planned paragraph text no longer matches the document.', 'stale-revision');
        directParagraphTextOps(edit).forEach(apply);
        break;
      }
      case 'format': {
        const target = edit.target;
        if (target.kind === 'run') {
          const lowered = trackedRunFormatOps(preview, [edit], LINK_SCRATCH);
          if (!lowered.success) refuse(lowered.error, lowered.code);
          for (const op of (lowered as { ops: readonly TreeDocOp[] }).ops) {
            const { revision: _revision, ...direct } = op as Extract<
              TreeDocOp,
              { op: 'setRunProperties' }
            >;
            apply(direct);
          }
        } else if (target.kind === 'paragraph') {
          if (findNode(preview, target.paragraphId)?.kind !== 'paragraph')
            refuse('The planned paragraph does not exist.', 'stale-revision');
          const properties = new Map(
            directParagraphProperties(preview, target.paragraphId).map((property) => [
              property.localName,
              property,
            ])
          );
          for (const change of edit.changes) {
            if (
              !sameState(
                properties.get(change.name)?.attributes ??
                  (properties.has(change.name) ? {} : undefined),
                change.before
              )
            )
              refuse(
                'Planned paragraph properties no longer match their before-values.',
                'stale-revision'
              );
            if (change.after.present)
              properties.set(change.name, {
                localName: change.name,
                attributes: attributesOf(change.after),
              });
            else properties.delete(change.name);
          }
          apply({
            op: 'setParagraphProperties',
            paragraphId: target.paragraphId,
            properties: [...properties.values()],
          });
        } else {
          const link = findNode(preview, target.hyperlinkId);
          if (link?.kind !== 'hyperlink')
            refuse('The planned hyperlink does not exist.', 'stale-revision');
          const current = hyperlinkTargetOf(
            link as Extract<OoxmlNode, { kind: 'hyperlink' }>,
            (id) => relationshipTargetIn(pkg, partName, id)
          );
          let next: LinkTarget = {
            linkId: target.hyperlinkId,
            ...(current.relationshipId !== undefined
              ? { relationshipId: current.relationshipId }
              : {}),
            ...(current.anchor !== undefined ? { anchor: current.anchor } : {}),
          };
          for (const change of edit.changes) {
            if (change.name === 'target') {
              const before = change.before.present
                ? (change.before.value as Record<string, unknown>)
                : null;
              if (!before || before.kind !== current.kind || before.authored !== current.authored)
                refuse('The planned hyperlink target no longer matches.', 'stale-revision');
              const after = change.after.present
                ? (change.after.value as Record<string, unknown>)
                : null;
              if (!after || typeof after.authored !== 'string')
                refuse(
                  'A hyperlink target cannot be removed by a format plan.',
                  'unsupported-operation'
                );
              next =
                after!.kind === 'external'
                  ? {
                      linkId: next.linkId,
                      url: after!.authored as string,
                      ...(next.tooltip ? { tooltip: next.tooltip } : {}),
                    }
                  : {
                      linkId: next.linkId,
                      anchor: after!.authored as string,
                      ...(next.tooltip ? { tooltip: next.tooltip } : {}),
                    };
            } else if (change.name === 'tooltip') {
              const actual = current.tooltip;
              if (change.before.present ? change.before.value !== actual : actual !== undefined)
                refuse('The planned hyperlink tooltip no longer matches.', 'stale-revision');
              if (!change.after.present)
                refuse(
                  'Removing a hyperlink tooltip needs native link structure planning (D27).',
                  'unsupported-operation'
                );
              next = { ...next, tooltip: String(change.after.present ? change.after.value : '') };
            } else refuse('Unknown native hyperlink change.');
          }
          // Validate the retarget on the preview with a scratch relationship; the live mint is gated.
          let scratch = pkg;
          const relationshipId =
            next.url !== undefined
              ? (() => {
                  const made = ensureHyperlinkRelationship(scratch, next.url!, partName);
                  if (!made) refuse('The native engine refuses this URL.', 'unsupported-operation');
                  scratch = made!.pkg;
                  return made!.relationshipId;
                })()
              : next.relationshipId;
          const op: TreeDocOp = {
            op: 'setHyperlinkTarget',
            linkId: next.linkId,
            ...(relationshipId !== undefined ? { relationshipId } : {}),
            ...(next.anchor !== undefined ? { anchor: next.anchor } : {}),
            ...(next.tooltip !== undefined ? { tooltip: next.tooltip } : {}),
          };
          const result = applyTreeOp(preview, op);
          if (!result.ok)
            refuse(`Native hyperlink retarget refused: ${result.reason}.`, 'apply-failed');
          preview = (result as { part: OoxmlPart }).part;
          staged.push({ retarget: edit, link: next });
        }
        break;
      }
      case 'siblings': {
        sibling(preview, edit);
        for (const id of edit.removedIds) {
          const node = findNode(preview, id);
          if (!node || (node.kind !== 'paragraph' && !wml(node, 'tbl')) || hasSection(node))
            refuse(
              'Direct sibling removal supports complete paragraphs and tables without section marks.',
              'unsupported-operation'
            );
        }
        const prepared = siblingFragmentSeeds(
          edit,
          edit.afterId !== null && findNode(preview, edit.afterId)?.kind === 'paragraph'
        );
        if (!prepared.success) refuse(prepared.error, prepared.code);
        const seeds = prepared.seeds;
        const lists = edit.lists ?? [];
        for (const binding of lists)
          if (binding.path.length === 1 && !seeds.paragraphOf.has(binding.path[0]!))
            refuse('An inserted list binding names no new paragraph.', 'apply-failed');
        for (const id of edit.removedIds) apply({ op: 'deleteBlock', blockId: id });
        if (seeds.paragraphs.length)
          apply({
            op: 'insertParagraphs',
            parentId: edit.parentId,
            beforeId: edit.afterId,
            paragraphs: seeds.paragraphs,
          });
        if (seeds.links.length || seeds.styles.size || seeds.placed.length || lists.length)
          deferred.push((working, original, listContext) => {
            const fresh = freshInGap(working.parts.get(partName)!, edit, original);
            const completed = completeSiblingFragments(working, partName, edit, seeds, fresh);
            return applyDirectListAssignments(
              completed.pkg,
              partName,
              lists.map((binding) => {
                const paragraphId = completed.paragraphs.get(binding.path.join('/'));
                if (!paragraphId)
                  refuse('An inserted list paragraph was not allocated.', 'apply-failed');
                return { paragraphId, before: null, target: binding.target };
              }),
              listContext
            );
          });
        break;
      }
      case 'link': {
        const text = paragraphTextOf(preview, edit.paragraphId);
        if (text === null || edit.end > text.length)
          refuse('The planned link range no longer exists.', 'stale-revision');
        let scratch: string | undefined;
        if (edit.target.kind === 'external') {
          const made = ensureHyperlinkRelationship(pkg, edit.target.url, partName);
          if (!made) refuse('The native engine refuses this URL.', 'unsupported-operation');
          scratch = made.relationshipId;
        }
        const op = wrapOp(edit, scratch);
        const result = applyTreeOp(preview, op);
        if (!result.ok) refuse(`Native link insertion refused: ${result.reason}.`, 'apply-failed');
        preview = result.part;
        staged.push({ linkEdit: edit });
        break;
      }
      case 'unlink':
        if (findNode(preview, edit.hyperlinkId)?.kind !== 'hyperlink')
          refuse('The planned hyperlink no longer exists.', 'stale-revision');
        apply({ op: 'removeHyperlink', linkId: edit.hyperlinkId });
        break;
      case 'list': {
        // Re-read the native source state through the planner that produced the plan.
        const current = planListMembership(
          { revision: plan.revision, pkg: withPart(pkg, preview), partName },
          [{ paragraphId: edit.paragraphId, target: edit.after }]
        );
        const reread = current.success ? current.plan.operations[0] : null;
        if (
          !reread ||
          reread.kind !== 'list' ||
          reread.before.fingerprint !== edit.before.fingerprint
        )
          refuse('Planned list state no longer matches the document.', 'stale-revision');
        lists.push({ paragraphId: edit.paragraphId, before: edit.before, target: edit.after });
        break;
      }
      case 'table-geometry':
        deferredTopology = true;
        deferred.push((working) => {
          const part = working.parts.get(partName)!;
          const resolved = resolveTableGeometry(part, edit.intent);
          if (!resolved.success) refuse(resolved.error, resolved.code);
          const result = applyTreeOp(part, resolved.operation);
          if (!result.ok)
            refuse(`Native table geometry refused: ${result.reason}.`, 'apply-failed');
          const next = withPart(working, result.part);
          if (edit.intent.kind === 'merge') return next;
          const cells = geometrySplitCells(result.part, edit.intent, resolved.cellId);
          if (!cells || cells.length !== edit.intent.cells.length)
            refuse('Native geometry split allocation differs from its plan.', 'apply-failed');
          return fillNewCells(
            next,
            partName,
            cells.map((cellId, index) => ({
              cellId,
              blocks: edit.intent.kind === 'split' ? edit.intent.cells[index]! : [],
              properties:
                edit.intent.kind === 'split'
                  ? cellPropertiesAt(edit.intent.cellProperties, index)
                  : undefined,
            }))
          );
        });
        break;
      case 'table':
        if (deferredTopology)
          deferred.push((working) => {
            const result = applyTreeOp(working.parts.get(partName)!, edit.operation as TreeDocOp);
            if (!result.ok)
              refuse(`Native table operation refused: ${result.reason}.`, 'apply-failed');
            return withPart(working, result.part);
          });
        else apply(edit.operation as TreeDocOp);
        break;
      case 'insert-row': {
        deferredTopology = true;
        const intent = edit.intent;
        if (intent.payloads.length)
          refuse(
            'Registered blocks in new rows need the D29 fragment executor.',
            'unsupported-operation'
          );
        if (!tableRowIds(preview, intent.tableId).includes(intent.rowId))
          refuse('The planned table row anchor no longer exists.', 'stale-revision');
        deferred.push((working, _original, listContext) => {
          let part = working.parts.get(partName)!;
          const before = new Set(tableRowIds(part, intent.tableId));
          const inserted = applyTreeOp(part, {
            op: 'insertTableRow',
            tableId: intent.tableId,
            rowId: intent.rowId,
            where: intent.where,
          });
          if (!inserted.ok)
            refuse(`Native row insertion refused: ${inserted.reason}.`, 'apply-failed');
          part = (inserted as { part: OoxmlPart }).part;
          const rows = tableRowIds(part, intent.tableId).filter((id) => !before.has(id));
          if (rows.length !== 1) refuse('Native row allocation is ambiguous.', 'apply-failed');
          if (intent.properties.heightTwips !== undefined) {
            const sized = applyTreeOp(part, {
              op: 'setTableRowHeight',
              tableId: intent.tableId,
              rowId: rows[0]!,
              heightTwips: intent.properties.heightTwips,
            });
            if (!sized.ok) refuse(`Native row height refused: ${sized.reason}.`, 'apply-failed');
            part = (sized as { part: OoxmlPart }).part;
          }
          const cells = rowCellIds(part, rows[0]!);
          if (cells.length !== intent.cells.length)
            refuse('The new native row does not match the planned cells.', 'apply-failed');
          return fillNewCells(
            withPart(working, part),
            partName,
            cells.map((cellId, index) => ({
              cellId,
              blocks: intent.cells[index]!,
              lists: cellListBindings(intent.lists, index),
              properties: cellPropertiesAt(intent.cellProperties, index),
            })),
            listContext
          );
        });
        break;
      }
      case 'insert-column': {
        deferredTopology = true;
        const intent = edit.intent;
        if (intent.payloads.length)
          refuse(
            'Registered blocks in new columns need the D29 fragment executor.',
            'unsupported-operation'
          );
        deferred.push((working, _original, listContext) => {
          const before = working.parts.get(partName)!;
          const inserted = applyTreeOp(before, insertedColumnOperation(intent));
          if (!inserted.ok)
            refuse(`Native column insertion refused: ${inserted.reason}.`, 'apply-failed');
          let part = (inserted as { part: OoxmlPart }).part;
          if (insertedColumnIndex(before, part, intent) === null)
            refuse('Native column allocation is ambiguous.', 'apply-failed');
          const sizing = insertedColumnWidthOperation(before, part, intent);
          if (!sizing.ok) refuse(`Native column width refused: ${sizing.reason}.`, 'apply-failed');
          if (sizing.operation) {
            const sized = applyTreeOp(part, sizing.operation);
            if (!sized.ok) refuse(`Native column width refused: ${sized.reason}.`, 'apply-failed');
            part = (sized as { part: OoxmlPart }).part;
          }
          const cells = intent.cells.flatMap((entry, index) => {
            const old = new Set(rowCellIds(before, entry.rowId)),
              created = rowCellIds(part, entry.rowId).filter((id) => !old.has(id));
            if (created.length > 1 || (!created.length && entry.blocks.length))
              refuse('The new native column cell is ambiguous.', 'apply-failed');
            return created.length
              ? [
                  {
                    cellId: created[0]!,
                    blocks: entry.blocks,
                    lists: cellListBindings(intent.lists, index),
                    properties: cellPropertiesAt(intent.cellProperties, index),
                  },
                ]
              : [];
          });
          return fillNewCells(withPart(working, part), partName, cells, listContext);
        });
        break;
      }
      case 'split-cell': {
        const intent = edit.intent;
        if (intent.payloads.length)
          refuse(
            'Registered blocks in split cells need the D29 fragment executor.',
            'unsupported-operation'
          );
        const size = nativeVisualSize(preview, intent.tableId, intent.cellId);
        if (!size) refuse('The planned merged cell no longer exists.', 'stale-revision');
        apply({
          op: 'splitTableCell',
          tableId: intent.tableId,
          cellId: intent.cellId,
          rows: intent.rows,
          columns: intent.columns,
          columnDistribution: 'existing-grid',
        });
        deferred.push((working) => {
          const part = working.parts.get(partName)!,
            cells = splitResultCells(part, intent.tableId, intent.cellId, size!);
          if (!cells || cells.length !== intent.cells.length)
            refuse('The split native cells do not match the planned cells.', 'apply-failed');
          return fillNewCells(
            working,
            partName,
            cells!.map((cellId, index) => ({
              cellId,
              blocks: intent.cells[index]!,
              properties: cellPropertiesAt(intent.cellProperties, index),
            }))
          );
        });
        break;
      }
      case 'payload': {
        const store = readPayloadStore(withPart(pkg, preview));
        const entry = store.status === 'ready' ? store.entries.get(edit.blockId) : undefined;
        const base =
          payloads.get(edit.blockId) ??
          (entry?.decoded.status === 'supported' ? entry.decoded.value : null);
        if (!entry || !base) refuse('The planned block payload does not exist.', 'stale-revision');
        let current: unknown = base;
        for (const step of edit.path)
          current =
            current && typeof current === 'object' && Object.hasOwn(current, step as string)
              ? (current as Record<string | number, unknown>)[step as string]
              : undefined;
        const present = edit.path.length === 0 || current !== undefined;
        if (
          present !== edit.before.present ||
          (edit.before.present && jsonDataKey(current as never) !== jsonDataKey(edit.before.value))
        )
          refuse('Planned block properties no longer match their before-values.', 'stale-revision');
        let next = JSON.parse(JSON.stringify(base)) as BlockPayload;
        if (!edit.path.length) {
          if (!edit.after.present)
            refuse('A block payload cannot be removed here.', 'unsupported-operation');
          next = edit.after.value as unknown as BlockPayload;
        } else {
          let holder = next as unknown as Record<string, unknown>;
          for (const step of edit.path.slice(0, -1)) {
            const child = holder[step as string];
            if (!child || typeof child !== 'object')
              refuse('The planned payload path does not exist.', 'stale-revision');
            holder = child as Record<string, unknown>;
          }
          const key = String(edit.path.at(-1));
          if (edit.after.present)
            Object.defineProperty(holder, key, {
              value: edit.after.value,
              enumerable: true,
              writable: true,
              configurable: true,
            });
          else delete holder[key];
        }
        if (next.id !== edit.blockId) refuse('A payload edit cannot change block identity.');
        payloads.set(edit.blockId, next);
        const label = entry!.label;
        deferred.push((working) => {
          const stored = withPayloadRecord(working, next, label);
          if (!stored.ok) refuse(`Payload transaction refused: ${stored.reason}`, 'apply-failed');
          return stored.pkg;
        });
        break;
      }
      default:
        refuse('Unknown plan operation.');
    }
  }
  if (lists.length)
    deferred.push((working, _original, listContext) =>
      applyDirectListAssignments(working, partName, lists, listContext)
    );
  // Code runs name the application character style (N18), whether formatted in place or
  // written into new fragments; its definition lands in the same transaction.
  if (
    staged.some(
      (op) =>
        'op' in op &&
        op.op === 'setRunProperties' &&
        typeof op.characterStyle === 'string' &&
        (op.characterStyle === SCIOBOT_CODE_STYLE_ID ||
          op.characterStyle.startsWith(`${SCIOBOT_CODE_STYLE_ID}__`))
    ) ||
    plan.operations.some(insertsInlineCode)
  )
    deferred.push((working) => {
      const styles = [
        ...inlineCodeCharacterStyles(working),
        ...staged.flatMap((op) =>
          'op' in op &&
          op.op === 'setRunProperties' &&
          typeof op.characterStyle === 'string' &&
          (op.characterStyle === SCIOBOT_CODE_STYLE_ID ||
            op.characterStyle.startsWith(`${SCIOBOT_CODE_STYLE_ID}__`))
            ? [op.characterStyle]
            : []
        ),
      ];
      const styled = withInlineCodeStyle(working, styles);
      if (!styled)
        refuse('The document cannot gain the inline code style.', 'unsupported-operation');
      return styled!;
    });
  return { staged, deferred };
}

function wrapOp(edit: Edit<'link'>, relationshipId: string | undefined): TreeDocOp {
  return {
    op: 'insertHyperlink',
    paragraphId: edit.paragraphId,
    start: edit.start,
    end: edit.end,
    ...(edit.target.kind === 'external' ? { relationshipId } : { anchor: edit.target.anchor }),
    ...(edit.title !== null ? { tooltip: edit.title } : {}),
  };
}
function stagedOps(
  staged: readonly Staged[],
  relate: (url: string) => string | null
): TreeDocOp[] | null {
  const ops: TreeDocOp[] = [];
  for (const entry of staged) {
    if ('linkEdit' in entry) {
      const target = entry.linkEdit.target;
      const relationshipId = target.kind === 'external' ? relate(target.url) : undefined;
      if (relationshipId === null) return null;
      ops.push(wrapOp(entry.linkEdit, relationshipId));
      continue;
    }
    if (!('retarget' in entry)) {
      ops.push(entry);
      continue;
    }
    const link = entry.link;
    const relationshipId = link.url !== undefined ? relate(link.url) : link.relationshipId;
    if (link.url !== undefined && relationshipId === null) return null;
    ops.push({
      op: 'setHyperlinkTarget',
      linkId: link.linkId,
      ...(relationshipId !== undefined && relationshipId !== null ? { relationshipId } : {}),
      ...(link.anchor !== undefined ? { anchor: link.anchor } : {}),
      ...(link.tooltip !== undefined ? { tooltip: link.tooltip } : {}),
    });
  }
  return ops;
}

/** Evaluate native operations and allocated resources in a detached package only. */
function previewLowering(pkg: OoxmlPackage, lowering: Lowering): OoxmlPackage {
  const original = nodeIds(pkg.parts.get(pkg.mainDocumentPart)!);
  let preview = pkg;
  const ops = stagedOps(lowering.staged, (url) => {
    const made = ensureHyperlinkRelationship(preview, url, pkg.mainDocumentPart);
    if (!made) return null;
    preview = made.pkg;
    return made.relationshipId;
  });
  if (!ops) refuse('The native engine refuses this hyperlink target.', 'unsupported-operation');
  let part = preview.parts.get(pkg.mainDocumentPart)!;
  for (const op of ops!) {
    const result = applyTreeOp(part, op);
    if (!result.ok) refuse(`Native operation ${op.op} refused: ${result.reason}.`, 'apply-failed');
    part = (result as { part: OoxmlPart }).part;
  }
  preview = withPart(preview, part);
  const listContext = createDirectListContext();
  for (const step of lowering.deferred) preview = step(preview, original, listContext);
  return preview;
}

/** Internal target read candidate. Never install this package into a live editor. */
export function previewDirectPlan(
  pkg: OoxmlPackage,
  plan: CanonicalEditPlan
): { readonly success: true; readonly pkg: OoxmlPackage } | NativeEditFailure {
  try {
    return { success: true, pkg: previewLowering(pkg, lowerDirectPlan(pkg, plan)) };
  } catch (error) {
    return failure(error);
  }
}

/**
 * Apply one validated mode-independent plan as a direct AI edit: one native transaction, one undo
 * step, one committed notification, and no review items. Every old value is rechecked first, and
 * a complete immutable preview runs before the live write, so refusals leave the document intact.
 * Suggestion mode is never switched off; direct edits refuse while tracking is on.
 */
export function applyDirectPlan(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  context: DirectPlanContext
): DirectPlanResult {
  if (
    !context ||
    typeof context.documentId !== 'string' ||
    !context.documentId ||
    typeof context.epoch !== 'string' ||
    !context.epoch
  )
    return { success: false, code: 'invalid-input', error: 'Invalid direct edit context.' };
  let plan: CanonicalEditPlan;
  try {
    const decoded = decodeEditPlan(serializeEditPlan(input));
    if (!decoded.success) return decoded;
    plan = decoded.plan;
  } catch (error) {
    return failure(new NativeWriteRefusal(error instanceof Error ? error.message : String(error)));
  }
  const actual = readEditorRevision(editor, context.documentId, context.epoch),
    surface = editor.surface;
  if (!actual || !surface)
    return {
      success: false,
      code: 'document-not-ready',
      error: 'The native editor is not attached.',
    };
  if (!sameRevision(actual, plan.revision))
    return {
      success: false,
      code: 'stale-revision',
      error: 'Read the current document before applying this edit.',
      expectedRevision: plan.revision,
      actualRevision: actual,
    };
  const mode = surface.editingMode();
  if (mode === 'view')
    return {
      success: false,
      code: 'unsupported-operation',
      error: 'The document is open for viewing.',
    };
  if (mode !== 'edit')
    return {
      success: false,
      code: 'unsupported-operation',
      error:
        'Direct AI edits need editing mode. Suggestion mode keeps tracking on; use a suggestion.',
    };
  if (!plan.operations.length) return { success: true, revision: actual, committed: false };
  const session = surface.session,
    pkg = session.currentPackage(),
    original = nodeIds(session.part());
  let lowering: Lowering;
  try {
    lowering = lowerDirectPlan(pkg, plan);
    previewLowering(pkg, lowering);
  } catch (error) {
    return failure(error);
  }
  const complete = lowering.deferred.length
    ? [
        (working: OoxmlPackage) => {
          const listContext = createDirectListContext();
          return lowering.deferred.reduce(
            (next, step) => step(next, original, listContext),
            working
          );
        },
      ]
    : undefined;
  try {
    const result = surface.applyAutomationOps(
      (relate) => stagedOps(lowering.staged, relate),
      { kind: 'body' },
      complete,
      false,
      true,
      'after'
    );
    if (result.rejected || !result.committed)
      return {
        success: false,
        code: 'apply-failed',
        error: String(result.reason ?? 'The direct edit did not commit.'),
        expectedRevision: actual,
        actualRevision: { ...actual, revision: session.packageRevision() },
      };
  } catch (error) {
    const changed = session.packageRevision() !== actual.revision;
    return {
      success: false,
      code: 'apply-failed',
      error: `${changed ? 'The document changed during a failed native call. Re-read before retrying. ' : ''}${error instanceof Error ? error.message : String(error)}`,
      expectedRevision: actual,
      actualRevision: { ...actual, revision: session.packageRevision() },
    };
  }
  return {
    success: true,
    revision: Object.freeze({ ...actual, revision: session.packageRevision() }),
    committed: true,
  };
}
