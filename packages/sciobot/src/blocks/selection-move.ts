import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  contentControlLevelOf,
  findNode,
  moveNode,
  parentNodeOf,
  validateTreeOp,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { ancestorContentWriteRefusal } from './ancestor-content-write';
import { resolveExactBlockSelection } from './selection-exact';
import { blockMovementRangeRefusal } from './selection-move-ranges';

export interface BlockSelectionMove {
  readonly nodeIds: readonly string[];
  /** A native body, table cell, or block content-control content node in the active body part. */
  readonly parentId: string;
  /** A direct destination child; null appends before final section properties. */
  readonly beforeId: string | null;
  readonly expectedRevision?: number;
}
export type BlockSelectionMoveResult =
  | { readonly success: true; readonly revision: number; readonly changed: boolean }
  | NativeEditFailure;
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

/** Move selected complete blocks in one transaction without cloning text, payloads, or native IDs. */
export function moveBlockSelection(
  editor: DocxEditorInstance,
  request: BlockSelectionMove
): BlockSelectionMoveResult {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  if (surface.activeScope().kind !== 'body')
    return fail('unsupported-operation', 'Block selection changes require the document body.');
  surface.flushPendingInput();
  if (surface.editingMode() !== 'edit')
    return fail('unsupported-operation', 'Block selection movement requires edit mode.');
  const revision = surface.session.packageRevision();
  if (request.expectedRevision !== undefined && request.expectedRevision !== revision)
    return fail('stale-revision', 'Read the current blocks before moving this selection.');
  const pkg = surface.session.currentPackage(),
    part = surface.session.part();
  const range = resolveExactBlockSelection(part, request.nodeIds);
  if (!range.success) return range;
  const destination = findNode(part, request.parentId);
  const blockContainer = (node: OoxmlNode | null): node is OoxmlElement => {
    if (!node || node.kind === 'textValue') return false;
    if (node.kind === 'body' || node.kind === 'tableCell') return true;
    const owner = parentNodeOf(part, node.id);
    return (
      node.namespaceUri === WML_NAMESPACE_URI &&
      node.localName === 'sdtContent' &&
      owner?.kind === 'contentControl' &&
      ['block', 'empty'].includes(contentControlLevelOf(owner))
    );
  };
  if (!blockContainer(destination) || range.owners.some((owner) => !blockContainer(owner.parent)))
    return fail('unsupported-operation', 'Move complete blocks between native block containers.');
  const selected = new Set(range.nodes.map((node) => node.id));
  for (
    let ancestor: OoxmlNode | null = destination;
    ancestor;
    ancestor = parentNodeOf(part, ancestor.id)
  )
    if (selected.has(ancestor.id))
      return fail('invalid-input', 'A block cannot move inside itself or its descendants.');
  const destinationRefusal = ancestorContentWriteRefusal(part, destination.id);
  if (destinationRefusal) return fail('unsupported-operation', destinationRefusal);
  if (
    request.beforeId !== null &&
    !destination.children.some((node) => node.id === request.beforeId)
  )
    return fail('invalid-input', 'The movement anchor is not a destination child.');
  const sameParent = range.owners.length === 1 && destination.id === range.owners[0]!.parent.id;
  const first = destination.children.findIndex((node) => node.id === range.nodes[0]!.id);
  const contiguous =
    sameParent &&
    destination.children
      .slice(first, first + range.nodes.length)
      .every((node) => selected.has(node.id));
  if (request.beforeId !== null && selected.has(request.beforeId))
    return contiguous
      ? { success: true, revision, changed: false }
      : fail('invalid-input', 'Choose a movement anchor outside the selected blocks.');
  if (contiguous) {
    const last = destination.children.findIndex((node) => node.id === range.nodes.at(-1)!.id);
    const next = destination.children[last + 1];
    const end = !next || (next.kind !== 'textValue' && next.localName === 'sectPr');
    if ((request.beforeId === null && end) || request.beforeId === next?.id)
      return { success: true, revision, changed: false };
  }
  const seedTemplate = findNode(part, surface.session.paragraphIds()[0]!);
  const seeds: TreeDocOp[] = [];
  for (const owner of range.owners) {
    if (owner.parent.id === destination.id) continue;
    const remaining = owner.parent.children.filter((node) => !selected.has(node.id));
    const seedRequired =
      !remaining.some((node) => ['paragraph', 'table', 'contentControl'].includes(node.kind)) ||
      (owner.parent.kind === 'tableCell' &&
        (!remaining.some((node) => node.kind === 'paragraph') ||
          remaining.at(-1)?.kind === 'table'));
    if (!seedRequired) continue;
    if (seedTemplate?.kind !== 'paragraph')
      return fail('invalid-document', 'The source has no native paragraph template.');
    seeds.push({
      op: 'insertParagraphs',
      parentId: owner.parent.id,
      beforeId: null,
      paragraphs: [{ ...seedTemplate, children: [] }],
    });
  }
  let preview = part;
  for (const seed of seeds) {
    const inserted = applyTreeOp(preview, seed);
    if (!inserted.ok)
      return fail(
        'unsupported-operation',
        `The source cannot retain an editable paragraph: ${inserted.reason}.`
      );
    preview = inserted.part;
  }
  for (const node of range.nodes) {
    const removal: TreeDocOp =
      node.kind === 'contentControl'
        ? { op: 'removeContentControl', controlId: node.id, keepContent: true }
        : { op: 'deleteBlock', blockId: node.id };
    const refusal = validateTreeOp(preview, removal);
    // Moving within one owner retains the very paragraph that a standalone removal needs.
    const staysInOwner = parentNodeOf(part, node.id)?.id === destination.id;
    if (refusal && !(staysInOwner && refusal === 'block-required'))
      return fail('unsupported-operation', `Native block movement refused: ${refusal}.`);
  }
  const move = (working: OoxmlPackage) => {
    let current = working.parts.get(part.name)!;
    for (const seed of seeds) {
      const inserted = applyTreeOp(current, seed);
      if (!inserted.ok)
        throw Error(`The source paragraph could not be retained: ${inserted.reason}.`);
      current = inserted.part;
    }
    for (const node of range.nodes) {
      const target = findNode(current, destination.id);
      if (!target || target.kind === 'textValue') throw Error('The destination disappeared.');
      const siblings = target.children.filter((child) => child.id !== node.id);
      let at =
        request.beforeId === null
          ? siblings.findIndex(
              (child) =>
                child.kind !== 'textValue' &&
                child.namespaceUri === WML_NAMESPACE_URI &&
                child.localName === 'sectPr'
            )
          : siblings.findIndex((child) => child.id === request.beforeId);
      if (request.beforeId !== null && at < 0) throw Error('The movement anchor disappeared.');
      if (at < 0) at = siblings.length;
      const moved = moveNode(current, node.id, destination.id, at);
      if (!moved.ok) throw Error('The native movement would create an invalid document.');
      current = moved.part;
    }
    const target = findNode(current, destination.id);
    if (target?.kind === 'tableCell' && target.children.at(-1)?.kind === 'table') {
      if (seedTemplate?.kind !== 'paragraph')
        throw Error('The destination has no paragraph template.');
      const inserted = applyTreeOp(current, {
        op: 'insertParagraphs',
        parentId: target.id,
        beforeId: null,
        paragraphs: [{ ...seedTemplate, children: [] }],
      });
      if (!inserted.ok)
        throw Error(`The destination paragraph could not be retained: ${inserted.reason}.`);
      current = inserted.part;
    }
    return current === working.parts.get(part.name) ? working : withPart(working, current);
  };
  try {
    const prepared = move(pkg);
    if (prepared === pkg) return { success: true, revision, changed: false };
    const rangeRefusal = blockMovementRangeRefusal(part, prepared.parts.get(part.name)!);
    if (rangeRefusal) return fail('unsupported-operation', rangeRefusal);
    const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [move], false, true);
    if (result.rejected || !result.committed)
      return fail('apply-failed', String(result.reason ?? 'The movement did not commit.'));
    return { success: true, revision: surface.session.packageRevision(), changed: true };
  } catch (error) {
    return fail('apply-failed', error instanceof Error ? error.message : 'The movement failed.');
  }
}
