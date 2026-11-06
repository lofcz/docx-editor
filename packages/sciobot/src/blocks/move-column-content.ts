import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  bodyStoryRoot,
  contentControlContentNodeOf,
  contentControlPropertiesOf,
  findNode,
  moveNode,
  parentNodeOf,
  readOoxmlPart,
  validateTreeOp,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlParagraphNode,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { buildBlockIndex } from './block-index';
import { readColumnGeometry } from './column-geometry';
import { blockMovementRangeRefusal } from './selection-move-ranges';

export interface ColumnContentMove {
  readonly groupId: string;
  /** One complete direct child of a column, with its existing native identity. */
  readonly nodeId: string;
  readonly destinationColumnId: string;
  /** A direct destination child; null appends. */
  readonly beforeId: string | null;
  readonly expectedRevision: number;
}
export type ColumnContentMoveResult =
  | { readonly success: true; readonly revision: number }
  | NativeEditFailure;
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

function emptySeed(): OoxmlParagraphNode {
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body><w:p/></w:body></w:document>`,
    {
      name: '/column-seed.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!parsed.ok) throw new Error('The empty column seed is invalid.');
  const body = bodyStoryRoot(parsed.part);
  const seed = body && body.kind !== 'textValue' ? body.children[0] : null;
  if (seed?.kind !== 'paragraph') throw new Error('The empty column seed has no paragraph.');
  return seed;
}

/** Move a native block across sibling columns. Its text, formatting, payloads, and identities survive. */
export function moveColumnContent(
  editor: DocxEditorInstance,
  request: ColumnContentMove
): ColumnContentMoveResult {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (surface.editingMode() !== 'edit')
    return fail(
      'unsupported-operation',
      'Column content movement is unavailable in this editing mode.'
    );
  if (request.expectedRevision !== surface.session.packageRevision())
    return fail('stale-revision', 'Read the current columns before moving their content.');
  const pkg = surface.session.currentPackage(),
    part = surface.session.part();
  const index = buildBlockIndex(pkg),
    group = index.byId.get(request.groupId);
  const geometry = group && readColumnGeometry(pkg, group, index);
  if (!geometry || group!.partName !== part.name)
    return fail('unsupported-operation', 'The column group has no editable native geometry.');
  const node = findNode(part, request.nodeId);
  if (!node || !['paragraph', 'table', 'contentControl'].includes(node.kind))
    return fail('invalid-input', 'Choose one complete native column block.');
  const sourceParent = parentNodeOf(part, node.id);
  const contentOf = (controlId: string) => {
    const control = findNode(part, controlId);
    return control ? contentControlContentNodeOf(control) : undefined;
  };
  const source = geometry.columns.find(
    (column) => contentOf(column.block.controlId)?.id === sourceParent?.id
  );
  const destination = geometry.columns.find(
    (column) => column.block.id === request.destinationColumnId
  );
  const target = destination && contentOf(destination.block.controlId);
  if (
    !source ||
    !sourceParent ||
    !destination ||
    !target ||
    source.block.id === destination.block.id
  )
    return fail('invalid-input', 'Choose another column in the same group.');
  const at =
    request.beforeId === null
      ? target.children.length
      : target.children.findIndex((child) => child.id === request.beforeId);
  if (at < 0)
    return fail('invalid-input', 'The destination anchor no longer belongs to that column.');
  const removal = validateTreeOp(
    part,
    node.kind === 'contentControl'
      ? { op: 'removeContentControl', controlId: node.id, keepContent: true }
      : { op: 'deleteBlock', blockId: node.id }
  );
  if (removal) return fail('unsupported-operation', `Column content movement refused: ${removal}.`);
  for (
    let ancestor: OoxmlNode | null = target;
    ancestor;
    ancestor = parentNodeOf(part, ancestor.id)
  ) {
    if (ancestor.kind !== 'contentControl') continue;
    const properties = contentControlPropertiesOf(ancestor);
    if (
      properties.dataBinding ||
      properties.lock === 'contentLocked' ||
      properties.lock === 'sdtContentLocked'
    )
      return fail('unsupported-operation', 'The destination column is protected.');
  }
  const sourceContent = contentOf(source.block.controlId)!;
  const remaining = sourceContent.children.filter((child) => child.id !== node.id);
  // Keep an editable paragraph in an empty column and after its final nested table.
  const needsSeed = !remaining.length || remaining.at(-1)?.kind === 'table';
  const change = (working: OoxmlPackage) => {
    let current = working.parts.get(part.name)!;
    if (needsSeed) {
      const seeded = applyTreeOp(current, {
        op: 'insertParagraphs',
        parentId: sourceContent.id,
        beforeId: null,
        paragraphs: [emptySeed()],
      });
      if (!seeded.ok) throw new Error(`Column seed insertion refused: ${seeded.reason}.`);
      current = seeded.part;
    }
    const moved = moveNode(current, node.id, target.id, at);
    if (!moved.ok) throw new Error('Native column content movement failed.');
    current = moved.part;
    if (node.kind === 'table' && request.beforeId === null) {
      const seeded = applyTreeOp(current, {
        op: 'insertParagraphs',
        parentId: target.id,
        beforeId: null,
        paragraphs: [emptySeed()],
      });
      if (!seeded.ok) throw new Error(`Column seed insertion refused: ${seeded.reason}.`);
      current = seeded.part;
    }
    return withPart(working, current);
  };
  try {
    const prepared = change(pkg);
    const rangeRefusal = blockMovementRangeRefusal(part, prepared.parts.get(part.name)!);
    if (rangeRefusal) return fail('unsupported-operation', rangeRefusal);
    const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [change], false, true);
    if (result.rejected || !result.committed)
      return fail(
        'apply-failed',
        String(result.reason ?? 'Column content movement did not commit.')
      );
    return { success: true, revision: surface.session.packageRevision() };
  } catch (error) {
    return fail(
      'apply-failed',
      error instanceof Error ? error.message : 'Column content movement failed.'
    );
  }
}
