import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  bodyStoryRoot,
  contentControlContentNodeOf,
  contentControlLevelOf,
  contentControlPropertiesOf,
  findNode,
  moveNode,
  parentNodeOf,
  withPart,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { buildBlockIndex } from './block-index';

export interface BlockMove {
  readonly controlId: string;
  /** Registered container, or null for the document body. */
  readonly parentControlId: string | null;
  /** A direct native child to insert before; null appends before final section properties. */
  readonly beforeId: string | null;
  readonly expectedRevision?: number;
}
export type BlockMoveResult =
  | { readonly success: true; readonly revision: number; readonly changed: boolean }
  | NativeEditFailure;
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

/** Move a registered native container without touching payload records or descendant IDs. */
export function moveBlock(editor: DocxEditorInstance, request: BlockMove): BlockMoveResult {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  const revision = surface.session.packageRevision();
  if (request.expectedRevision !== undefined && request.expectedRevision !== revision)
    return fail('stale-revision', 'Read the current document before moving this block.');
  if (surface.editingMode() === 'suggest')
    return fail('unsupported-operation', 'Tracked block movement requires a review transaction.');
  const pkg = surface.session.currentPackage(),
    part = surface.session.part();
  const registered = new Map(
    [...buildBlockIndex(pkg).byId.values()]
      .filter((block) => block.partName === part.name)
      .map((block) => [block.controlId, block])
  );
  const source = findNode(part, request.controlId);
  if (!registered.has(request.controlId) || !source || source.kind !== 'contentControl')
    return fail('invalid-input', 'The registered source block is missing.');
  if (!['block', 'empty'].includes(contentControlLevelOf(source)))
    return fail('invalid-input', 'Only native container blocks move between containers.');
  // Use the native wrapper and ancestor lock policy for the existence operation.
  const removal = applyTreeOp(part, {
    op: 'removeContentControl',
    controlId: source.id,
    keepContent: true,
  });
  if (!removal.ok) return fail('apply-failed', `Native block movement refused: ${removal.reason}`);
  let root: OoxmlElement = source;
  // A structural insertion proposal commonly owns one complete control. Carry its wrapper
  // with the control, rather than silently accepting that pending insertion by moving out.
  for (
    let parent = parentNodeOf(part, root.id);
    parent &&
    ['revisionInsert', 'revisionDelete', 'revisionMoveFrom', 'revisionMoveTo'].includes(
      parent.kind
    );
    parent = parentNodeOf(part, root.id)
  ) {
    if (parent.children.length !== 1)
      return fail(
        'unsupported-operation',
        'Move the complete pending insertion group through its review transaction.'
      );
    root = parent;
  }
  let destination: OoxmlNode | null = bodyStoryRoot(part);
  if (request.parentControlId !== null) {
    const parent = findNode(part, request.parentControlId);
    if (!registered.has(request.parentControlId) || !parent || parent.kind !== 'contentControl')
      return fail('invalid-input', 'The registered destination container is missing.');
    if (!['block', 'empty'].includes(contentControlLevelOf(parent)))
      return fail('invalid-input', 'The destination is not a block container.');
    for (
      let ancestor: OoxmlNode | null = parent;
      ancestor;
      ancestor = parentNodeOf(part, ancestor.id)
    ) {
      if (ancestor.id === root.id || ancestor.id === source.id)
        return fail('invalid-input', 'A block cannot move inside itself or its descendants.');
      if (ancestor.kind !== 'contentControl') continue;
      const properties = contentControlPropertiesOf(ancestor);
      if (
        properties.lock === 'contentLocked' ||
        properties.lock === 'sdtContentLocked' ||
        properties.dataBinding
      )
        return fail('apply-failed', 'The destination content is protected.');
    }
    destination = contentControlContentNodeOf(parent) ?? null;
  }
  if (!destination || destination.kind === 'textValue')
    return fail('invalid-document', 'The destination has no native content container.');
  if (destination.children.some((child) => child.kind === 'run' || child.kind === 'textValue'))
    return fail('invalid-input', 'The destination is an inline control.');
  const siblings = destination.children.filter((child) => child.id !== root.id);
  let index =
    request.beforeId === null
      ? siblings.findIndex((child) => child.kind !== 'textValue' && child.localName === 'sectPr')
      : siblings.findIndex((child) => child.id === request.beforeId);
  if (
    request.beforeId !== null &&
    request.beforeId === root.id &&
    parentNodeOf(part, root.id)?.id === destination.id
  )
    index = destination.children.findIndex((child) => child.id === root.id);
  else if (request.beforeId !== null && index < 0)
    return fail('invalid-input', 'The movement anchor is not a destination child.');
  if (index < 0) index = siblings.length;
  const prepared = moveNode(part, root.id, destination.id, index);
  if (!prepared.ok)
    return fail('invalid-input', 'The native movement would create an invalid document.');
  if (prepared.part === part) return { success: true, revision, changed: false };
  const destinationId = destination.id;
  const move = (working: OoxmlPackage) => {
    const current = working.parts.get(part.name);
    if (!current) throw new Error('The document part disappeared.');
    const moved = moveNode(current, root.id, destinationId, index);
    if (!moved.ok) throw new Error('The native block movement failed.');
    return withPart(working, moved.part);
  };
  try {
    const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [move], false, true);
    if (result.rejected || !result.committed)
      return fail('apply-failed', String(result.reason ?? 'The block movement did not commit.'));
    return { success: true, revision: surface.session.packageRevision(), changed: true };
  } catch (error) {
    return fail(
      'apply-failed',
      error instanceof Error ? error.message : 'The block movement failed.'
    );
  }
}
