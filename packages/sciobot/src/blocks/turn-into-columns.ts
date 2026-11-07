import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { paragraphFragmentsOf } from '@docx-editor.dev/core/layout';
import {
  findNode,
  insertChildren,
  moveNode,
  parentNodeOf,
  removeNode,
  withPart,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { activeMdxRegistry } from '../registry/active';
import { buildBlockIndex } from './block-index';
import { columnWidthsTwips, readColumnGeometry } from './column-geometry';
import { columnContent, columnScaffold, columnStructureRefusal } from './column-structure';
import { readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import { setColumnCount } from './set-column-count';

export interface TurnIntoColumns {
  readonly count: number;
  readonly expectedRevision: number;
}
export type TurnIntoColumnsResult =
  { readonly success: true; readonly revision: number } | NativeEditFailure;
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

/** Wrap the current native paragraph, or change its enclosing column group. */
export function turnIntoColumns(
  editor: DocxEditorInstance,
  request: TurnIntoColumns
): TurnIntoColumnsResult {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (surface.editingMode() !== 'edit')
    return fail('unsupported-operation', 'Column structure changes require editing mode.');
  if (surface.session.packageRevision() !== request.expectedRevision)
    return fail('stale-revision', 'Read the current selection before creating columns.');
  if (!Number.isInteger(request.count) || request.count < 1 || request.count > 63)
    return fail('invalid-input', 'The column count must be between 1 and 63.');
  const { anchor, head } = surface.state().selection;
  if (anchor.paragraphId !== head.paragraphId || surface.storyScope().kind !== 'body')
    return fail('unsupported-operation', 'Choose one native body paragraph for this action.');
  const pkg = surface.session.currentPackage(),
    part = surface.session.part();
  // Shared infrastructure belongs to room bootstrap, never to one actor's undo history.
  if (surface.collaborationSession() && readPayloadStore(pkg).status !== 'ready')
    return fail(
      'document-not-ready',
      'Prepare the native document before starting shared column editing.'
    );
  const paragraph = findNode(part, head.paragraphId);
  const parent = paragraph && parentNodeOf(part, paragraph.id);
  if (paragraph?.kind !== 'paragraph' || !parent)
    return fail('unsupported-operation', 'The selected paragraph is unavailable.');
  const index = buildBlockIndex(pkg);
  const controls = new Map([...index.byId.values()].map((block) => [block.controlId, block]));
  for (let node: OoxmlNode | null = parent; node; node = parentNodeOf(part, node.id)) {
    const block = controls.get(node.id);
    if (!block) continue;
    if (block.payload.type === 'column_group')
      return setColumnCount(editor, { ...request, groupId: block.id });
    if (activeMdxRegistry().components.get(block.payload.type)?.content !== 'native')
      return fail('unsupported-operation', 'This application block owns its fallback content.');
  }
  const refusal = columnStructureRefusal(part, paragraph.id);
  if (refusal) return fail('unsupported-operation', refusal);
  const fragment = surface
    .layout()
    .pages.flatMap((page) => [...paragraphFragmentsOf(page)])
    .find((entry) => entry.paragraphId === paragraph.id);
  if (!fragment)
    return fail('unsupported-operation', 'The selected paragraph has no native layout.');
  try {
    const widths = columnWidthsTwips(
      Array(request.count).fill(`${100 / request.count}%`),
      Math.round(fragment.box.width * 20)
    );
    const total = widths.reduce((sum, value) => sum + value, 0);
    const groupId = `columns-${crypto.randomUUID()}`;
    const ids = widths.map(() => `column-${crypto.randomUUID()}`);
    const change = (working: OoxmlPackage) => {
      const current = working.parts.get(part.name)!;
      const scaffold = columnScaffold(current, widths, ids, groupId);
      const at = parent.children.findIndex((node) => node.id === paragraph.id);
      const inserted = insertChildren(current, parent.id, at, [scaffold.node]);
      if (!inserted.ok) throw Error('The column group could not be inserted.');
      const control = scaffold.cells[0]!.children.find((node) => node.kind === 'contentControl')!;
      const content = columnContent(inserted.part, control.id);
      const moved = moveNode(inserted.part, paragraph.id, content.id, 0);
      if (!moved.ok) throw Error('The selected paragraph could not move into its column.');
      const removed = removeNode(moved.part, content.children[0]!.id);
      if (!removed.ok) throw Error('The unused column placeholder could not be removed.');
      let next = withPart(working, removed.part);
      for (const [i, id] of [groupId, ...ids].entries()) {
        const saved = withPayloadRecord(
          next,
          {
            id,
            version: 1,
            type: i === 0 ? 'column_group' : 'column',
            attributes:
              i === 0 ? {} : { width: `${Number(((widths[i - 1]! * 100) / total).toFixed(10))}%` },
          },
          i === 0 ? 'column_group' : 'column'
        );
        if (!saved.ok) throw Error(saved.reason);
        next = saved.pkg;
      }
      const group = buildBlockIndex(next).byId.get(groupId);
      if (!group || !readColumnGeometry(next, group))
        throw Error('The new column geometry is invalid.');
      return next;
    };
    const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [change], false, true);
    if (result.rejected || !result.committed)
      return fail('apply-failed', String(result.reason ?? 'The column group did not commit.'));
    const point = { paragraphId: paragraph.id, offset: 0 };
    surface.setSelection({ anchor: point, head: point });
    surface.focus();
    return { success: true, revision: surface.session.packageRevision() };
  } catch (error) {
    return fail('apply-failed', error instanceof Error ? error.message : 'Column creation failed.');
  }
}
