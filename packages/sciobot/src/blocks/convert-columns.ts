import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  contentControlContentNodeOf,
  contentControlPropertiesOf,
  createNodeIdAllocator,
  findNode,
  insertChildren,
  moveNode,
  parentNodeOf,
  paragraphTextOf,
  readOoxmlPart,
  usedParaIds,
  validateTreeOp,
  w14RootPrefix,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { buildBlockIndex, type IndexedBlock } from './block-index';
import { columnTableXml, columnWidthsTwips, readColumnGeometry } from './column-geometry';
import { withFreshIds } from './fresh-ids';

/** Only the former, complete stacked representation qualifies for explicit conversion. */
export function readLegacyColumnGroup(pkg: OoxmlPackage, groupId: string) {
  const index = buildBlockIndex(pkg),
    group = index.byId.get(groupId);
  if (group?.payload.type !== 'column_group') return null;
  const part = pkg.parts.get(group.partName),
    control = part && findNode(part, group.controlId);
  const content = control && contentControlContentNodeOf(control);
  if (!part || !content) return null;
  const byControl = new Map([...index.byId.values()].map((block) => [block.controlId, block]));
  const columns: IndexedBlock[] = [];
  let trailing = false;
  for (const child of content.children) {
    const column = byControl.get(child.id);
    if (!trailing && column?.payload.type === 'column' && column.parentId === group.id) {
      columns.push(column);
      continue;
    }
    if (
      child.kind !== 'paragraph' ||
      paragraphTextOf(part, child.id) !== '' ||
      child.children.some(
        (node) => node.namespaceUri !== WML_NAMESPACE_URI || node.localName !== 'pPr'
      )
    )
      return null;
    trailing = true;
  }
  if (!columns.length || columns.length > 63) return null;
  return { group, content, columns, hasTrailingParagraph: trailing };
}

export interface ColumnGroupConversion {
  readonly groupId: string;
  readonly expectedRevision: number;
}
export type ColumnGroupConversionResult =
  | { readonly success: true; readonly revision: number }
  | NativeEditFailure;
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

/**
 * Convert stacked columns before joining collaboration. Existing control and text identities survive.
 * Shared sessions, including disconnected sessions, refuse: concurrent wrapper creation needs coordination.
 */
export function convertColumnGroup(
  editor: DocxEditorInstance,
  request: ColumnGroupConversion
): ColumnGroupConversionResult {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (surface.editingMode() !== 'edit')
    return fail('unsupported-operation', 'Column conversion requires editing mode.');
  if (surface.session.packageRevision() !== request.expectedRevision)
    return fail('stale-revision', 'Read the current columns before converting their layout.');
  if (surface.collaborationSession())
    return fail(
      'unsupported-operation',
      'Convert stacked columns before joining a collaboration session.'
    );
  const pkg = surface.session.currentPackage(),
    part = surface.session.part();
  const group = buildBlockIndex(pkg).byId.get(request.groupId);
  if (group && readColumnGeometry(pkg, group))
    return { success: true, revision: request.expectedRevision };
  const legacy = readLegacyColumnGroup(pkg, request.groupId);
  if (!legacy || legacy.group.partName !== part.name || surface.storyScope().kind !== 'body')
    return fail(
      'unsupported-operation',
      'The group has no complete stacked column representation in the active body.'
    );
  for (const column of legacy.columns) {
    const reason = validateTreeOp(part, {
      op: 'removeContentControl',
      controlId: column.controlId,
      keepContent: true,
    });
    if (reason) return fail('unsupported-operation', `Column conversion refused: ${reason}.`);
    for (
      let node: OoxmlNode | null = findNode(part, column.controlId);
      node;
      node = parentNodeOf(part, node.id)
    ) {
      if (node.kind !== 'contentControl') continue;
      const properties = contentControlPropertiesOf(node);
      if (
        properties.dataBinding ||
        (properties.lock !== undefined && properties.lock !== 'unlocked')
      )
        return fail('unsupported-operation', 'The column group contains protected content.');
    }
  }
  const fragment = surface
    .layout()
    .contentControls?.find((control) => control.id === legacy.group.controlId)?.fragments[0];
  if (!fragment)
    return fail('unsupported-operation', 'The column group has no native layout geometry.');
  try {
    const widths = columnWidthsTwips(
      legacy.columns.map((column) => column.payload.attributes.width),
      Math.round(fragment.box.width * 20)
    );
    const xml = columnTableXml(
      widths,
      () => '',
      () => '<w:p/>'
    );
    const parsed = readOoxmlPart(
      `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body>${xml}<w:p/></w:body></w:document>`,
      { name: '/column-conversion.xml', contentType: part.contentType }
    );
    const body = parsed.ok ? bodyStoryRoot(parsed.part) : null;
    if (!body || body.kind === 'textValue')
      return fail('invalid-input', 'The native column template is invalid.');
    const identitySeed = `column-conversion:${request.groupId}:${crypto.randomUUID()}`;
    const convert = (working: OoxmlPackage) => {
      let current = working.parts.get(part.name)!;
      const allocate = createNodeIdAllocator(current, 'new');
      const identity = {
        prefix: w14RootPrefix(current.root),
        used: new Set(usedParaIds(current.root)),
        seed: identitySeed,
      };
      const table = withFreshIds(body.children[0]!, allocate, identity);
      if (table.kind !== 'table') throw new Error('The native template contains no table.');
      const row = table.children.find((node) => node.kind === 'tableRow');
      const cells = row?.children.filter((node) => node.kind === 'tableCell') ?? [];
      if (cells.length !== legacy.columns.length)
        throw new Error('The native column count differs.');
      const inserted = insertChildren(current, legacy.content.id, 0, [table]);
      if (!inserted.ok) throw new Error('The native column table could not be inserted.');
      current = inserted.part;
      for (const [index, column] of legacy.columns.entries()) {
        const moved = moveNode(current, column.controlId, cells[index]!.id, 1);
        if (!moved.ok) throw new Error('A column control could not move into its native cell.');
        current = moved.part;
      }
      if (!legacy.hasTrailingParagraph) {
        const trailing = insertChildren(current, legacy.content.id, 1, [
          withFreshIds(body.children[1]!, allocate, identity),
        ]);
        if (!trailing.ok) throw new Error('The native column group needs its terminal paragraph.');
        current = trailing.part;
      }
      const converted = withPart(working, current);
      if (!readColumnGeometry(converted, legacy.group))
        throw new Error('The converted column structure is invalid.');
      return converted;
    };
    const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [convert], false, true);
    if (result.rejected || !result.committed)
      return fail('apply-failed', String(result.reason ?? 'The column conversion did not commit.'));
    return { success: true, revision: surface.session.packageRevision() };
  } catch (error) {
    return fail(
      'apply-failed',
      error instanceof Error ? error.message : 'The column conversion failed.'
    );
  }
}
