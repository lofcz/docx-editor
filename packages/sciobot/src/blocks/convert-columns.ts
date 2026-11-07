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
  removeNode,
  usedParaIds,
  validateTreeOp,
  w14RootPrefix,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlPackage,
  type OoxmlNode,
  type OoxmlPart,
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
  { readonly success: true; readonly revision: number } | NativeEditFailure;
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

function retainParagraphs(part: OoxmlPart, sourceId: string, controlId: string): OoxmlPart {
  const source = findNode(part, sourceId)!;
  const target = contentControlContentNodeOf(findNode(part, controlId)!)!;
  if (source.kind === 'textValue') throw new Error('The paragraph owner is invalid.');
  let at = target.children.length;
  for (const child of source.children) {
    if (child.kind !== 'paragraph') continue;
    const moved = moveNode(part, child.id, target.id, at++);
    if (!moved.ok) throw new Error('A trailing paragraph could not move into its column.');
    part = moved.part;
  }
  return part;
}

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
  const native = group && readColumnGeometry(pkg, group);
  if (native && !native.separatorIds.size)
    return { success: true, revision: request.expectedRevision };
  const legacy = readLegacyColumnGroup(pkg, request.groupId);
  const owner = legacy?.group ?? native?.group;
  const columns = legacy?.columns ?? native?.columns.map((column) => column.block);
  if (!owner || !columns || owner.partName !== part.name || surface.storyScope().kind !== 'body')
    return fail(
      'unsupported-operation',
      'The group has no complete column representation in the active body.'
    );
  for (const column of columns) {
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
  try {
    const commit = (
      change: (working: OoxmlPackage) => OoxmlPackage
    ): ColumnGroupConversionResult => {
      const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [change], false, true);
      if (result.rejected || !result.committed)
        return fail(
          'apply-failed',
          String(result.reason ?? 'The column conversion did not commit.')
        );
      return { success: true, revision: surface.session.packageRevision() };
    };
    if (native)
      return commit((working) => {
        let current = working.parts.get(part.name)!;
        // Older native prototypes have cell and group paragraphs outside the controls.
        // Move them in reading order. Their identities, formatting, and undo remain intact.
        for (const column of native.columns)
          current = retainParagraphs(current, column.cellId, column.block.controlId);
        const content = contentControlContentNodeOf(findNode(current, owner.controlId)!)!;
        current = retainParagraphs(current, content.id, columns.at(-1)!.controlId);
        return withPart(working, current);
      });
    if (!legacy) throw new Error('The stacked column representation is unavailable.');
    const fragment = surface
      .layout()
      .contentControls?.find((control) => control.id === legacy.group.controlId)?.fragments[0];
    if (!fragment)
      return fail('unsupported-operation', 'The column group has no native layout geometry.');
    const widths = columnWidthsTwips(
      legacy.columns.map((column) => column.payload.attributes.width),
      Math.round(fragment.box.width * 20)
    );
    // Temporary cells receive the existing controls before the transaction commits.
    const xml = columnTableXml(widths, () => '<w:p/>');
    const parsed = readOoxmlPart(
      `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body>${xml}</w:body></w:document>`,
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
      const template = body.children[0]!;
      // Each existing column supplies the cell's required block content.
      const table = withFreshIds(template, allocate, identity);
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
        const seed = cells[index]!.children.find((node) => node.kind === 'paragraph')!;
        const removed = removeNode(current, seed.id);
        if (!removed.ok) throw new Error('The temporary cell seed could not be removed.');
        current = removed.part;
      }
      // Retain imported trailing paragraph identities inside the final column's ownership.
      if (legacy.hasTrailingParagraph)
        current = retainParagraphs(current, legacy.content.id, legacy.columns.at(-1)!.controlId);
      const converted = withPart(working, current);
      if (!readColumnGeometry(converted, legacy.group))
        throw new Error('The converted column structure is invalid.');
      return converted;
    };
    return commit(convert);
  } catch (error) {
    return fail(
      'apply-failed',
      error instanceof Error ? error.message : 'The column conversion failed.'
    );
  }
}
