import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  readOoxmlPart,
  storyParagraphs,
  WML_NAMESPACE_URI,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { buildBlockIndex } from './block-index';
import { readColumnGeometry } from './column-geometry';
import { columnContent } from './column-structure';

export interface ColumnEntry {
  readonly columnId: string;
  readonly expectedRevision: number;
}
export type ColumnEntryResult =
  | { readonly success: true; readonly paragraphId: string; readonly revision: number }
  | NativeEditFailure;
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

/** Enter an existing native paragraph, or explicitly create one in an empty column. */
export function enterColumn(editor: DocxEditorInstance, request: ColumnEntry): ColumnEntryResult {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (surface.editingMode() !== 'edit')
    return fail('unsupported-operation', 'Empty column entry requires editing mode.');
  if (surface.session.packageRevision() !== request.expectedRevision)
    return fail('stale-revision', 'Read the current column before entering it.');
  const pkg = surface.session.currentPackage(),
    part = surface.session.part();
  const index = buildBlockIndex(pkg),
    column = index.byId.get(request.columnId);
  const group = column?.parentId ? index.byId.get(column.parentId) : undefined;
  const geometry = group && readColumnGeometry(pkg, group, index);
  if (
    !column ||
    !geometry?.columns.some((entry) => entry.block.id === column.id) ||
    column.partName !== part.name ||
    surface.storyScope().kind !== 'body'
  )
    return fail('unsupported-operation', 'The native column is unavailable.');
  try {
    const content = columnContent(part, column.controlId);
    let paragraph = storyParagraphs(content)[0];
    if (!paragraph) {
      const parsed = readOoxmlPart(
        `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body><w:p/></w:body></w:document>`,
        { name: part.name, contentType: part.contentType }
      );
      const body = parsed.ok ? bodyStoryRoot(parsed.part) : null;
      const seed = body && body.kind !== 'textValue' ? body.children[0] : null;
      if (seed?.kind !== 'paragraph') throw Error('The empty column seed is invalid.');
      const result = surface.applyAutomationOps(
        () => [
          { op: 'insertParagraphs', parentId: content.id, beforeId: null, paragraphs: [seed] },
        ],
        { kind: 'body' },
        [],
        false,
        true
      );
      if (result.rejected || !result.committed)
        return fail('apply-failed', String(result.reason ?? 'The column entry did not commit.'));
      paragraph = storyParagraphs(columnContent(surface.session.part(), column.controlId))[0];
    }
    if (!paragraph) throw Error('The column has no editable paragraph.');
    const point = { paragraphId: paragraph.id, offset: 0 };
    surface.setSelection({ anchor: point, head: point });
    surface.focus();
    return {
      success: true,
      paragraphId: paragraph.id,
      revision: surface.session.packageRevision(),
    };
  } catch (error) {
    return fail(
      'apply-failed',
      error instanceof Error ? error.message : 'The column entry failed.'
    );
  }
}
