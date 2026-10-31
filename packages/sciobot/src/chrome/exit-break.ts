import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  parentNodeOf,
  readOoxmlPart,
  WML_NAMESPACE_URI,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';

export type ExitDirection = 'after' | 'before';
export type ExitBlockResult =
  | { readonly success: true; readonly paragraphId: string }
  | NativeEditFailure;

const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});

/**
 * Leave every enclosing structure at the caret (table, application block, container) and put
 * the caret in a new empty body paragraph after or before the outermost one, in one undo unit.
 * At a plain body paragraph this adds a paragraph next to it. This is the production exit break
 * (Mod+Enter and Mod+Shift+Enter).
 */
export function exitBlock(editor: DocxEditorInstance, direction: ExitDirection): ExitBlockResult {
  const surface = editor.surface;
  if (!surface) return fail('The native editor is not attached.');
  surface.flushPendingInput();
  if (surface.editingMode() !== 'edit')
    return fail('Exit break inserts a paragraph only in editing mode.');
  const session = surface.session,
    pkg = session.currentPackage(),
    part = session.part();
  if (part.name !== pkg.mainDocumentPart) return fail('Exit break requires the body story.');
  const body = bodyStoryRoot(part);
  if (!body || body.kind === 'textValue') return fail('The document has no body.');
  let top = surface.state().selection.head.paragraphId;
  for (let parent = parentNodeOf(part, top); parent && parent.id !== body.id; ) {
    top = parent.id;
    parent = parentNodeOf(part, top);
  }
  const index = body.children.findIndex((child) => child.id === top);
  if (index < 0) return fail('The caret is outside the document body.');
  const beforeId = direction === 'before' ? top : (body.children[index + 1]?.id ?? null);
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body><w:p/></w:body></w:document>`,
    { name: part.name, contentType: part.contentType }
  );
  const seedBody = parsed.ok ? bodyStoryRoot(parsed.part) : null;
  const paragraph = seedBody && seedBody.kind !== 'textValue' ? seedBody.children[0] : null;
  if (!paragraph || paragraph.kind !== 'paragraph') return fail('The paragraph seed is invalid.');
  const before = new Set(session.paragraphIds());
  const result = surface.applyAutomationOps(
    () => [{ op: 'insertParagraphs', parentId: body.id, beforeId, paragraphs: [paragraph] }],
    { kind: 'body' }
  );
  if (!result.committed || result.rejected)
    return fail(String(result.reason ?? 'The paragraph insertion did not commit.'));
  const id = session.paragraphIds().find((paragraphId) => !before.has(paragraphId));
  if (!id) return fail('The inserted paragraph was not allocated.');
  const point = { paragraphId: id, offset: 0 };
  surface.setSelection({ anchor: point, head: point });
  return { success: true, paragraphId: id };
}

/** Mod+Enter exits after, Mod+Shift+Enter before. Widget forms keep their own keys. */
export function installExitBreak(editor: DocxEditorInstance): () => void {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const host = surface.hostElement();
  const keydown = (event: KeyboardEvent) => {
    const target = event.target as Element | null;
    if (
      event.key !== 'Enter' ||
      !(event.ctrlKey || event.metaKey) ||
      event.altKey ||
      event.isComposing ||
      !target?.closest?.('[contenteditable="true"]') ||
      target.closest('[data-sciobot-block-ui]')
    )
      return;
    event.preventDefault();
    event.stopImmediatePropagation();
    exitBlock(editor, event.shiftKey ? 'before' : 'after');
  };
  host.addEventListener('keydown', keydown, true);
  return () => host.removeEventListener('keydown', keydown, true);
}
