import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  bodyStoryRoot,
  directParagraphProperties,
  storyParagraphs,
  withPart,
  parentNodeOf,
  readOoxmlPart,
  WML_NAMESPACE_URI,
} from '@docx-editor.dev/core/store';
import {
  horizontalRuleStyleState,
  HORIZONTAL_RULE_PARAGRAPH_XML,
  SCIOBOT_HORIZONTAL_RULE_STYLE_ID,
} from '../native/horizontal-rule';
import { BLOCK_PARAGRAPH_STYLES, type TextBlockType } from './block-type';
import type { NativeEditFailure } from '../contracts/document';

export type InsertBlockType = TextBlockType | 'hr';
const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});

/** Append a new native block after the selected paragraph without replacing selected text. */
export function insertBlock(
  editor: DocxEditorInstance,
  type: InsertBlockType
): { readonly success: true; readonly paragraphId: string } | NativeEditFailure {
  const surface = editor.surface;
  if (!surface) return fail('The native editor is not attached.');
  surface.flushPendingInput();
  const session = surface.session,
    pkg = session.currentPackage(),
    part = session.part();
  if (part.name !== pkg.mainDocumentPart) return fail('Block controls require the body story.');
  const style = type === 'hr' ? null : BLOCK_PARAGRAPH_STYLES[type];
  if (
    type === 'hr'
      ? horizontalRuleStyleState(pkg) !== 'present'
      : !style ||
        !session
          .documentStyles()
          .some((entry) => entry.type === 'paragraph' && entry.styleId === style)
  )
    return fail('The document does not define this block style.');
  const selection = surface.state().selection;
  const order = session.paragraphIds();
  const a = order.indexOf(selection.anchor.paragraphId),
    h = order.indexOf(selection.head.paragraphId);
  const selected = order[Math.max(a, h)];
  if (!selected || a < 0 || h < 0) return fail('The native selection has no paragraph.');
  const parent = parentNodeOf(part, selected);
  if (!parent) return fail('The paragraph has no native parent.');
  const index = parent.children.findIndex((node) => node.id === selected);
  const beforeId = parent.children[index + 1]?.id ?? null;
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body>${type === 'hr' ? HORIZONTAL_RULE_PARAGRAPH_XML.replace(/<w:pStyle[^>]*\/>/, '') : '<w:p/>'}</w:body></w:document>`,
    { name: part.name, contentType: part.contentType }
  );
  if (!parsed.ok) return fail('The native block seed is invalid.');
  const body = bodyStoryRoot(parsed.part);
  const paragraph = body && body.kind !== 'textValue' ? body.children[0] : null;
  if (!paragraph || paragraph.kind !== 'paragraph')
    return fail('The native block seed has no paragraph.');
  const before = new Set(order);
  const result = surface.applyAutomationOps(
    () => [{ op: 'insertParagraphs', parentId: parent.id, beforeId, paragraphs: [paragraph] }],
    { kind: 'body' },
    [
      (working) => {
        const current = working.parts.get(part.name)!;
        const allocated = storyParagraphs(bodyStoryRoot(current)!).filter(
          (node) => !before.has(node.id)
        );
        if (allocated.length !== 1) throw new Error('Native block allocation is ambiguous.');
        const paragraphId = allocated[0]!.id;
        const properties = directParagraphProperties(current, paragraphId);
        const styled = applyTreeOp(current, {
          op: 'setParagraphProperties',
          paragraphId,
          properties: [
            ...properties,
            {
              localName: 'pStyle',
              attributes: { val: style ?? SCIOBOT_HORIZONTAL_RULE_STYLE_ID },
            },
          ],
        });
        if (!styled.ok) throw new Error(`The native block style was refused: ${styled.reason}.`);
        return withPart(working, styled.part);
      },
    ],
    false,
    true,
    'after'
  );
  if (!result.committed || result.rejected)
    return fail(String(result.reason ?? 'The native insertion did not commit.'));
  const id = session.paragraphIds().find((id) => !before.has(id));
  if (!id) return fail('The inserted paragraph was not allocated.');
  const point = { paragraphId: id, offset: 0 };
  surface.setSelection({ anchor: point, head: point });
  return { success: true, paragraphId: id };
}
