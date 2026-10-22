import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  bodyStoryRoot,
  directParagraphProperties,
  findNode,
  paragraphTextOf,
  parentNodeOf,
  readOoxmlPart,
  storyParagraphs,
  withPart,
  WML_NAMESPACE_URI,
} from '@docx-editor.dev/core/store';
import {
  horizontalRuleStyleState,
  HORIZONTAL_RULE_PARAGRAPH_XML,
  SCIOBOT_HORIZONTAL_RULE_STYLE_ID,
} from '../native/horizontal-rule';

/**
 * Insert the rule before the existing text owner, then consume only its typed prefix.
 * The source paragraph remains the editable follower, so offline peer text keeps its owner.
 */
export function applyHorizontalRuleInput(
  editor: DocxEditorInstance,
  prefix: '--' | '___'
): boolean {
  const surface = editor.surface;
  if (!surface || surface.editingMode() !== 'edit') return false;
  surface.flushPendingInput();
  const session = surface.session,
    part = session.part(),
    pkg = session.currentPackage();
  const { anchor, head } = surface.state().selection;
  if (
    part.name !== pkg.mainDocumentPart ||
    horizontalRuleStyleState(pkg) !== 'present' ||
    anchor.paragraphId !== head.paragraphId ||
    anchor.offset !== head.offset ||
    head.offset !== prefix.length ||
    paragraphTextOf(part, head.paragraphId) !== prefix
  )
    return false;
  const paragraph = findNode(part, head.paragraphId),
    parent = parentNodeOf(part, head.paragraphId);
  if (
    paragraph?.kind !== 'paragraph' ||
    !parent ||
    directParagraphProperties(part, paragraph.id).some((p) => p.localName === 'numPr')
  )
    return false;
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body>${HORIZONTAL_RULE_PARAGRAPH_XML.replace(/<w:pStyle[^>]*\/>/, '')}</w:body></w:document>`,
    { name: part.name, contentType: part.contentType }
  );
  if (!parsed.ok) return false;
  const seed = storyParagraphs(bodyStoryRoot(parsed.part)!)[0];
  if (seed?.kind !== 'paragraph') return false;
  const before = new Set(session.paragraphIds());
  const result = surface.applyAutomationOps(
    () => [
      { op: 'insertParagraphs', parentId: parent.id, beforeId: paragraph.id, paragraphs: [seed] },
      { op: 'deleteText', paragraphId: paragraph.id, start: 0, end: prefix.length },
    ],
    { kind: 'body' },
    [
      (working) => {
        const current = working.parts.get(part.name)!;
        const allocated = storyParagraphs(bodyStoryRoot(current)!).filter(
          (node) => !before.has(node.id)
        );
        if (allocated.length !== 1) throw new Error('The native rule allocation is ambiguous.');
        const id = allocated[0]!.id;
        const styled = applyTreeOp(current, {
          op: 'setParagraphProperties',
          paragraphId: id,
          properties: [
            ...directParagraphProperties(current, id),
            {
              localName: 'pStyle',
              attributes: { val: SCIOBOT_HORIZONTAL_RULE_STYLE_ID },
            },
          ],
        });
        if (!styled.ok) throw new Error(`The native rule style was refused: ${styled.reason}.`);
        return withPart(working, styled.part);
      },
    ],
    false,
    true,
    'after'
  );
  if (!result.committed || result.rejected) return false;
  const point = { paragraphId: paragraph.id, offset: 0 };
  surface.setSelection({ anchor: point, head: point });
  return true;
}
