import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  hyperlinkTargetOf,
  relationshipTargetIn,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { planHyperlinkFormat, type HyperlinkFormatTarget } from '../../src/diff/run-format';
import type { CanonicalEditPlan } from '../../src/diff/plan';
import { readEditorRevision } from '../../src/projection/revision-cache';
import { readCommittedView } from '../../src/review/committed-view';
import { readWorkingView } from '../../src/review/working-view';
import { docxFixture } from './document';

export const LINK_CONTEXT = Object.freeze({
  documentId: 'link-review-test',
  epoch: 'mount-1',
  author: 'ScioBot',
  date: '2026-10-20T09:00:00Z',
});
const RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
/** Two links share one relationship; the first carries a tooltip and marked runs. */
export const linkDocument = () =>
  docxFixture(
    '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:hyperlink r:id="shared" w:tooltip="Old tip"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:b/></w:rPr><w:t xml:space="preserve">the </w:t></w:r><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:i/></w:rPr><w:t>site</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> or </w:t></w:r><w:hyperlink r:id="shared"><w:r><w:t>twin</w:t></w:r></w:hyperlink></w:p><w:p><w:bookmarkStart w:id="3" w:name="Goal"/><w:r><w:t>Tail</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p>',
    {
      relationships: `<Relationship Id="shared" Type="${RELATIONSHIP}/hyperlink" Target="https://example.com" TargetMode="External"/>`,
    }
  );
export const children = (node: OoxmlNode | null | undefined) =>
  node && node.kind !== 'textValue' ? node.children : [];
export const linksOf = (part: OoxmlPart) =>
  children(part.root.children[0]).flatMap((paragraph) =>
    children(paragraph).filter((node) => node.kind === 'hyperlink')
  );
export function linkPlan(
  editor: DocxEditorInstance,
  target: HyperlinkFormatTarget,
  index = 0
): CanonicalEditPlan {
  const session = editor.surface!.session;
  const result = planHyperlinkFormat(
    {
      revision: readEditorRevision(editor, LINK_CONTEXT.documentId, LINK_CONTEXT.epoch)!,
      pkg: session.currentPackage(),
      partName: session.part().name,
      hyperlinkId: linksOf(session.part())[index]!.id,
    },
    target
  );
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
/** Each live link as text and resolved target, in reading order. */
export const linkTargets = (editor: DocxEditorInstance) => {
  const session = editor.surface!.session,
    pkg = session.currentPackage(),
    part = session.part();
  return linksOf(part).map((link) => {
    const target = hyperlinkTargetOf(link as Extract<OoxmlNode, { kind: 'hyperlink' }>, (id) =>
      relationshipTargetIn(pkg, part.name, id)
    );
    return [target.authored, target.tooltip ?? null];
  });
};
export async function linkViews(editor: DocxEditorInstance) {
  const session = editor.surface!.session;
  const revision = readEditorRevision(editor, LINK_CONTEXT.documentId, LINK_CONTEXT.epoch)!;
  const original = await readCommittedView(session, revision),
    working = await readWorkingView(session, revision);
  if (!original.success || !working.success) throw new Error('Missing review view');
  return [original.view.projection.text, working.view.projection.text];
}
