import { readOoxmlPackage, canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';
import { docxFixture } from './document';
import { propertyEditor, propertyContext } from './property-execution';
import { groupToken } from './review-groups';
import { readWorkingView } from '../../src/review/working-view';
import { readCommittedView } from '../../src/review/committed-view';
import { parseTargetMarkdown } from '../../src/projection/parse-target';
import { targetAstKey } from '../../src/diff/target-key';
import { planTargetDocument } from '../../src/diff/plan-target';
export const context = { ...propertyContext, date: '2026-10-22T00:00:00.000Z' };
export type Editor = ReturnType<typeof propertyEditor>;
export const run = (text: string, marks = '') =>
  `<w:r><w:rPr>${marks}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
export const paragraph = (content: string) => `<w:p><w:pPr><w:rPr/></w:pPr>${content}</w:p>`;
export const linked = (text: string, tooltip = '') =>
  `<w:hyperlink r:id="shared"${tooltip ? ` w:tooltip="${tooltip}"` : ''}>${run(text)}</w:hyperlink>`;
export const document_ = (body = paragraph(run('Alpha end'))) =>
  docxFixture(body + paragraph(run('Tail')), {
    relationships:
      '<Relationship Id="shared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>',
  });
export function mount(body?: string) {
  const read = readOoxmlPackage(document_(body));
  if (!read.ok) throw Error(read.reason);
  return propertyEditor(read.package);
}
export const fingerprint = (editor: Editor) =>
  canonicalOoxmlFingerprint(editor.surface!.session.part());
export async function projection(editor: Editor, committed = false) {
  const result = await (committed ? readCommittedView : readWorkingView)(
    editor.surface!.session,
    groupToken(editor)
  );
  if (!result.success) throw Error(result.error);
  return result.view.projection;
}
export function ast(text: string) {
  const parsed = parseTargetMarkdown(text);
  if (!parsed.success) throw Error(parsed.error);
  return targetAstKey(parsed.target.normalizedAst);
}
export async function plan(editor: Editor, target: string) {
  const result = planTargetDocument(
    {
      pkg: editor.surface!.session.currentPackage(),
      revision: groupToken(editor),
      projection: await projection(editor),
    },
    target
  );
  if (!result.success) throw Error(result.error);
  return result.plan;
}
export const cases = [
  {
    name: 'full insertion',
    body: paragraph(run('Alpha end')),
    target: '[Alpha end](https://example.org/new)',
  },
  {
    name: 'two disjoint partial insertions',
    body: paragraph(run('Alpha end')),
    target: '[Alpha](https://example.org/a) [end](https://example.org/b)',
  },
  {
    name: 'marked source',
    body: paragraph(run('Alpha', '<w:b/>') + run(' end')),
    target: '[**Alpha** end](https://example.org/new)',
  },
  { name: 'whole unlink', body: paragraph(linked('Alpha end')), target: 'Alpha end' },
  {
    name: 'tooltip removal',
    body: paragraph(linked('Alpha end', 'Old')),
    target: '[Alpha end](https://example.com)',
  },
  {
    name: 'changed link text',
    body: paragraph(run('See ') + linked('docs') + run(' now.')),
    target: 'See [the manual](https://example.com) now.',
  },
  {
    name: 'fresh linked text',
    body: paragraph(run('Alpha end')),
    target: 'Alpha [new](https://example.org/new) end',
  },
  {
    name: 'replacement deletion gap',
    body: paragraph(run('old right')),
    target: '[new right](https://example.org/new)',
  },
  {
    name: 'linked run formatting',
    body: paragraph(run('Alpha end')),
    target: '[**Alpha end**](https://example.org/new)',
  },
  {
    name: 'two links and marks',
    body: paragraph(run('Alpha end')),
    target: '[**Alpha**](https://example.org/a) [*end*](https://example.org/b)',
  },
  {
    name: 'expanded old wrapper',
    body: paragraph(linked('Alpha') + run(' end')),
    target: '[Alpha end](https://example.com)',
  },
] as const;
