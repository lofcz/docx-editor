import { expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart, type OoxmlPart } from '../package/ooxml-tree.ts';
import { canonicalOoxmlFingerprint } from '../package/ooxml-serialize.ts';
import { bodyStoryRoot, storyParagraphs } from '../package/story-blocks.ts';
import { contentControlsIn } from '../package/content-control-nodes.ts';
import { applyTreeOp } from '../store/tree-ops.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text: string) => '<w:r><w:rPr><w:b/></w:rPr><w:t>' + text + '</w:t></w:r>';
const insertion = (content: string) =>
  '<w:ins w:id="7" w:author="AI" w:date="2026-10-24T12:00:00Z" e:keep="source">' +
  content +
  '</w:ins>';
function load(body: string): OoxmlPart {
  const result = readOoxmlPart(
    '<w:document xmlns:w="' +
      W +
      '" xmlns:e="urn:imported"><w:body>' +
      body +
      '</w:body></w:document>',
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
const paragraph = (part: OoxmlPart) => storyParagraphs(bodyStoryRoot(part)!)[0]!;
const apply = (part: OoxmlPart, op: TreeDocOp) => {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw Error(result.reason);
  return result.part;
};
const scopes: Record<
  'plain' | 'link' | 'namedLink' | 'control' | 'custom' | 'smart',
  (content: string) => string
> = {
  plain: (content: string) => content,
  link: (content: string) =>
    '<w:hyperlink w:anchor="target" w:tooltip="Imported" e:keep="link">' +
    content +
    '</w:hyperlink>',
  namedLink: (content: string) => scopes.link(content),
  custom: (content: string) =>
    '<w:customXml w:uri="urn:custom" w:element="slot"><w:customXmlPr><w:attr w:name="tag" w:val="keep"/></w:customXmlPr>' +
    content +
    '</w:customXml>',
  smart: (content: string) =>
    '<w:smartTag w:uri="urn:smart" w:element="slot"><w:smartTagPr><w:attr w:name="tag" w:val="keep"/></w:smartTagPr>' +
    content +
    '</w:smartTag>',
  control: (content: string) =>
    '<w:sdt><w:sdtPr><w:id w:val="4"/><w:alias w:val="Imported"/><w:tag w:val="imported"/></w:sdtPr><w:sdtContent>' +
    content +
    '</w:sdtContent></w:sdt>',
};
for (const kind of ['plain', 'link', 'namedLink', 'control', 'custom', 'smart'] as const)
  for (const decision of ['accept', 'reject'] as const)
    test(`committed insertion in imported pending ${kind}/${decision} keeps independent text, metadata, and stable owners`, () => {
      const wrap = scopes[kind];
      const before = load('<w:p>' + insertion(wrap(run('abcd'))) + '</w:p>');
      const source = canonicalOoxmlFingerprint(before);
      const owner =
        kind === 'control'
          ? contentControlsIn(before.root)[0]!.node.id
          : kind === 'namedLink'
            ? paragraph(before)
                .children.flatMap((node) => (node.kind === 'revisionInsert' ? node.children : []))
                .find((node) => node.kind === 'hyperlink')!.id
            : undefined;
      const changed = apply(before, {
        op: 'insertText',
        paragraphId: paragraph(before).id,
        offset: 2,
        text: 'XY',
        ...(owner ? { inside: owner } : {}),
      });
      const pending = load(
        '<w:p>' + wrap(insertion(run('ab')) + run('XY') + insertion(run('cd'))) + '</w:p>'
      );
      expect(canonicalOoxmlFingerprint(changed)).toBe(canonicalOoxmlFingerprint(pending));
      expect(canonicalOoxmlFingerprint(before)).toBe(source);
      if (kind === 'control') expect(contentControlsIn(changed.root)[0]!.node.id).toBe(owner);
      if (kind === 'namedLink')
        expect(paragraph(changed).children.find((node) => node.kind === 'hyperlink')!.id).toBe(
          owner
        );
      const resolved = apply(changed, {
        op: decision === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
      });
      const expected = load(
        '<w:p>' +
          wrap(decision === 'accept' ? run('ab') + run('XY') + run('cd') : run('XY')) +
          '</w:p>'
      );
      expect(canonicalOoxmlFingerprint(resolved)).toBe(canonicalOoxmlFingerprint(expected));
      const reopened = readOoxmlPart(serializeOoxmlPart(resolved), {
        name: resolved.name,
        contentType: resolved.contentType,
      });
      if (!reopened.ok) throw Error(reopened.reason);
      expect(canonicalOoxmlFingerprint(reopened.part)).toBe(canonicalOoxmlFingerprint(expected));
    });
for (const [op, element] of [
  ['insertTab', '<w:tab/>'],
  ['insertHardBreak', '<w:br/>'],
  ['insertPageBreak', '<w:br w:type="page"/>'],
] as const)
  test(`${op} stays committed inside an imported insertion`, () => {
    const before = load('<w:p>' + insertion(run('ab')) + '</w:p>');
    const changed = apply(before, { op, paragraphId: paragraph(before).id, offset: 1 });
    const rejected = apply(changed, { op: 'rejectAllRevisions' });
    expect(canonicalOoxmlFingerprint(rejected)).toBe(
      canonicalOoxmlFingerprint(load('<w:p><w:r><w:rPr><w:b/></w:rPr>' + element + '</w:r></w:p>'))
    );
  });

test('committed insertion keeps the established depth budget across nested controls', () => {
  const nested = (content: string) => {
    let result = content;
    for (let index = 0; index < 20; index++)
      result =
        '<w:sdt><w:sdtPr><w:id w:val="' +
        (100 + index) +
        '"/><w:tag w:val="level-' +
        index +
        '"/></w:sdtPr><w:sdtContent>' +
        result +
        '</w:sdtContent></w:sdt>';
    return result;
  };
  const before = load('<w:p>' + insertion(nested(run('abcd'))) + '</w:p>');
  const controls = contentControlsIn(before.root);
  const changed = apply(before, {
    op: 'insertText',
    paragraphId: paragraph(before).id,
    offset: 2,
    text: 'XY',
    inside: controls.at(-1)!.node.id,
  });
  expect(contentControlsIn(changed.root).map(({ node }) => node.id)).toEqual(
    controls.map(({ node }) => node.id)
  );
  const rejected = apply(changed, { op: 'rejectAllRevisions' });
  expect(canonicalOoxmlFingerprint(rejected)).toBe(
    canonicalOoxmlFingerprint(load('<w:p>' + nested(run('XY')) + '</w:p>'))
  );
});
