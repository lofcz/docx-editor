import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  contentControlsIn,
  findNode,
  readOoxmlPart,
  serializeOoxmlPart,
  type OoxmlNode,
  type OoxmlPart,
  type TreeDocOp,
} from '../index.ts';
const paragraph = '<w:p><w:r><w:t>A </w:t></w:r>CONTROL<w:r><w:t> Z</w:t></w:r></w:p>';
const control = (properties: string, content = '<w:r><w:t>Atom</w:t></w:r>') =>
  `<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
function load(body: string) {
  const result = readOoxmlPart(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
function elements(node: OoxmlNode, kind: OoxmlNode['kind']): OoxmlNode[] {
  return [
    ...(node.kind === kind ? [node] : []),
    ...(node.kind === 'textValue' ? [] : node.children.flatMap((child) => elements(child, kind))),
  ];
}
function apply(part: OoxmlPart, op: TreeDocOp) {
  const r = applyTreeOp(part, op);
  if (!r.ok) throw Error(r.reason);
  return r.part;
}
for (const properties of [
  '',
  '<w:lock w:val="contentLocked"/>',
  '<w:lock w:val="sdtContentLocked"/>',
  '<w:dataBinding w:xpath="/a" w:storeItemID="{G}"/>',
  '<w:temporary/><w:lock w:val="sdtLocked"/>',
])
  test(`native link envelopes retain entire controls: ${properties}`, () => {
    const original = load(paragraph.replace('CONTROL', control(properties))),
      p = elements(original.root, 'paragraph')[0]!;
    const before = contentControlsIn(original.root).map(({ node }) => ({
      id: node.id,
      word: canonicalOoxmlFingerprint(node),
    }));
    let linked = apply(original, {
      op: 'insertHyperlink',
      paragraphId: p.id,
      start: 0,
      end: 8,
      relationshipId: 'rId1',
      tooltip: 'Read',
    });
    const link = elements(linked.root, 'hyperlink')[0]!;
    expect(elements(link, 'contentControl').length).toBe(1);
    for (const entry of before)
      expect(canonicalOoxmlFingerprint(findNode(linked, entry.id)!)).toBe(entry.word);
    linked = apply(linked, {
      op: 'setHyperlinkTarget',
      linkId: link.id,
      relationshipId: 'rId2',
      tooltip: 'Changed',
    });
    expect(findNode(linked, link.id)?.id).toBe(link.id);
    for (const entry of before)
      expect(canonicalOoxmlFingerprint(findNode(linked, entry.id)!)).toBe(entry.word);
    const reopened = load(
      serializeOoxmlPart(linked)
        .replace(/^[\s\S]*?<w:body>/, '')
        .replace(/<\/w:body>[\s\S]*$/, '')
    );
    expect(canonicalOoxmlFingerprint(reopened)).toBe(canonicalOoxmlFingerprint(linked));
    linked = apply(linked, { op: 'removeHyperlink', linkId: link.id });
    expect(canonicalOoxmlFingerprint(linked)).toBe(canonicalOoxmlFingerprint(original));
  });
for (const boundary of [
  [1, 4],
  [3, 7],
  [3, 5],
] as const)
  test(`a link boundary cannot split locked content ${boundary}`, () => {
    const part = load(paragraph.replace('CONTROL', control('<w:lock w:val="sdtContentLocked"/>'))),
      p = elements(part.root, 'paragraph')[0]!;
    const before = canonicalOoxmlFingerprint(part);
    expect(
      applyTreeOp(part, {
        op: 'insertHyperlink',
        paragraphId: p.id,
        start: boundary[0],
        end: boundary[1],
        relationshipId: 'rId1',
      }).ok
    ).toBe(false);
    expect(canonicalOoxmlFingerprint(part)).toBe(before);
  });
for (const variant of ['ancestor', 'styled', 'tracked'] as const)
  test(`link envelopes retain ${variant} restrictions`, () => {
    const inner = paragraph.replace('CONTROL', control('<w:lock w:val="sdtContentLocked"/>'));
    const part = load(
        variant === 'ancestor' ? control('<w:lock w:val="contentLocked"/>', inner) : inner
      ),
      p = elements(part.root, 'paragraph')[0]!;
    const result = applyTreeOp(part, {
      op: 'insertHyperlink',
      paragraphId: p.id,
      start: 0,
      end: 8,
      relationshipId: 'rId1',
      ...(variant === 'styled' ? { styleId: 'Hyperlink' } : {}),
      ...(variant === 'tracked' ? { revision: { author: 'A', date: '2026-10-23T09:00:00Z' } } : {}),
    });
    expect(result.ok).toBe(false);
  });
