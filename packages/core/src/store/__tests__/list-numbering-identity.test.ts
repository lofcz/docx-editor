import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPart,
  serializeOoxmlPart,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { applyTreeOp, type TreeDocOp } from '../store/tree-ops.ts';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function load(body: string) {
  const read = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:u="urn:unknown"><w:body>${body}</w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  return read.part;
}
function find(part: OoxmlPart, name: string): Exclude<OoxmlNode, { kind: 'textValue' }> {
  const pending: OoxmlNode[] = [part.root];
  while (pending.length) {
    const node = pending.shift()!;
    if (node.kind === 'textValue') continue;
    if (node.localName === name && node.namespaceUri === W) return node;
    pending.unshift(...node.children);
  }
  throw new Error(`Missing ${name}`);
}
function apply(part: OoxmlPart, op: TreeDocOp) {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}
const paragraph = (level: string, num: string) =>
  `<w:p><w:pPr><w:numPr u:guard="keep"><w:ilvl w:val="${level}" u:level="yes"/><w:numId w:val="${num}" u:counter="yes"/><u:opaque u:key="stable"><u:leaf/></u:opaque></w:numPr><w:spacing w:after="120"/></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>`;
for (const op of ['setListNumbering', 'setListLevel'] as const)
  for (const tracked of [false, true])
    test(`${op} retains numbering identities and complete unknown metadata; tracked=${tracked}`, () => {
      const before = load(paragraph('0', '4'));
      const ids = ['numPr', 'ilvl', 'numId'].map((name) => find(before, name).id);
      const revision = tracked ? { author: 'QA', date: '2026-10-24T12:00:00Z' } : undefined;
      const after = apply(before, {
        op,
        paragraphId: find(before, 'p').id,
        level: 2,
        ...(op === 'setListNumbering' ? { numId: '9' } : {}),
        revision,
      } as TreeDocOp);
      expect(['numPr', 'ilvl', 'numId'].map((name) => find(after, name).id)).toEqual(ids);
      const expected = load(paragraph('2', op === 'setListNumbering' ? '9' : '4'));
      const accepted = tracked ? apply(after, { op: 'acceptAllRevisions' }) : after;
      expect(canonicalOoxmlFingerprint(accepted)).toEqual(canonicalOoxmlFingerprint(expected));
      const reopened = readOoxmlPart(serializeOoxmlPart(accepted), {
        name: accepted.name,
        contentType: accepted.contentType,
      });
      if (!reopened.ok) throw new Error(reopened.reason);
      expect(canonicalOoxmlFingerprint(reopened.part)).toEqual(canonicalOoxmlFingerprint(expected));
      if (tracked) {
        const rejected = apply(after, { op: 'rejectAllRevisions' });
        expect(canonicalOoxmlFingerprint(rejected)).toEqual(canonicalOoxmlFingerprint(before));
      }
    });

test('adding a missing level retains the counter and foreign same-name metadata', () => {
  const before = load(
    '<w:p><w:pPr><w:numPr u:keep="yes"><u:ilvl u:val="opaque"/><w:numId w:val="4" u:keep="counter"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
  );
  const after = apply(before, { op: 'setListLevel', paragraphId: find(before, 'p').id, level: 3 });
  expect(find(after, 'numPr').id).toBe(find(before, 'numPr').id);
  expect(find(after, 'numId').id).toBe(find(before, 'numId').id);
  const expected = load(
    '<w:p><w:pPr><w:numPr u:keep="yes"><w:ilvl w:val="3"/><u:ilvl u:val="opaque"/><w:numId w:val="4" u:keep="counter"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
  );
  expect(canonicalOoxmlFingerprint(after)).toEqual(canonicalOoxmlFingerprint(expected));
});

test('adding a missing counter retains level identity and schema position', () => {
  const before = load(
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="2" u:keep="level"/><u:opaque/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
  );
  const after = apply(before, {
    op: 'setListNumbering',
    paragraphId: find(before, 'p').id,
    level: 2,
    numId: '9',
  });
  expect(find(after, 'numPr').id).toBe(find(before, 'numPr').id);
  expect(find(after, 'ilvl').id).toBe(find(before, 'ilvl').id);
  const expected = load(
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="2" u:keep="level"/><w:numId w:val="9"/><u:opaque/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
  );
  expect(canonicalOoxmlFingerprint(after)).toEqual(canonicalOoxmlFingerprint(expected));
});
