// Tracked hyperlink retargeting through Word content revisions.

import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPart,
  serializeOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { planRevisionBatch, revisionItemsOf } from '../index.ts';
import { applyTreeOp, validateTreeOp, type TreeDocOp } from '../store/tree-ops.ts';

const W = WML_NAMESPACE_URI;
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const revision = { author: 'ScioBot', date: '2026-10-20T08:00:00Z' };
const LINK =
  '<w:hyperlink r:id="rId1" w:tooltip="Old" w:history="1"><w:proofErr w:type="spellStart"/><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:b/></w:rPr><w:t xml:space="preserve">the </w:t></w:r><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:i/></w:rPr><w:t>site</w:t><w:tab/></w:r><w:proofErr w:type="spellEnd"/></w:hyperlink>';
const BODY = `<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>${LINK}<w:r><w:t xml:space="preserve"> and </w:t></w:r><w:hyperlink r:id="rId1"><w:r><w:t>twin</w:t></w:r></w:hyperlink></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>`;
function load(xml = BODY): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${xml}</w:body></w:document>`,
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}
const reopen = (part: OoxmlPart) => {
  const result = readOoxmlPart(serializeOoxmlPart(part), {
    name: part.name,
    contentType: part.contentType,
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
};
const kids = (node: OoxmlNode) => (node.kind === 'textValue' ? [] : node.children);
const links = (part: OoxmlPart) =>
  kids(kids(kids(part.root)[0]!)[0]!).filter((node) => node.kind === 'hyperlink');
function apply(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`${result.reason}`);
  return result.part;
}
function decide(part: OoxmlPart, action: 'accept' | 'reject'): OoxmlPart {
  const batch = planRevisionBatch(part, action);
  expect(batch.result.skipped).toEqual([]);
  let current = part;
  for (const op of batch.ops) current = apply(current, op);
  expect(revisionItemsOf(current)).toHaveLength(0);
  return current;
}
const xmlOf = (part: OoxmlPart) => serializeOoxmlPart(part);
/** Proofing marks carry no meaning; accepting hoists them out of the emptied old link. */
const proofless = (part: OoxmlPart) =>
  xmlOf(part)
    .replace(/<w:proofErr [^>]*\/>/g, '')
    .replace(/ xmlns:r="[^"]*"/g, '');
/** Visible text: struck text is markup, not reading text. */
const textOf = (part: OoxmlPart) =>
  xmlOf(part)
    .replace(/<w:delText[^>]*>[^<]*<\/w:delText>/g, '')
    .replace(/<[^>]+>/g, '');

const retargets: [string, (part: OoxmlPart) => TreeDocOp][] = [
  [
    'a URL',
    (part) =>
      ({
        op: 'setHyperlinkTarget',
        linkId: links(part)[0]!.id,
        relationshipId: 'rId9',
      }) as TreeDocOp,
  ],
  [
    'a bookmark target',
    (part) =>
      ({ op: 'setHyperlinkTarget', linkId: links(part)[0]!.id, anchor: 'Tail' }) as TreeDocOp,
  ],
  [
    'a tooltip only',
    (part) =>
      ({
        op: 'setHyperlinkTarget',
        linkId: links(part)[0]!.id,
        relationshipId: 'rId1',
        tooltip: 'New tip',
      }) as TreeDocOp,
  ],
];
for (const [name, build] of retargets)
  test(`tracked retarget to ${name} accepts to the direct command and rejects exactly`, () => {
    const part = load();
    const op = build(part);
    const tracked = { ...op, revision } as TreeDocOp;
    expect(validateTreeOp(part, tracked)).toBeNull();
    const pending = apply(part, tracked);
    const xml = xmlOf(pending);
    // The old link keeps its target and holds struck runs; the copy follows with the new one.
    expect(xml).toContain('<w:hyperlink r:id="rId1" w:history="1" w:tooltip="Old"><w:proofErr');
    expect(xml.match(/<w:delText/g)).toHaveLength(2);
    expect(xml.match(/<w:ins /g)).toHaveLength(1);
    expect(xml.indexOf('<w:del ')).toBeLessThan(xml.indexOf('<w:ins '));
    expect(xml.match(/<w:proofErr /g)).toHaveLength(2);
    // The twin link sharing the old relationship is untouched.
    expect(xml).toContain('<w:hyperlink r:id="rId1"><w:r><w:t>twin</w:t></w:r></w:hyperlink>');
    expect(textOf(pending)).toBe(textOf(part));
    expect(revisionItemsOf(pending).length).toBeGreaterThan(0);
    const direct = apply(part, op);
    const accepted = decide(pending, 'accept');
    expect(proofless(accepted)).toBe(proofless(direct));
    expect(canonicalOoxmlFingerprint(decide(pending, 'reject'))).toBe(
      canonicalOoxmlFingerprint(part)
    );
    // Saved bytes keep both halves and both decisions.
    const reopened = reopen(pending);
    expect(proofless(decide(reopened, 'accept'))).toBe(proofless(reopen(direct)));
    expect(canonicalOoxmlFingerprint(decide(reopened, 'reject'))).toBe(
      canonicalOoxmlFingerprint(part)
    );
  });

test('the proposal keeps every run property and the link attributes the command did not name', () => {
  const part = load();
  const pending = apply(part, { ...retargets[0]![1](part), revision } as TreeDocOp);
  const inserted =
    /<w:hyperlink [^>]*r:id="rId9"[^>]*><w:ins [^>]*>(.*?)<\/w:ins><\/w:hyperlink>/.exec(
      xmlOf(pending)
    );
  expect(inserted).not.toBeNull();
  expect(inserted![0]).toContain('w:history="1"');
  expect(inserted![0]).toContain('w:tooltip="Old"');
  expect(inserted![1]).toBe(
    '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:b/></w:rPr><w:t xml:space="preserve">the </w:t></w:r><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:i/></w:rPr><w:t>site</w:t><w:tab/></w:r>'
  );
});

test('a retarget naming the current target proposes nothing', () => {
  const part = load();
  const same = apply(part, {
    op: 'setHyperlinkTarget',
    linkId: links(part)[0]!.id,
    relationshipId: 'rId1',
    revision,
  } as TreeDocOp);
  expect(revisionItemsOf(same)).toHaveLength(0);
  expect(canonicalOoxmlFingerprint(same)).toBe(canonicalOoxmlFingerprint(part));
});

test('subranges, uncopyable content, nested revisions, and bad attribution refuse atomically', () => {
  const part = load();
  const id = links(part)[0]!.id;
  const base = { op: 'setHyperlinkTarget', linkId: id, relationshipId: 'rId9' };
  const refused = (target: OoxmlPart, op: object, reason: string) => {
    expect(validateTreeOp(target, op as TreeDocOp)).toBe(reason as never);
    const result = applyTreeOp(target, op as TreeDocOp);
    expect(result.ok).toBe(false);
  };
  refused(
    part,
    { ...base, linkId: links(part)[1]!.id, range: { start: 19, end: 21 }, revision },
    'unsupported'
  );
  expect(
    validateTreeOp(part, {
      ...base,
      linkId: links(part)[1]!.id,
      range: { start: 19, end: 21 },
    } as TreeDocOp)
  ).toBeNull();
  refused(part, { ...base, revision: { author: '' } }, 'invalidArgs');
  const field = load(
    `<w:p><w:hyperlink r:id="rId1"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>`
  );
  refused(field, { ...base, linkId: links(field)[0]!.id, revision }, 'unsupported');
  const nested = load(
    `<w:p><w:hyperlink r:id="rId1"><w:ins w:id="1" w:author="Other"><w:r><w:t>x</w:t></w:r></w:ins></w:hyperlink></w:p>`
  );
  refused(nested, { ...base, linkId: links(nested)[0]!.id, revision }, 'unsupported');
  // The untracked command is unchanged.
  expect(validateTreeOp(part, base as TreeDocOp)).toBeNull();
});
