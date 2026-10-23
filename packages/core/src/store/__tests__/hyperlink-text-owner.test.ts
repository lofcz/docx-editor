import { formsProtectionRefusal } from '../store/tree-op-content-controls.ts';
import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  contentControlsIn,
  findNode,
  paragraphOffsetIndex,
  paragraphTextOf,
  readOoxmlPart,
  serializeOoxmlPart,
  validateTreeOp,
  type OoxmlNode,
  type OoxmlPart,
  type TreeDocOp,
} from '../index.ts';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const run = '<w:r><w:rPr><w:color w:val="112233"/></w:rPr><w:t>label</w:t></w:r>';
const link = (content = run) =>
  `<w:hyperlink r:id="existing" w:history="1" w:tgtFrame="_blank">${content}</w:hyperlink>`;
const paragraph = (content: string) =>
  `<w:p><w:r><w:t>Keep. </w:t></w:r>${content}<w:r><w:t> End.</w:t></w:r></w:p>`;
const control = (properties: string, content: string) =>
  `<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
function load(body: string, name = '/word/document.xml') {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}</w:body></w:document>`,
    { name, contentType: 'application/xml' }
  );
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
function nodes(node: OoxmlNode): OoxmlNode[] {
  return [node, ...(node.kind === 'textValue' ? [] : node.children.flatMap(nodes))];
}
function insertion(part: OoxmlPart, offset = 11): TreeDocOp {
  return {
    op: 'insertText',
    paragraphId: nodes(part.root).find((n) => n.kind === 'paragraph')!.id,
    inside: nodes(part.root).find((n) => n.kind === 'hyperlink')!.id,
    offset,
    text: ' ',
  };
}
for (const offset of [6, 11])
  test(`explicit link ownership preserves the native envelope at ${offset}`, () => {
    const part = load(paragraph(link())),
      op = insertion(part, offset);
    const original = findNode(part, op.op === 'insertText' ? op.inside! : '')!;
    const result = applyTreeOp(part, op);
    if (!result.ok) throw Error(result.reason);
    const kept = findNode(result.part, original.id)!;
    expect(kept.kind).toBe('hyperlink');
    if (kept.kind !== 'hyperlink' || original.kind !== 'hyperlink') throw Error('Missing link');
    expect(kept.attributes).toEqual(original.attributes);
    expect(paragraphTextOf(result.part, op.op === 'insertText' ? op.paragraphId : '')).toBe(
      offset === 6 ? 'Keep.  label End.' : 'Keep. label  End.'
    );
    const p = nodes(result.part.root).find((n) => n.kind === 'paragraph')!;
    if (p.kind !== 'paragraph') throw Error('Missing paragraph');
    expect(paragraphOffsetIndex(p).spanOf(kept)).toEqual({ start: 6, end: 12 });
    expect(serializeOoxmlPart(result.part)).toContain('112233');
    const reopened = readOoxmlPart(serializeOoxmlPart(result.part), {
      name: part.name,
      contentType: part.contentType,
    });
    if (!reopened.ok) throw Error(reopened.reason);
    expect(canonicalOoxmlFingerprint(reopened.part)).toBe(canonicalOoxmlFingerprint(result.part));
  });
test('ordinary trailing-link typing keeps its established outside ownership', () => {
  const part = load(paragraph(link())),
    { inside: _, ...op } = insertion(part) as Extract<TreeDocOp, { op: 'insertText' }>;
  const result = applyTreeOp(part, op);
  if (!result.ok) throw Error(result.reason);
  const p = nodes(result.part.root).find((n) => n.kind === 'paragraph')!;
  if (p.kind !== 'paragraph') throw Error('Missing paragraph');
  expect(paragraphOffsetIndex(p).spanOf(nodes(p).find((n) => n.kind === 'hyperlink')!)).toEqual({
    start: 6,
    end: 11,
  });
});
for (const properties of [
  '<w:lock w:val="contentLocked"/>',
  '<w:lock w:val="sdtContentLocked"/>',
  '<w:dataBinding w:xpath="/a" w:storeItemID="{G}"/>',
  '<w:temporary/><w:lock w:val="sdtLocked"/>',
])
  for (const placement of ['ancestor', 'nested'] as const)
    test(`named hyperlink insertion enforces ${placement} restrictions: ${properties}`, () => {
      const body =
        placement === 'ancestor'
          ? control(properties, paragraph(link()))
          : paragraph(link(control(properties, run)));
      const part = load(body),
        before = canonicalOoxmlFingerprint(part);
      const op = insertion(part);
      expect(validateTreeOp(part, op)).toBe(
        properties.includes('dataBinding') ? 'bound' : 'locked'
      );
      expect(applyTreeOp(part, op).ok).toBe(false);
      expect(canonicalOoxmlFingerprint(part)).toBe(before);
    });
test('an unlocked temporary control at the actual named landing unwraps on the first edit', () => {
  const part = load(paragraph(link(control('<w:temporary/>', run))));
  const result = applyTreeOp(part, insertion(part));
  if (!result.ok) throw Error(result.reason);
  expect(contentControlsIn(result.part.root)).toEqual([]);
  expect(nodes(result.part.root).filter((n) => n.kind === 'hyperlink')).toHaveLength(1);
});
test('a named hyperlink cannot bypass document forms protection', () => {
  const part = load(paragraph(link())),
    settings = readOoxmlPart(
      `<w:settings xmlns:w="${W}"><w:documentProtection w:edit="forms" w:enforcement="1"/></w:settings>`,
      { name: '/word/settings.xml', contentType: 'application/xml' }
    );
  if (!settings.ok) throw Error(settings.reason);
  expect(formsProtectionRefusal(part, settings.part, insertion(part))).toBe('locked');
});
test('invalid owners, foreign paragraphs, outside offsets and tracked ownership refuse atomically', () => {
  const part = load(paragraph(link()) + '<w:p><w:r><w:t>Other</w:t></w:r></w:p>'),
    op = insertion(part) as Extract<TreeDocOp, { op: 'insertText' }>;
  const p = nodes(part.root).filter((n) => n.kind === 'paragraph');
  for (const changed of [
    { ...op, inside: 'missing' },
    { ...op, inside: p[0]!.id },
    { ...op, paragraphId: p[1]!.id, offset: 0 },
    { ...op, offset: 5 },
    { ...op, offset: 12 },
    { ...op, revision: { author: 'Alice', date: '2026-10-23T10:00:00.000Z' } },
  ]) {
    const before = canonicalOoxmlFingerprint(part);
    expect(applyTreeOp(part, changed).ok).toBe(false);
    expect(canonicalOoxmlFingerprint(part)).toBe(before);
  }
});
