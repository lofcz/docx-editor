import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPart,
  serializeOoxmlPart,
  propertyContainer,
  runPropertyEdits,
  runsCovering,
} from '../index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
test('run background edits preserve foreign metadata and unchanged themed shading through other marks', () => {
  const read = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:x="urn:keep"><w:body><w:p><w:r><w:rPr><w:color w:val="112233" w:themeColor="accent1" x:keep="color"/><w:shd w:val="clear" w:fill="DDEEFF" w:themeFill="accent2" x:keep="shade"><x:extra/></w:shd><w:lang w:val="cs-CZ"/></w:rPr><w:t>Alpha</w:t></w:r></w:p></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  let part = read.part;
  const paragraph =
    part.root.children[0]!.kind === 'textValue'
      ? null
      : part.root.children[0]!.children.find((n) => n.kind === 'paragraph');
  if (!paragraph) throw new Error('Missing paragraph');
  const id = paragraph.id;
  const run = runsCovering(part, id, 0, 5)[0]!;
  const props = propertyContainer(run, 'runProperties', 'rPr');
  if (!props || props.kind === 'textValue') throw new Error('Missing properties');
  const shade = props.children.find((n) => n.kind !== 'textValue' && n.localName === 'shd')!;
  const original = canonicalOoxmlFingerprint(shade);
  for (const incoming of [
    { localName: 'b' },
    { localName: 'shd', attributes: { val: 'clear', fill: '17A2B8', color: 'auto' } },
  ]) {
    const edit = runPropertyEdits(part, id, 0, 5, incoming)[0]!;
    const result = applyTreeOp(part, {
      op: 'setRunProperties',
      paragraphId: id,
      start: 0,
      end: 5,
      properties: edit.properties,
    });
    if (!result.ok) throw new Error(result.reason);
    part = result.part;
    const xml = serializeOoxmlPart(part);
    expect(xml).toContain('x:keep="shade"');
    expect(xml).toContain('x:keep="color"');
    expect(xml).toContain('<x:extra/>');
    expect(xml).toContain('w:themeColor="accent1"');
    expect(xml).toContain('w:val="cs-CZ"');
    if (incoming.localName === 'b')
      expect(canonicalOoxmlFingerprint(findNode(part, shade.id)!)).toEqual(original);
    else {
      expect(xml).toContain('w:fill="17A2B8"');
      expect(xml).not.toContain('w:themeFill="accent2"');
    }
    expect(paragraphTextOf(part, id)).toBe('Alpha');
  }
});

test('empty color properties remove their native elements without changing unrelated properties', () => {
  const read = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:b/><w:color w:val="CC0000"/><w:highlight w:val="yellow"/><w:shd w:val="clear" w:fill="123456"/></w:rPr><w:t>Alpha</w:t></w:r></w:p></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  let part = read.part;
  const body = part.root.children[0]!;
  if (body.kind === 'textValue') throw new Error('Missing body');
  const id = body.children[0]!.id;
  for (const localName of ['color', 'shd']) {
    const edit = runPropertyEdits(part, id, 0, 5, { localName })[0]!;
    const result = applyTreeOp(part, {
      op: 'setRunProperties',
      paragraphId: id,
      start: 0,
      end: 5,
      properties: edit.properties,
    });
    if (!result.ok) throw new Error(result.reason);
    part = result.part;
    const xml = serializeOoxmlPart(part);
    expect(xml).not.toContain(`<w:${localName}`);
    expect(xml).toContain('<w:b/>');
  }
  expect(serializeOoxmlPart(part)).not.toContain('<w:highlight');
});
