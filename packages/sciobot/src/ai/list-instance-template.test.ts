import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  readOoxmlPart,
} from '@docx-editor.dev/core/store';
import { docxFixture, W } from '../../test/support/document';
import { copyListInstanceTemplate, setListInstanceStart } from './list-instance-template';
const xml = `<w:numbering xmlns:w="${W}" xmlns:u="urn:opaque"><w:num w:numId="1" u:keep="yes"><w:abstractNumId w:val="7"/><w:lvlOverride w:ilvl="+02"><w:startOverride w:val="8" u:keep="start"/><w:lvl w:ilvl="2"><w:numFmt w:val="decimal"/><w:lvlText w:val="§ %3)"/><w:rPr><w:rFonts w:ascii="Georgia"/></w:rPr></w:lvl><u:keep/></w:lvlOverride><w:lvlOverride w:ilvl="4"><w:startOverride w:val="6"/></w:lvlOverride></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
function fixture(numbering = xml) {
  const read = readOoxmlPackage(docxFixture('<w:p/>', { numbering }));
  if (!read.ok) throw new Error(read.reason);
  return read.package;
}
test('counter fork and start replacement preserve the complete imported descriptor', () => {
  const original = fixture(),
    part = original.parts.get('/word/numbering.xml')!;
  const source = part.root.children[0]!;
  const copied = copyListInstanceTemplate(original, '1', '2');
  const changed = setListInstanceStart(copied, '2', 2, 1);
  const expected = readOoxmlPart(
    xml.replace(
      '<w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>',
      xml
        .slice(xml.indexOf('<w:num w:numId="1"'), xml.indexOf('<w:num w:numId="2"'))
        .replace('w:numId="1"', 'w:numId="2"')
        .replace('w:val="8"', 'w:val="1"')
    ),
    { name: part.name, contentType: part.contentType }
  );
  if (!expected.ok) throw new Error(expected.reason);
  expect(canonicalOoxmlFingerprint(changed.parts.get(part.name)!)).toEqual(
    canonicalOoxmlFingerprint(expected.part)
  );
  expect(canonicalOoxmlFingerprint(changed.parts.get(part.name)!.root.children[0]!)).toEqual(
    canonicalOoxmlFingerprint(source)
  );
  const ids = new Set<string>();
  const visit = (node: typeof source) => {
    expect(ids.has(node.id)).toBe(false);
    ids.add(node.id);
    if (node.kind !== 'textValue') node.children.forEach(visit);
  };
  visit(changed.parts.get(part.name)!.root);
});
test('start insertion keeps other overrides and has one native start per level', () => {
  const original = fixture();
  const changed = setListInstanceStart(
    setListInstanceStart(original, '1', 0, 0),
    '1',
    0,
    2147483647
  );
  const part = changed.parts.get('/word/numbering.xml')!;
  const expected = readOoxmlPart(
    xml.replace(
      '</w:num>',
      '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="2147483647"/></w:lvlOverride></w:num>'
    ),
    { name: part.name, contentType: part.contentType }
  );
  if (!expected.ok) throw new Error(expected.reason);
  expect(canonicalOoxmlFingerprint(part)).toEqual(canonicalOoxmlFingerprint(expected.part));
});
test('ambiguous imported starts and invalid bounded inputs refuse without mutation', () => {
  const original = fixture(
    xml.replace(
      '<w:startOverride w:val="8" u:keep="start"/>',
      '<w:startOverride w:val="8"/><w:startOverride w:val="9"/>'
    )
  );
  const before = canonicalOoxmlFingerprint(original.parts.get('/word/numbering.xml')!);
  expect(() => setListInstanceStart(original, '1', 2, 1)).toThrow('ambiguous');
  for (const [level, start] of [
    [-1, 1],
    [9, 1],
    [2, -1],
    [2, 2147483648],
    [2, NaN],
  ])
    expect(() => setListInstanceStart(original, '1', level!, start!)).toThrow();
  expect(() => copyListInstanceTemplate(original, '0', '2')).toThrow();
  expect(canonicalOoxmlFingerprint(original.parts.get('/word/numbering.xml')!)).toEqual(before);
});
