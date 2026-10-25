import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  directParagraphProperties,
  readOoxmlPart,
  serializeOoxmlPart,
  canonicalOoxmlFingerprint,
  type OoxmlNode,
} from '../index';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function indent(node: OoxmlNode): OoxmlNode | undefined {
  if (node.kind === 'textValue') return;
  if (node.namespaceUri === W && node.localName === 'ind') return node;
  for (const child of node.children) {
    const found = indent(child);
    if (found) return found;
  }
}
for (const tracked of [false, true])
  test(`paragraph indent attribute changes preserve the existing scalar holder and opaque children: tracked=${tracked}`, () => {
    const parsed = readOoxmlPart(
      `<w:document xmlns:w="${W}"><w:body><w:p><w:pPr><w:ind xmlns:u="urn:opaque" w:left="720" w:right="240" u:keep="yes"><u:child/></w:ind></w:pPr><w:r><w:t>Item</w:t></w:r></w:p></w:body></w:document>`,
      {
        name: '/word/document.xml',
        contentType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
      }
    );
    if (!parsed.ok) throw new Error(parsed.reason);
    const body = parsed.part.root.children[0]!;
    if (body.kind === 'textValue') throw new Error('Missing body.');
    const id = body.children[0]!.id,
      original = indent(parsed.part.root)!;
    const changed = applyTreeOp(parsed.part, {
      op: 'setParagraphProperties',
      paragraphId: id,
      properties: directParagraphProperties(parsed.part, id).map((property) =>
        property.localName === 'ind'
          ? { ...property, attributes: { ...property.attributes, right: '480' } }
          : property
      ),
      ...(tracked ? { revision: { author: 'User', date: '2026-01-01T00:00:00Z' } } : {}),
    });
    if (!changed.ok) throw new Error(changed.reason);
    const next = indent(changed.part.root)!;
    expect(next.id).toBe(original.id);
    if (next.kind === 'textValue' || original.kind === 'textValue')
      throw new Error('Missing indentation.');
    expect(next.children).toEqual(original.children);
    expect(next.namespaceBindings).toEqual(original.namespaceBindings);
    expect(next.attributes.find((attribute) => attribute.namespaceUri === 'urn:opaque')).toEqual(
      original.attributes.find((attribute) => attribute.namespaceUri === 'urn:opaque')
    );
    const reopened = readOoxmlPart(serializeOoxmlPart(changed.part), {
      name: changed.part.name,
      contentType: changed.part.contentType,
    });
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(canonicalOoxmlFingerprint(reopened.part)).toEqual(
      canonicalOoxmlFingerprint(changed.part)
    );
  });
