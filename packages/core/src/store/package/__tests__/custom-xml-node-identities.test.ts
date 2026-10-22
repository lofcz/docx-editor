import { expect, test } from 'bun:test';
import { blankDocumentBytes } from '../../../index.ts';
import { canonicalOoxmlFingerprint } from '../index.ts';
import { readOoxmlPackage, writeOoxmlPackage, withPart } from '../ooxml-package.ts';
import { readOoxmlPart } from '../ooxml-tree.ts';
import { withCustomXmlDataPart } from '../custom-xml-part.ts';
import { customXmlNodes, withCustomXmlNode } from '../custom-xml-nodes.ts';

function imported(content: string) {
  const source = readOoxmlPackage(blankDocumentBytes());
  if (!source.ok) throw new Error(source.reason);
  const created = withCustomXmlDataPart(
    source.package,
    source.package.mainDocumentPart,
    'urn:test:identity',
    'store'
  );
  if (!created.part) throw new Error('No custom XML part.');
  const old = created.pkg.parts.get(created.part.partName)!;
  const part = readOoxmlPart(
    `<store xmlns="urn:test:identity" xmlns:f="urn:test:foreign">${content}</store>`,
    { name: old.name, contentType: old.contentType }
  );
  if (!part.ok) throw new Error(part.reason);
  return { pkg: withPart(created.pkg, part.part), partName: old.name, part: part.part };
}

test('updating an imported payload retains wrapper, value, and unrelated XML identities', () => {
  const original = imported(
    '<node id="item" f:keep="node"><label f:keep="label">Old</label><data f:keep="data">{"before":true}</data><f:extra>untouched</f:extra></node><node id="other"><label>Other</label><data>42</data></node>'
  );
  const before = original.part.root.children[0]!;
  if (before.kind === 'textValue') throw new Error('Expected node.');
  const other = original.part.root.children[1]!;
  const body = original.pkg.parts.get(original.pkg.mainDocumentPart)!;
  const bodyFingerprint = canonicalOoxmlFingerprint(body);
  const next = withCustomXmlNode(original.pkg, original.partName, {
    id: 'item',
    label: ' New ',
    data: '{"after":true}',
  });
  const part = next.parts.get(original.partName)!;
  const updated = part.root.children[0]!;
  if (updated.kind === 'textValue') throw new Error('Expected node.');
  expect(updated.id).toBe(before.id);
  expect(updated.attributes).toEqual(before.attributes);
  expect(updated.namespaceBindings).toEqual(before.namespaceBindings);
  for (let index = 0; index < 2; index++) {
    const a = before.children[index]!,
      b = updated.children[index]!;
    expect(b.id).toBe(a.id);
    if (a.kind === 'textValue' || b.kind === 'textValue')
      throw new Error('Expected value elements.');
    expect(b.children[0]!.id).toBe(a.children[0]!.id);
    expect(b.attributes.some((attr) => attr.namespaceUri === 'urn:test:foreign')).toBe(true);
  }
  expect(updated.children[2] === before.children[2]).toBe(true);
  expect(part.root.children[1] === other).toBe(true);
  expect(next.parts.get(body.name) === body).toBe(true);
  expect(canonicalOoxmlFingerprint(body)).toBe(bodyFingerprint);
  expect(customXmlNodes(next, original.partName)).toEqual([
    { id: 'item', label: ' New ', data: '{"after":true}' },
    { id: 'other', label: 'Other', data: '42' },
  ]);
  const reopened = readOoxmlPackage(writeOoxmlPackage(next));
  if (!reopened.ok) throw new Error(reopened.reason);
  expect(customXmlNodes(reopened.package, original.partName)).toEqual(
    customXmlNodes(next, original.partName)
  );
});

test('missing values are added while an existing empty value and foreign siblings retain identity', () => {
  const original = imported('<node id="item"><data/><f:extra keep="yes"/></node>');
  const before = original.part.root.children[0]!;
  if (before.kind === 'textValue') throw new Error('Expected node.');
  const next = withCustomXmlNode(original.pkg, original.partName, {
    id: 'item',
    label: 'Label',
    data: 'Value',
  });
  const after = next.parts.get(original.partName)!.root.children[0]!;
  if (after.kind === 'textValue') throw new Error('Expected node.');
  expect(after.id).toBe(before.id);
  expect(after.children[0]!.id).toBe(before.children[0]!.id);
  expect(after.children[1] === before.children[1]).toBe(true);
  expect(customXmlNodes(next, original.partName)).toEqual([
    { id: 'item', label: 'Label', data: 'Value' },
  ]);
});
