import { expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlNode } from '../package/ooxml-tree.ts';
import { sourceOriginIdentity } from '../store/run-origin-identity.ts';
import { mapOriginNodes, originNodes } from '../store/run-origin-metadata.ts';

function paragraph() {
  const read = readOoxmlPart(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Original</w:t><w:t> atom</w:t></w:r><w:r><w:t>Other</w:t></w:r></w:p></w:body></w:document>',
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  return originNodes(read.part.root).find((node) => node.kind === 'paragraph')!;
}

test('replicas share source identities while distinct native runs and text atoms stay distinct', () => {
  const nodes = originNodes(paragraph());
  const source = nodes.filter((node) => node.kind === 'run' || node.kind === 'text');
  const a = sourceOriginIdentity(nodes),
    b = sourceOriginIdentity(originNodes(paragraph()));
  expect(source.map(a)).toEqual(source.map(b));
  expect(new Set(source.map(a)).size).toBe(source.length);
});

test('source identities exclude text, formatting, and opaque attribute contents', () => {
  const original = paragraph(),
    nodes = originNodes(original);
  const changed = mapOriginNodes(nodes, (node) =>
    node.kind === 'textValue'
      ? { ...node, value: 'different confidential contents' }
      : ({
          ...node,
          attributes: [
            ...node.attributes,
            {
              kind: 'genericExtension',
              namespaceUri: 'urn:opaque',
              localName: 'value',
              prefix: 'opaque',
              value: 'private data',
            },
          ],
        } as OoxmlNode)
  );
  const before = sourceOriginIdentity(nodes),
    after = sourceOriginIdentity(originNodes(changed));
  expect(nodes.map(before)).toEqual(nodes.map(after));
});

test('an inserted native identity separates later preparation epochs without storing content', () => {
  const original = paragraph(),
    nodes = originNodes(original);
  const changed = { ...original, id: `${original.id}:new-generation` } as OoxmlNode;
  const before = sourceOriginIdentity(nodes),
    after = sourceOriginIdentity(originNodes(changed));
  for (const node of nodes) expect(after(node)).not.toBe(before(node));
});
