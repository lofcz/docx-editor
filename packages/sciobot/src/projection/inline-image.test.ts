import { expect, test } from 'bun:test';
import { createInlineImagePayload } from '../blocks/inline-image-codec';
import { serializeInlineImage } from './inline-image';
import { parseTargetMarkdown, type TargetAstNode } from './parse-target';
const images = (node: TargetAstNode): TargetAstNode[] =>
  node.type === 'image' ? [node] : ((node.children ?? []) as TargetAstNode[]).flatMap(images);
for (const input of [
  { reference: 'asset:cat', alt: 'Cat', title: 'Photo' },
  {
    reference: 'https://example.test/a_(b).png?q=a&b=𝛼',
    alt: 'α [*] | ` <> \\ \n',
    title: '"Quoted"\r\n | \\ end',
  },
  { reference: 'asset: spaced reference\n', alt: 'a\tb\r\nc | d', title: null },
  { reference: 'asset:empty', alt: '', title: '' },
])
  test(`inline image provenance retains exact reference, alt and title in prose and GFM: ${JSON.stringify(input)}`, () => {
    const payload = createInlineImagePayload('picture', {
      ...input,
      contentHash: '0'.repeat(64),
      mediaPartName: '/word/media/x.png',
      widthPoints: 12,
      heightPoints: 8,
    });
    const serialized = serializeInlineImage(payload);
    for (const source of [
      `Before ${serialized.text} after`,
      `| H |\n| - |\n| Before ${serialized.text} after |`,
    ]) {
      const parsed = parseTargetMarkdown(source);
      if (!parsed.success) throw Error(parsed.error);
      const recovered = images(parsed.target.normalizedAst);
      expect(recovered).toHaveLength(1);
      const node = recovered[0]!;
      expect({ url: node.url, alt: node.alt, title: node.title }).toEqual({
        url: input.reference,
        alt: input.alt,
        title: input.title,
      });
      expect(serialized.properties).toEqual([]);
    }
  });
for (const syntax of [
  '<scio-inline-image reference="asset:x" alt="A" extra="keep" />',
  '<scio-inline-image reference="asset:x" reference="asset:y" alt="A" />',
  '<scio-inline-image reference={fetch("asset:x")} alt="A" />',
  '<scio-inline-image reference="asset:x" alt="A">keep</scio-inline-image>',
])
  test(`reserved image data syntax preserves unsupported fields as opaque AST: ${syntax}`, () => {
    const parsed = parseTargetMarkdown(`Before ${syntax} after`);
    if (!parsed.success) throw Error(parsed.error);
    expect(images(parsed.target.normalizedAst)).toEqual([]);
  });
