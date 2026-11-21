import { unified } from 'unified';
import remarkParse from 'remark-parse';
import remarkStringify from 'remark-stringify';
import remarkGfm from 'remark-gfm';
import remarkMdx from 'remark-mdx';
import remarkMath from 'remark-math';
import type { DocumentNode } from './document-types';
import { MAX_SOURCE_BYTES, syntaxData } from './document-data';

const syntax = unified()
  .use(remarkParse)
  .use(remarkMdx)
  .use(remarkGfm)
  .use(remarkMath)
  .use(remarkStringify, { bullet: '-', fences: true })
  .freeze();

export function parseDocumentSyntax(source: string): DocumentNode {
  if (
    source.length > MAX_SOURCE_BYTES ||
    new TextEncoder().encode(source).length > MAX_SOURCE_BYTES
  )
    throw new RangeError('MDX source exceeds the document limit.');
  return syntaxData(syntax.parse(source)) as unknown as DocumentNode;
}

export function printDocumentSyntax(root: DocumentNode): string {
  return syntax.stringify(root as unknown as ReturnType<typeof syntax.parse>);
}
