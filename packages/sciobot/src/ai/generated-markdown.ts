import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { normalizeLogicalBlocks } from '../diff/logical-blocks';
import {
  documentQuoteStyle,
  insertedBlockProblem,
  logicalBlockAst,
} from '../diff/inserted-fragment-policy';
import { planInsertedBlocks } from '../diff/inserted-blocks';
import { resolveTargetLists } from '../diff/target-lists';
import { quoteParagraphStream, quoteSiblingInsertion } from '../diff/target-quotes';
import { parseTargetMarkdown } from '../projection/parse-target';
import type { GeneratedFragment } from './stream-insert';

/** Compile completed generated Markdown with the same native policies as D05. */
export function prepareGeneratedMarkdown(
  source: {
    readonly pkg: OoxmlPackage;
    readonly revision: RevisionToken;
    readonly partName: string;
  },
  markdown: string,
  allocateBlockId?: () => string
): { readonly success: true; readonly fragment: GeneratedFragment } | NativeEditFailure {
  const parsed = parseTargetMarkdown(markdown);
  if (!parsed.success) return parsed;
  const normalized = normalizeLogicalBlocks(parsed.target);
  if (!normalized.success) return normalized;
  const quotes = quoteParagraphStream(normalized.document);
  const blocks = quotes.document.blocks;
  const targets = blocks.map(logicalBlockAst);
  const lists = resolveTargetLists(source, blocks, []);
  if (!lists.success) return lists;
  for (const [index, block] of blocks.entries()) {
    const problem = insertedBlockProblem(block, targets[index]!, lists.targets[index]!);
    if (problem) return problem;
  }
  const wrapped = quoteSiblingInsertion(
    targets,
    blocks.map((block, index) =>
      quotes.quoted.has(block) ? quotes.origins.get(String(index))!.slice(0, -1).join('/') : null
    )
  );
  const quoted = wrapped.inserted.flatMap((ast, index) =>
    ast.type === 'blockquote' ? [index] : []
  );
  const quoteStyle = quoted.length ? documentQuoteStyle(source.pkg) : null;
  if (quoted.length && !quoteStyle)
    return {
      success: false,
      code: 'unsupported-operation',
      error: 'A generated block quote needs a Quote paragraph style.',
    };
  const nested = planInsertedBlocks(
    source,
    blocks.map((block, index) => ({ block, ast: targets[index]! })),
    allocateBlockId
  );
  if (!nested.success) return nested;
  return {
    success: true,
    fragment: {
      inserted: wrapped.inserted,
      lists: [
        ...blocks.flatMap((block, index) =>
          block.kind === 'paragraph' && !quotes.quoted.has(block)
            ? [{ path: wrapped.path([index]), target: lists.targets[index]! }]
            : []
        ),
        ...nested.plan.lists.map((binding) => ({ ...binding, path: wrapped.path(binding.path) })),
      ],
      blocks: nested.plan.blocks.map((binding) => ({
        ...binding,
        path: wrapped.path(binding.path),
      })),
      styles: quoted.map((index) => ({ path: [index], styleId: quoteStyle! })),
    },
  };
}
