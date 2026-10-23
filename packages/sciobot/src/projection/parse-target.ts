import { remarkNativeImage } from './native-image';
import { remarkNativeMath } from './native-math';
import remarkGfm from 'remark-gfm';
import remarkMath from 'remark-math';
import remarkMdx from 'remark-mdx';
import remarkParse from 'remark-parse';
import { unified } from 'unified';
import type { JsonObject, JsonValue } from '../blocks/payload-schema';
import type { NativeEditFailure } from '../contracts/document';
import { normalizeTargetSource } from './normalize-target';
import {
  normalizationSourceRangeAt,
  splitsSurrogate,
  type TargetNormalizationMap,
  type TargetSourceRange,
} from './target-source-map';
export type {
  TargetNormalizationMap,
  TargetNormalizationStep,
  TargetNormalizationChange,
  TargetSourceRange,
} from './target-source-map';
import { remarkImageGalleries } from './target-galleries';
import { remarkTargetMentions } from './target-mentions';

/** A JSON read model. Canonical document nodes never enter this tree. */
export interface TargetAstNode extends JsonObject {
  readonly type: string;
}
export interface TargetAstRoot extends TargetAstNode {
  readonly type: 'root';
  readonly children: readonly TargetAstNode[];
}

export interface TargetParseLimits {
  readonly maxBytes: number;
  readonly maxNodes: number;
  readonly maxDepth: number;
}
export const DEFAULT_TARGET_PARSE_LIMITS: TargetParseLimits = Object.freeze({
  maxBytes: 2 * 1024 * 1024,
  maxNodes: 100_000,
  maxDepth: 64,
});

export interface ParsedTargetMarkdown {
  /** Original tool input, before the application normalizers run. */
  readonly sourceMarkdown: string;
  /** AST positions address this normalized string using UTF-16 offsets. */
  readonly markdown: string;
  /** Exact string transformation trace. This does not establish native ownership. */
  readonly normalization: TargetNormalizationMap;
  /** Retains source positions and complete expression source, without compiler ESTree caches. */
  readonly ast: TargetAstRoot;
  /** Position-independent data for symmetric before/after comparisons. */
  readonly normalizedAst: TargetAstRoot;
}

export interface TargetParseIssue {
  readonly code: 'syntax' | 'byte-limit' | 'node-limit' | 'depth-limit' | 'invalid-input';
  readonly message: string;
  readonly line?: number;
  readonly column?: number;
  /** Original parser diagnostic, retained for application tool feedback. */
  readonly ruleId?: string;
  readonly details?: readonly string[];
}
export type TargetParseResult =
  | { readonly success: true; readonly target: ParsedTargetMarkdown }
  | (NativeEditFailure & {
      readonly issues: readonly TargetParseIssue[];
      /** Diagnostic positions address this normalized, rejected string. */
      readonly markdown?: string;
    });

const createTargetProcessor = (preserveLinkedImages: boolean) =>
  unified()
    .use(remarkParse)
    .use(remarkMath)
    .use(remarkGfm)
    .use(remarkMdx)
    .use(remarkNativeMath)
    .use(remarkNativeImage)
    .use(remarkTargetMentions)
    .use(remarkImageGalleries, { preserveLinkedImages });

const processor = createTargetProcessor(true);
let plateProcessor: ReturnType<typeof createTargetProcessor> | undefined;

/** Native parsing retains authored links. Reference capture can request exact Plate grouping. */
export interface TargetParseOptions extends Partial<TargetParseLimits> {
  readonly preserveLinkedImages?: boolean;
}

class ParseLimit extends Error {
  constructor(readonly code: TargetParseIssue['code']) {
    super(`Target Markdown exceeds its ${code}.`);
  }
}

function limitsOf(options: Partial<TargetParseLimits>): TargetParseLimits {
  const limits = { ...DEFAULT_TARGET_PARSE_LIMITS, ...options };
  for (const key of ['maxBytes', 'maxNodes', 'maxDepth'] as const)
    if (
      !Number.isSafeInteger(limits[key]) ||
      limits[key] < 1 ||
      limits[key] > DEFAULT_TARGET_PARSE_LIMITS[key]
    )
      throw new TypeError(`Invalid target parse limit: ${key}`);
  return limits;
}

function checkBytes(markdown: string, limits: TargetParseLimits): void {
  if (
    markdown.length > limits.maxBytes ||
    new TextEncoder().encode(markdown).byteLength > limits.maxBytes
  )
    throw new ParseLimit('byte-limit');
}

/** Bound parser-owned data before recursive transforms or JSON cloning. */
function checkTree(root: unknown, limits: TargetParseLimits): void {
  const stack = [{ value: root, depth: 0 }];
  let count = 0;
  while (stack.length) {
    const { value, depth } = stack.pop()!;
    if (++count > limits.maxNodes) throw new ParseLimit('node-limit');
    if (depth > limits.maxDepth) throw new ParseLimit('depth-limit');
    if (value === null || typeof value !== 'object') continue;
    for (const [key, child] of Object.entries(value)) {
      // MDX expressions retain source text. Their derived compiler tree can contain BigInt/RegExp.
      if (key !== 'estree') stack.push({ value: child, depth: depth + 1 });
    }
  }
}

function freezeJson(value: JsonValue): void {
  if (value === null || typeof value !== 'object') return;
  for (const child of Object.values(value)) freezeJson(child);
  Object.freeze(value);
}

function normalized(value: JsonValue): JsonValue {
  if (value === null || typeof value !== 'object') return value;
  if (Array.isArray(value)) return value.map(normalized);
  const node = value as JsonObject;
  return Object.fromEntries(
    Object.entries(node)
      .filter(([key]) => key !== 'position' || typeof node.type !== 'string')
      .map(([key, child]) => [key, normalized(child)])
      .filter(
        ([key, child]) =>
          key !== 'data' ||
          child === null ||
          typeof child !== 'object' ||
          Object.keys(child).length > 0
      )
  );
}

function issueOf(cause: unknown): TargetParseIssue {
  if (cause instanceof ParseLimit) return { code: cause.code, message: cause.message };
  if (cause instanceof TypeError) return { code: 'invalid-input', message: cause.message };
  const error = cause as {
    message?: string;
    reason?: string;
    line?: number;
    column?: number;
    ruleId?: string;
    actual?: string;
    expected?: string | string[];
    note?: string;
  } | null;
  return {
    code: 'syntax',
    message: error?.reason ?? error?.message ?? String(cause),
    ...(Number.isSafeInteger(error?.line) ? { line: error!.line } : {}),
    ...(Number.isSafeInteger(error?.column) ? { column: error!.column } : {}),
    ...(error?.ruleId ? { ruleId: error.ruleId } : {}),
    ...(error?.actual || error?.expected || error?.note
      ? {
          details: Object.freeze([
            ...(error.actual ? [`actual=${error.actual}`] : []),
            ...(error.expected
              ? [
                  `expected=${Array.isArray(error.expected) ? error.expected.join(' | ') : error.expected}`,
                ]
              : []),
            ...(error.note ? [error.note] : []),
          ]),
        }
      : {}),
  };
}

/**
 * Parse a proposed target with the production Markdown plugins. Never construct or replace a DOCX tree.
 * Domain planners consume preserved MDX attributes and children. P11 owns lesson shape validation.
 * Apply this same function to both read Markdown and target Markdown before semantic comparison.
 */
export function parseTargetMarkdown(
  input: string,
  options: TargetParseOptions = {}
): TargetParseResult {
  let markdown: string | undefined;
  const failure = (issue: TargetParseIssue): TargetParseResult => ({
    success: false,
    code: 'invalid-input',
    error: issue.message,
    issues: Object.freeze([Object.freeze(issue)]),
    ...(markdown === undefined ? {} : { markdown }),
  });
  if (typeof input !== 'string')
    return failure({ code: 'invalid-input', message: 'Target Markdown must be a string.' });
  try {
    if (
      options.preserveLinkedImages !== undefined &&
      typeof options.preserveLinkedImages !== 'boolean'
    )
      throw new TypeError('Invalid linked image preservation mode.');
    const { preserveLinkedImages, ...limitOverrides } = options;
    const limits = limitsOf(limitOverrides);
    const parser =
      preserveLinkedImages === false
        ? (plateProcessor ??= createTargetProcessor(false))
        : processor;
    checkBytes(input, limits);
    const normalizedSource = normalizeTargetSource(input);
    markdown = normalizedSource.markdown;
    checkBytes(markdown, limits);
    const parsed = parser.parse(markdown);
    checkTree(parsed, limits);
    const transformed = parser.runSync(parsed);
    checkTree(transformed, limits);
    // Only parser-generated compiler caches are omitted. Raw expression source remains inert data.
    const ast = JSON.parse(
      JSON.stringify(transformed, (key, value) => (key === 'estree' ? undefined : value))
    ) as TargetAstRoot;
    const normalizedAst = normalized(ast) as TargetAstRoot;
    freezeJson(ast);
    freezeJson(normalizedAst);
    return {
      success: true,
      target: Object.freeze({
        sourceMarkdown: input,
        markdown,
        normalization: normalizedSource.normalization,
        ast,
        normalizedAst,
      }),
    };
  } catch (cause) {
    return failure(issueOf(cause));
  }
}

/**
 * Map one parsed AST range to original Markdown, retaining generated/rewritten status.
 * Use the parser-owned target and its own trace. Resolve native identities separately.
 * Rewritten tag interiors and UTF-16 surrogate splits have no source coordinate.
 */
export function targetSourceRangeAt(
  target: Pick<ParsedTargetMarkdown, 'markdown' | 'sourceMarkdown' | 'normalization'>,
  start: number,
  end: number
): TargetSourceRange | null {
  if (splitsSurrogate(target.markdown, start) || splitsSurrogate(target.markdown, end)) return null;
  const range = normalizationSourceRangeAt(target.normalization, start, end);
  if (
    !range ||
    range.end > target.sourceMarkdown.length ||
    splitsSurrogate(target.sourceMarkdown, range.start) ||
    splitsSurrogate(target.sourceMarkdown, range.end)
  )
    return null;
  return range;
}

/** Exact original boundary, including rewritten tag edges. This alone grants no native write authority. */
export function targetSourceOffsetAt(
  target: Pick<ParsedTargetMarkdown, 'markdown' | 'sourceMarkdown' | 'normalization'>,
  offset: number
): number | null {
  return targetSourceRangeAt(target, offset, offset)?.start ?? null;
}
