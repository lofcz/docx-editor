import type { NativeEditFailure } from '../contracts/document';
import { jsonDataProblem } from '../data/json';
import {
  parseTargetMarkdown,
  type ParsedTargetMarkdown,
  type TargetParseLimits,
} from '../projection/parse-target';
import { validateMdxTree, type JsxNode, type ValidationIssue } from './validation/schema-engine';
import { lessonPlanSchema } from './validation/lesson-plan-schema';
import { analyzeEditTagBalance } from './validation/tag-balance';
import { formatEditRejection, formatSchemaWarning } from './validation/format-issues';

export interface TargetValidationOptions {
  readonly materialKind: 'document' | 'lesson_plan' | 'curriculum';
  readonly operation: 'edit' | 'set';
  /** Exact requested strings supply hints only. They never determine syntax acceptance. */
  readonly replacement?: { readonly oldString: string; readonly newString: string };
  readonly limits?: Partial<TargetParseLimits>;
}
export type TargetValidationIssue = Readonly<ValidationIssue>;
export type TargetValidationResult =
  | {
      readonly success: true;
      readonly target: ParsedTargetMarkdown;
      readonly warnings: readonly TargetValidationIssue[];
      readonly warning?: string;
    }
  | (NativeEditFailure & {
      readonly stage: 'input' | 'syntax' | 'schema';
      readonly issues: readonly TargetValidationIssue[];
      readonly hints: readonly string[];
    });

const freezeIssues = (issues: readonly ValidationIssue[]) =>
  Object.freeze(issues.map((issue) => Object.freeze({ ...issue })));
const inputFailure = (message: string, code = 'invalid-input'): TargetValidationResult =>
  Object.freeze({
    success: false,
    code: 'invalid-input',
    error: message,
    stage: 'input',
    issues: freezeIssues([{ code, message }]),
    hints: Object.freeze([]),
  });

/** Pure preflight. No native document, editor, apply callback, or persistent state enters this API. */
export function validateTargetMarkdown(
  input: string,
  options: TargetValidationOptions
): TargetValidationResult {
  const problem = jsonDataProblem(options, { maxDepth: 5, maxNodes: 32 });
  if (problem) return inputFailure(problem);
  if (
    !options ||
    typeof options !== 'object' ||
    Array.isArray(options) ||
    !['document', 'lesson_plan', 'curriculum'].includes(options.materialKind) ||
    !['edit', 'set'].includes(options.operation) ||
    Object.keys(options).some(
      (key) => !['materialKind', 'operation', 'replacement', 'limits'].includes(key)
    )
  )
    return inputFailure('Invalid target validation options.');
  if (options.replacement !== undefined) {
    const replacement = options.replacement;
    if (
      !replacement ||
      typeof replacement.oldString !== 'string' ||
      typeof replacement.newString !== 'string' ||
      Object.keys(replacement).some((key) => key !== 'oldString' && key !== 'newString') ||
      replacement.oldString.length > 2 * 1024 * 1024 ||
      replacement.newString.length > 2 * 1024 * 1024
    )
      return inputFailure('Invalid or excessive replacement hint input.');
  }
  if (
    options.limits !== undefined &&
    (!options.limits ||
      typeof options.limits !== 'object' ||
      Array.isArray(options.limits) ||
      Object.keys(options.limits).some(
        (key) => !['maxBytes', 'maxNodes', 'maxDepth'].includes(key)
      ))
  )
    return inputFailure('Invalid target parse limits.');
  const parsed = parseTargetMarkdown(input, options.limits);
  const reject = (
    stage: 'syntax' | 'schema',
    issues: readonly ValidationIssue[],
    markdown: string
  ): TargetValidationResult => {
    const hints =
      stage === 'syntax' && options.replacement
        ? analyzeEditTagBalance(options.replacement.oldString, options.replacement.newString)
        : [];
    return Object.freeze({
      success: false,
      code: 'invalid-input',
      stage,
      error: formatEditRejection([...issues], markdown, hints, stage),
      issues: freezeIssues(issues),
      hints: Object.freeze(hints),
    });
  };
  if (!parsed.success) {
    if (parsed.issues.some((issue) => issue.code !== 'syntax'))
      return inputFailure(parsed.error, parsed.issues[0]!.code);
    return reject(
      'syntax',
      parsed.issues.map((issue) => ({
        code: issue.ruleId ?? 'mdx-parse-error',
        message: [issue.message, ...(issue.details ?? [])].join(' — '),
        ...(issue.line === undefined ? {} : { line: issue.line }),
        ...(issue.column === undefined ? {} : { column: issue.column }),
      })),
      parsed.markdown ?? input
    );
  }
  if (!parsed.target.markdown.trim())
    return reject(
      'syntax',
      [{ code: 'empty-document', message: 'Document is empty.' }],
      parsed.target.markdown
    );
  const issues =
    options.materialKind === 'lesson_plan'
      ? validateMdxTree(parsed.target.ast as unknown as JsxNode, lessonPlanSchema).issues
      : [];
  if (issues.length && options.operation === 'set')
    return reject('schema', issues, parsed.target.markdown);
  return Object.freeze({
    success: true,
    target: parsed.target,
    warnings: freezeIssues(issues),
    ...(issues.length ? { warning: formatSchemaWarning(issues) } : {}),
  });
}
