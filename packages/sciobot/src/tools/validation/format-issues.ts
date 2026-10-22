// Application-owned tool feedback. See fixtures/target-validation.json.
import type { ValidationIssue } from './schema-engine';

function formatIssueLines(issues: ValidationIssue[]): string[] {
  return issues.map((issue, i) => {
    const loc =
      issue.line !== undefined
        ? ` line ${issue.line}${issue.column !== undefined ? `:${issue.column}` : ''}`
        : '';
    return `${i + 1}. [${issue.code}]${loc}: ${issue.message}`;
  });
}

/**
 * Quote the rejected document around the first located issue so the agent
 * sees the exact orphaned tag instead of guessing from a line number. Line
 * numbers refer to the document as it would look WITH the rejected patch.
 */
function formatIssueContext(issues: ValidationIssue[], rejectedContent: string): string[] {
  const firstLocated = issues.find((issue) => issue.line !== undefined);
  if (!firstLocated?.line) return [];
  const lines = rejectedContent.split('\n');
  const start = Math.max(1, firstLocated.line - 3);
  const end = Math.min(lines.length, firstLocated.line + 3);
  const snippet: string[] = [
    '',
    'Context of the rejected result around the first issue (these lines were NOT applied):',
  ];
  for (let n = start; n <= end; n++) {
    const marker = n === firstLocated.line ? '>>>' : '   ';
    snippet.push(`${marker} ${String(n).padStart(4)} | ${lines[n - 1]}`);
  }
  return snippet;
}

/**
 * Tool-result error for a REJECTED edit: the patch would corrupt the
 * document, so it was NOT applied. Includes the offending region of the
 * WOULD-BE document so the agent can locate its own mistake (usually a tag
 * it deleted one half of, or a line range that swallowed a closing tag).
 */
export function formatEditRejection(
  issues: ValidationIssue[],
  rejectedContent?: string,
  hints?: string[],
  kind: 'syntax' | 'schema' = 'syntax'
): string {
  return [
    kind === 'schema'
      ? 'Edit REJECTED — the document parses correctly but does not satisfy the lesson-plan structure. It was NOT applied. Fix the listed schema requirements.'
      : 'Edit REJECTED — the resulting document is not valid MDX and was NOT applied.',
    ...(kind === 'syntax'
      ? [
          'The most common cause: old_string and new_string disagree about a tag pair (e.g. old_string swallowed a closing </activity> that new_string does not re-emit, or new_string left an opening tag unclosed). Keep BOTH halves of every tag pair consistently kept or removed.',
        ]
      : []),
    '',
    ...formatIssueLines(issues),
    ...(hints && hints.length > 0
      ? ['', 'How to fix your range:', ...hints.map((h) => `- ${h}`)]
      : []),
    ...(rejectedContent ? formatIssueContext(issues, rejectedContent) : []),
  ].join('\n');
}

/**
 * Tool-result warning appended to a SUCCESSFUL edit whose result violates
 * the lesson-plan schema. The edit is applied (incremental edits may pass
 * through transiently-invalid shapes), but the agent must restore a valid
 * document before finishing its turn.
 */
export function formatSchemaWarning(issues: ValidationIssue[]): string {
  return [
    'WARNING: the document now violates the lesson-plan schema. You MUST fix every issue below with further edits before ending your turn:',
    ...formatIssueLines(issues),
  ].join('\n');
}
