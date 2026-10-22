// Pure application normalization. Reference hashes live in target-parsing.json.
import { unified } from 'unified';
import remarkParse from 'remark-parse';
import {
  normalizationStep,
  type NormalizationEdit,
  type TargetNormalizationMap,
} from './target-source-map';

/** Match the editor's HTML <br> support without rewriting code examples. */
export function normalizeMdxLineBreaks(
  markdown: string,
  capture?: (edit: NormalizationEdit) => void
): string {
  type Node = {
    type: string;
    children?: Node[];
    position?: { start: { offset?: number }; end: { offset?: number } };
  };
  const root = unified().use(remarkParse).parse(markdown) as Node;
  const replacements: { start: number; end: number; text: string }[] = [];
  const visit = (node: Node) => {
    if (node.type === 'html') {
      const start = node.position?.start.offset;
      const end = node.position?.end.offset;
      if (start != null && end != null) {
        const original = markdown.slice(start, end);
        const text = original.replace(/<br\s*>/gi, (match, offset: number) => {
          capture?.({ start: start + offset, end: start + offset + match.length, text: '<br />' });
          return '<br />';
        });
        if (text !== original) replacements.push({ start, end, text });
      }
    }
    for (const child of node.children ?? []) visit(child);
  };
  visit(root);
  for (const { start, end, text } of replacements.sort((a, b) => b.start - a.start)) {
    markdown = markdown.slice(0, start) + text + markdown.slice(end);
  }
  return markdown;
}

const ORDERED_ITEM = /^ {0,3}(\d{1,3})[.)][ \t]+\S/;
const LIST_ITEM = /^ {0,3}(?:\d{1,9}[.)]|[-+*])(?:[ \t]|$)/;
const FENCE = /^ {0,3}(```|~~~)/;
const MATH_FENCE = /^ {0,3}\$\$/;
const TAG_LINE = /^\s*<[^>]*>\s*$/;

/**
 * True when the next item line of the same block carries the following number,
 * so the line opens a real numbered list rather than being a sentence that
 * wrapped before a number ("začala\n2. světová válka", "dne\n15. května").
 */
function continuesNumbering(lines: readonly string[], from: number, start: number): boolean {
  for (let index = from; index < lines.length; index += 1) {
    const line = lines[index] ?? '';
    if (line.trim() === '' || TAG_LINE.test(line) || FENCE.test(line)) return false;
    const ordered = ORDERED_ITEM.exec(line);
    if (ordered) return Number(ordered[1]) === start + 1;
  }
  return false;
}

/**
 * Generated plans continue a numbered list right under a label
 * ("Logaritmický tvar na mocninu:\n6. …\n7. …"). CommonMark lets only a list
 * that starts at 1 interrupt a paragraph, so items 6–10 were glued onto the
 * label as one line. A blank line before such an item makes it a list again,
 * but only when the next item continues the numbering: a paragraph line that
 * merely starts with a number stays text (a lone "6. …" under a label does too).
 * Lines in a block that already holds list items are left alone, so tight
 * lists and their continuation lines keep their meaning.
 */
export function separateInterruptingLists(
  markdown: string,
  capture?: (edit: NormalizationEdit) => void
): string {
  const lines = markdown.split('\n');
  const out: string[] = [];
  let fence: RegExp | null = null;
  let blockHasList = false;
  let previous = '';
  let sourceOffset = 0;
  for (const [lineIndex, line] of lines.entries()) {
    const lineStart = sourceOffset;
    sourceOffset += line.length + 1;
    if (fence) {
      if (fence.test(line)) fence = null;
      out.push(line);
      previous = line;
      continue;
    }
    if (FENCE.test(line)) fence = FENCE;
    // A one-line `$$…$$` block opens and closes on the same line.
    else if (MATH_FENCE.test(line) && !line.trim().slice(2).includes('$$')) fence = MATH_FENCE;
    if (fence) {
      out.push(line);
      previous = line;
      continue;
    }
    if (line.trim() === '' || TAG_LINE.test(line)) {
      blockHasList = false;
    } else {
      const ordered = ORDERED_ITEM.exec(line);
      if (
        ordered &&
        Number(ordered[1]) !== 1 &&
        !blockHasList &&
        previous.trim() !== '' &&
        !TAG_LINE.test(previous) &&
        continuesNumbering(lines, lineIndex + 1, Number(ordered[1]))
      ) {
        out.push('');
        capture?.({ start: lineStart, end: lineStart, text: '\n' });
      }
      if (LIST_ITEM.test(line)) blockHasList = true;
    }
    out.push(line);
    previous = line;
  }
  return out.join('\n');
}

/** Keep each pass's exact edits. AST locations address the final string. */
export function normalizeTargetSource(source: string): {
  readonly markdown: string;
  readonly normalization: TargetNormalizationMap;
} {
  const breaks: NormalizationEdit[] = [],
    lists: NormalizationEdit[] = [];
  const intermediate = normalizeMdxLineBreaks(source, (edit) => breaks.push(edit));
  const markdown = separateInterruptingLists(intermediate, (edit) => lists.push(edit));
  return Object.freeze({
    markdown,
    normalization: Object.freeze({
      unit: 'utf16-code-unit' as const,
      steps: Object.freeze([
        normalizationStep(source.length, intermediate.length, breaks),
        normalizationStep(intermediate.length, markdown.length, lists),
      ]),
    }),
  });
}
