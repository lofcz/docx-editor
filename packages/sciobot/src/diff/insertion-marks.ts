import {
  applyTreeOp,
  type FormattingDisplayMode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import type { InlineTargetMark } from './inline-target';
import { directParagraphTextOps } from './paragraph-text';
import type { CanonicalEdit } from './plan';
import { planRunFormat } from './run-format';
import { inlineCodeRunChanges } from '../native/inline-code';
import { jsonDataKey } from './target-key';

export type MarkRun = { value: string; marks: readonly InlineTargetMark[] };
const MARK_PROPERTIES: Readonly<Record<string, string>> = {
  bold: 'b',
  italic: 'i',
  strike: 'strike',
};

/** Marks of the character beside `offset`, read in the direction native insertion inherits. */
export function markAt(
  runs: readonly MarkRun[],
  offset: number,
  bias: 'left' | 'right'
): readonly InlineTargetMark[] {
  const length = runs.reduce((total, run) => total + run.value.length, 0);
  const at = bias === 'left' ? Math.max(0, offset - 1) : Math.min(length - 1, offset);
  let position = 0;
  for (const run of runs) {
    position += run.value.length;
    if (at < position) return run.marks;
  }
  return [];
}

export interface MarkDeltaSource {
  readonly revision: RevisionToken;
  readonly part: OoxmlPart;
  readonly paragraphId: string;
  /** Native offset of the first model character, after any owned task marker. */
  readonly base: number;
  readonly length: number;
  readonly displayMode: FormattingDisplayMode;
  /** The document's Hyperlink character style, restored on link text that leaves code. */
  readonly hyperlinkStyle?: string | null;
  readonly codeStyles?: readonly string[];
}

/**
 * Plan run-format intentions that turn `current` mark runs into `target` mark runs over the same
 * text. Old values come from the supplied native part, so callers pass the part the executor will
 * see when these intentions apply.
 */
export function planMarkDeltas(
  source: MarkDeltaSource,
  current: readonly MarkRun[],
  target: readonly MarkRun[]
): { success: true; operations: readonly CanonicalEdit[] } | NativeEditFailure {
  const positions = new Set([0, source.length]);
  for (const runs of [current, target]) {
    let offset = 0;
    for (const run of runs) {
      offset += run.value.length;
      positions.add(offset);
    }
  }
  const boundaries = [...positions].sort((a, b) => a - b),
    operations: CanonicalEdit[] = [];
  for (let i = 0; i + 1 < boundaries.length; i++) {
    const start = boundaries[i]!,
      end = boundaries[i + 1]!;
    if (start === end) continue;
    const old = markAt(current, start, 'right'),
      next = markAt(target, start, 'right');
    const changed = [...new Set([...old, ...next])].filter(
      (mark) => old.includes(mark) !== next.includes(mark)
    );
    if (!changed.length) continue;
    const segment: CanonicalEdit[] = [];
    const toggles = changed.filter((mark) => mark !== 'code');
    if (toggles.length) {
      const result = planRunFormat(
        {
          revision: source.revision,
          part: source.part,
          paragraphId: source.paragraphId,
          start: source.base + start,
          end: source.base + end,
          displayMode: source.displayMode,
        },
        {
          properties: toggles.map((mark) => ({
            localName: MARK_PROPERTIES[mark]!,
            attributes: { val: next.includes(mark) ? '1' : '0' },
          })),
        }
      );
      if (!result.success) return result;
      segment.push(...result.plan.operations);
    }
    if (changed.includes('code')) {
      // Code is the run's own character style (N18), planned as an explicit `rStyle` change
      // on the same native run slices the toggle deltas use.
      const code = inlineCodeRunChanges(
        source.part,
        {
          paragraphId: source.paragraphId,
          start: source.base + start,
          end: source.base + end,
          code: next.includes('code'),
        },
        {
          displayMode: source.displayMode,
          hyperlinkStyle: source.hyperlinkStyle ?? null,
          codeStyles: source.codeStyles,
        }
      );
      if (!code.success) return code;
      for (const change of code.changes) {
        if (change.targetRunIds)
          return {
            success: false,
            code: 'unsupported-operation',
            error: 'Field-result code formatting requires explicit run ownership.',
          };
        const rStyle = {
          name: 'rStyle',
          before:
            change.before === undefined
              ? { present: false as const }
              : { present: true as const, value: { val: change.before } },
          after:
            change.after === null
              ? { present: false as const }
              : { present: true as const, value: { val: change.after } },
        };
        const index = segment.findIndex(
          (edit) =>
            edit.kind === 'format' &&
            edit.target.kind === 'run' &&
            edit.target.start === change.start &&
            edit.target.end === change.end
        );
        const merged = index >= 0 ? segment[index] : undefined;
        if (merged?.kind === 'format')
          segment[index] = { ...merged, changes: [...merged.changes, rStyle] };
        else
          segment.push({
            kind: 'format',
            partName: source.part.name,
            target: {
              kind: 'run',
              paragraphId: source.paragraphId,
              start: change.start,
              end: change.end,
            },
            changes: [rStyle],
          });
      }
    }
    // Projection tokens can split identical deltas at whitespace boundaries. Keep one
    // continuous intention so native formatting does not split unchanged source text there.
    for (const edit of segment) {
      const previous = operations.at(-1);
      if (
        edit.kind === 'format' &&
        edit.target.kind === 'run' &&
        previous?.kind === 'format' &&
        previous.target.kind === 'run' &&
        previous.partName === edit.partName &&
        previous.target.paragraphId === edit.target.paragraphId &&
        previous.target.end === edit.target.start &&
        jsonDataKey(previous.changes) === jsonDataKey(edit.changes)
      )
        operations[operations.length - 1] = {
          ...previous,
          target: { ...previous.target, end: edit.target.end },
        };
      else operations.push(edit);
    }
  }
  return { success: true, operations };
}

/** Apply planned text edits to an immutable copy, as a direct executor would before later edits. */
export function previewTextEdits(
  part: OoxmlPart,
  edits: readonly CanonicalEdit[]
): { success: true; part: OoxmlPart } | NativeEditFailure {
  let preview = part;
  for (const edit of edits) {
    if (edit.kind !== 'text')
      return {
        success: false,
        code: 'invalid-input',
        error: 'Unexpected paragraph text operation.',
      };
    for (const op of directParagraphTextOps(edit)) {
      const result = applyTreeOp(preview, op);
      if (!result.ok)
        return {
          success: false,
          code: 'unsupported-operation',
          error: `The inserted text preview was refused: ${result.reason}.`,
        };
      preview = result.part;
    }
  }
  return { success: true, part: preview };
}
