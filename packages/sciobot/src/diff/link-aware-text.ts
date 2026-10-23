import {
  findNode,
  paragraphOffsetIndex,
  paragraphTextOf,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { CanonicalEdit } from './plan';
import type { CompiledInlineTarget } from './inline-target';
import { previewTextEdits } from './insertion-marks';

/** Keep exact source hyperlink ownership when repeated characters make a lexical diff ambiguous. */
export function linkAwareBias(
  part: OoxmlPart,
  paragraphId: string,
  base: number,
  edits: readonly CanonicalEdit[],
  target: CompiledInlineTarget
): CanonicalEdit[] {
  const paragraph = findNode(part, paragraphId);
  if (paragraph?.kind !== 'paragraph' || !target.tokens.length) return [...edits];
  const offsets = paragraphOffsetIndex(paragraph),
    spans: { id: string; start: number; end: number }[] = [],
    stack: OoxmlNode[] = [paragraph];
  while (stack.length) {
    const node = stack.pop()!;
    if (node.kind === 'hyperlink') {
      const span = offsets.spanOf(node);
      if (span) spans.push({ id: node.id, ...span });
    } else if (node.kind !== 'textValue') stack.push(...node.children);
  }
  if (!spans.length) return [...edits];
  const linked: boolean[] = [],
    wanted = new Map<string, { start: number; end: number }>();
  let offset = base;
  for (const token of target.tokens) {
    if (token.kind !== 'text') continue;
    for (let i = 0; i < token.value.length; i++) linked.push(token.hyperlink !== null);
    if (token.hyperlink) {
      const key = token.hyperlink.join('/');
      wanted.set(key, {
        start: wanted.get(key)?.start ?? offset,
        end: offset + token.value.length,
      });
    }
    offset += token.value.length;
  }
  const wantedRanges = [...wanted.values()];
  const exactRanges = new Set(wantedRanges.map((range) => `${range.start}/${range.end}`));
  const bias = (operations: readonly CanonicalEdit[]) =>
    operations.map((edit) => {
      if (edit.kind !== 'text' || edit.before || !edit.after) return edit;
      const atStart = spans.some((span) => span.start === edit.start),
        atEnd = spans.some((span) => span.end === edit.start);
      if (!atStart && !atEnd) return edit;
      const shift = operations.reduce(
        (total, other) =>
          other.kind === 'text' && other.start < edit.start
            ? total + other.after.length - (other.end - other.start)
            : total,
        0
      );
      const start = edit.start + shift;
      const inside = linked[start - base] ?? false;
      const owner =
        inside && !/[\n\t\u001e\u001f]/.test(edit.after)
          ? spans.find((span) => {
              if (span.start !== edit.start && span.end !== edit.start) return false;
              const deltaBefore = (position: number) =>
                operations.reduce(
                  (total, other) =>
                    other !== edit && other.kind === 'text' && other.start < position
                      ? total + other.after.length - (other.end - other.start)
                      : total,
                  0
                );
              const from = span.start + deltaBefore(span.start),
                to = span.end + deltaBefore(span.end) + edit.after.length;
              return exactRanges.has(`${from}/${to}`);
            })
          : undefined;
      return {
        ...edit,
        bias: atStart === inside ? ('right' as const) : ('left' as const),
        ...(owner && owner.end === edit.start ? { insideHyperlink: owner.id } : {}),
      };
    });
  const inspect = (operations: readonly CanonicalEdit[]) => {
    const result = previewTextEdits(part, bias(operations));
    if (!result.success) return null;
    const node = findNode(result.part, paragraphId);
    if (node?.kind !== 'paragraph') return null;
    const index = paragraphOffsetIndex(node);
    return {
      text: paragraphTextOf(result.part, paragraphId),
      score: spans.filter((span) => {
        const retained = findNode(result.part, span.id),
          current = retained && index.spanOf(retained);
        return current && exactRanges.has(`${current.start}/${current.end}`);
      }).length,
    };
  };
  let operations = [...edits],
    current = inspect(operations);
  const native = paragraphTextOf(part, paragraphId)!;
  for (let i = 0; current && current.score < spans.length && i < operations.length; i++) {
    const edit = operations[i]!;
    if (edit.kind !== 'text' || edit.before || !edit.after) continue;
    const boundaries = [...new Set(spans.flatMap((span) => [span.start, span.end]))].sort(
      (a, b) => Math.abs(a - edit.start) - Math.abs(b - edit.start) || a - b
    );
    for (const at of boundaries) {
      if (at === edit.start) continue;
      const lo = Math.min(at, edit.start),
        hi = Math.max(at, edit.start),
        between = native.slice(lo, hi);
      if (between + edit.after !== edit.after + between) continue;
      if (
        operations.some(
          (other, j) => j !== i && other.kind === 'text' && other.start <= hi && other.end >= lo
        )
      )
        continue;
      const candidate = operations.map((other, j) =>
        j === i ? { ...edit, start: at, end: at } : other
      );
      const next = inspect(candidate);
      if (next && next.text === current.text && next.score > current.score) {
        operations = candidate;
        current = next;
        break;
      }
    }
  }
  return bias(operations);
}
