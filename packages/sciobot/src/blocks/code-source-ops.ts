import { diffChars } from 'diff';
import { trackedReplacementLanding, type TreeDocOp } from '@docx-editor.dev/core/store';
import type { nativeCodeSource } from './code-native-source';

interface SourceChange {
  start: number;
  end: number;
  insert: string;
}

/** Preserve unchanged spans while bounding source comparison work. */
function changes(before: string, after: string): SourceChange[] {
  const parts = diffChars(before, after, { timeout: 100, maxEditLength: 10000 });
  if (!parts) throw Error('The code change is too large. Apply smaller source edits.');
  const edits: SourceChange[] = [];
  let offset = 0;
  let current: SourceChange | undefined;
  for (const part of parts) {
    if (!part.added && !part.removed) {
      offset += part.value.length;
      current = undefined;
      continue;
    }
    if (!current) {
      current = { start: offset, end: offset, insert: '' };
      edits.push(current);
    }
    if (part.removed) {
      offset += part.value.length;
      current.end = offset;
    } else current.insert += part.value;
  }
  return edits;
}

/** Address original paragraphs in reverse source order. Inserted lines need no regenerated IDs. */
export function codeSourceOps(
  source: NonNullable<ReturnType<typeof nativeCodeSource>>,
  value: string,
  revision?: Extract<TreeDocOp, { op: 'insertText' }>['revision']
): readonly TreeDocOp[] {
  const ops: TreeDocOp[] = [];
  const attribution = revision ? { revision } : {};
  const locate = (offset: number) => {
    for (let row = 0; row < source.lines.length; row++) {
      const length = source.lines[row]!.length;
      if (offset <= length) return { row, offset };
      offset -= length + 1;
    }
    throw Error('The source position is invalid.');
  };
  for (const change of changes(source.value, value).reverse()) {
    let start = locate(change.start),
      end = locate(change.end),
      inserted = change.insert;
    // A full line before an existing line must retain that following line's identity.
    if (
      start.row > 0 &&
      start.offset === 0 &&
      change.start === change.end &&
      inserted.endsWith('\n')
    ) {
      start = { row: start.row - 1, offset: source.lines[start.row - 1]!.length };
      end = start;
      inserted = `\n${inserted.slice(0, -1)}`;
    }
    const paragraphId = source.paragraphs[start.row]!.id;
    for (let row = start.row; row <= end.row; row++) {
      const from = row === start.row ? start.offset : 0;
      const to = row === end.row ? end.offset : source.lines[row]!.length;
      if (to > from)
        ops.push({
          op: 'deleteText',
          paragraphId: source.paragraphs[row]!.id,
          start: from,
          end: to,
          ...attribution,
        });
      if (row > start.row)
        ops.push(
          revision
            ? { op: 'proposeParagraphMerge', paragraphId: source.paragraphs[row]!.id, revision }
            : { op: 'joinParagraphs', firstId: paragraphId, secondId: source.paragraphs[row]!.id }
        );
    }
    let offset = revision
      ? trackedReplacementLanding(
          source.paragraphs[start.row]!,
          start.offset,
          start.row === end.row ? end.offset : source.lines[start.row]!.length,
          revision.author,
          revision.date
        )
      : start.offset;
    const boundaries: number[] = [];
    for (const token of inserted.split(/(\n|\t)/)) {
      if (!token) continue;
      if (token === '\n') boundaries.push(offset);
      else {
        ops.push(
          token === '\t'
            ? { op: 'insertTab', paragraphId, offset, ...attribution }
            : { op: 'insertText', paragraphId, offset, text: token, ...attribution }
        );
        offset += token.length;
      }
    }
    if (boundaries.length)
      ops.push({ op: 'splitParagraphMany', paragraphId, offsets: boundaries, ...attribution });
    if (ops.length > 20000) throw Error('The code edit exceeds the operation limit.');
  }
  return ops;
}
