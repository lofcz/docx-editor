// Map case over the selected text before distributing its letters back to native runs.
// Whole-language lowercasing supplies contextual forms such as the Greek final sigma.
import type { EditorCommand, ParagraphTextProofing } from '../contracts/editor.ts';

export type CaseMode = Extract<EditorCommand, { type: 'changeCase' }>['mode'];

let graphemes: Intl.Segmenter | undefined;
let words: Intl.Segmenter | undefined;
let sentences: Intl.Segmenter | undefined;
const CASED = /\p{Cased}/u;

export function caseLetters(
  text: string,
  start: number,
  end: number,
  mode: CaseMode,
  languages: readonly ParagraphTextProofing[]
): { start: number; end: number; text: string }[] | null {
  if (typeof Intl.Segmenter !== 'function') return null;
  graphemes ??= new Intl.Segmenter('und', { granularity: 'grapheme' });
  const input = text.slice(start, end);
  const initial = new Set<number>();
  if (mode === 'capitalizeWords' || mode === 'sentence') {
    const boundaries =
      mode === 'sentence'
        ? (sentences ??= new Intl.Segmenter('und', { granularity: 'sentence' }))
        : (words ??= new Intl.Segmenter('und', { granularity: 'word' }));
    for (const { segment, index } of boundaries.segment(input)) {
      const first = segment.search(CASED);
      if (first >= 0) initial.add(index + first);
    }
  }
  const changes: { start: number; end: number; text: string }[] = [];
  const boundaries = graphemes.segment(text);
  if (
    boundaries.containing(start)?.index !== start ||
    (end !== text.length && boundaries.containing(end)?.index !== end)
  )
    return null;
  for (let i = 0; i < languages.length; i++) {
    const span = languages[i]!;
    const lang = span.lang ?? 'und';
    let limit = span.acceptedStart + span.length;
    // Proofing can split a word without changing its casing language.
    while (i + 1 < languages.length) {
      const next = languages[i + 1]!;
      if (next.acceptedStart !== limit || (next.lang ?? 'und') !== lang) break;
      limit += next.length;
      i++;
    }
    const from = Math.max(start, span.acceptedStart);
    const to = Math.min(end, limit);
    if (from >= to) continue;
    const part = text.slice(from, to);
    let lower: string;
    try {
      lower = part.toLocaleLowerCase(lang);
    } catch {
      return null;
    }
    let offset = 0;
    for (const { segment, index } of graphemes.segment(part)) {
      const upper = segment.toLocaleUpperCase(lang);
      const size = segment.toLocaleLowerCase(lang).length;
      const small = lower.slice(offset, offset + size);
      offset += size;
      const useUpper =
        mode === 'uppercase' ||
        initial.has(from - start + index) ||
        (mode === 'toggle' && segment !== upper);
      const replacement = useUpper ? upper : small;
      if (replacement !== segment)
        changes.push({
          start: from + index,
          end: from + index + segment.length,
          text: replacement,
        });
    }
    if (offset !== lower.length) return null;
  }
  return changes;
}
