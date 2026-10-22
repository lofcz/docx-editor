import { applyPatch, createTwoFilesPatch, diffArrays, type StructuredPatch } from 'diff';

/**
 * Application find/replace engine, ported from Sciobot’s shared pure implementation.
 * See fixtures/find-replace.json for reference hashes and expected behavior.
 */

const LEFT_SINGLE = '‘';
const RIGHT_SINGLE = '’';
const LEFT_DOUBLE = '“';
const RIGHT_DOUBLE = '”';
const LEFT_DOUBLE_LOW = '„'; // Czech/German opening double quote (U+201E)
const LEFT_SINGLE_LOW = '‚'; // Czech/German opening single quote (U+201A)

/** Curly → straight, for matching only (never written back verbatim). */
function normalizeQuotes(s: string): string {
  return s
    .replaceAll(LEFT_SINGLE, "'")
    .replaceAll(RIGHT_SINGLE, "'")
    .replaceAll(LEFT_SINGLE_LOW, "'")
    .replaceAll(LEFT_DOUBLE, '"')
    .replaceAll(RIGHT_DOUBLE, '"')
    .replaceAll(LEFT_DOUBLE_LOW, '"');
}

/** Decode the small set of HTML entities a model tends to emit. */
function decodeEntities(s: string): string {
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => {
      const code = Number.parseInt(hex, 16);
      return Number.isFinite(code) ? String.fromCodePoint(code) : _;
    })
    .replace(/&#(\d+);/g, (_, dec) => {
      const code = Number.parseInt(dec, 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : _;
    })
    .replaceAll('&amp;', '&')
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&nbsp;', ' ');
}

/** Combined normal form used for typography/entity matching. */
function normalizeForMatch(s: string): string {
  return decodeEntities(normalizeQuotes(s));
}

export type FindAndReplaceResult =
  | { success: true; newContent: string }
  | { success: false; error: string };

/**
 * Locate `oldString` in `content`. Returns the exact substring of `content`
 * that matched (which may differ from `oldString` when typography/entities
 * were normalized), or null.
 */
function findActual(content: string, oldString: string): string | null {
  // Layer 1: exact.
  if (content.includes(oldString)) return oldString;

  // Layer 2: typography + entity normalization. The lengths of the normalized
  // and original strings can differ (entity decoding shrinks), so we map the
  // normalized index back to the original via a per-char scan.
  const nContent = normalizeForMatch(content);
  const nOld = normalizeForMatch(oldString);
  const idx = nContent.indexOf(nOld);
  if (idx === -1) return null;

  // Walk the original content, tracking the normalized cursor, until we've
  // covered the whole normalized match. Each original unit (single char, or a
  // multi-char HTML entity) contributes exactly one normalized char.
  let origStart = -1;
  let ni = 0;
  for (let oi = 0; oi < content.length; oi++) {
    let origLen = 1;
    const rest = content.slice(oi);
    const entMatch = /^&#x[0-9a-fA-F]+;|^&#\d+;|^&(amp|lt|gt|quot|apos|nbsp);/.exec(rest);
    if (entMatch) origLen = entMatch[0].length;

    if (ni === idx && origStart === -1) origStart = oi;
    ni += 1;
    if (origStart !== -1 && ni === idx + nOld.length) {
      return content.substring(origStart, oi + origLen);
    }
    if (origLen > 1) oi += origLen - 1;
  }
  return null;
}

/**
 * When oldString matched via typography normalization (curly quotes in the
 * document, straight quotes from the model), restyle newString to the
 * document's quote style so the edit preserves typography instead of
 * flattening it. Mirrors Claude Code's preserveQuoteStyle.
 */
function preserveQuoteStyle(oldString: string, actualOld: string, newString: string): string {
  if (normalizeQuotes(oldString) === normalizeQuotes(actualOld) && oldString !== actualOld) {
    // Only quotes differed — detect which curly forms the document uses.
    const usesDouble =
      actualOld.includes(LEFT_DOUBLE) ||
      actualOld.includes(RIGHT_DOUBLE) ||
      actualOld.includes(LEFT_DOUBLE_LOW);
    const usesSingle =
      actualOld.includes(LEFT_SINGLE) ||
      actualOld.includes(RIGHT_SINGLE) ||
      actualOld.includes(LEFT_SINGLE_LOW);
    if (!usesDouble && !usesSingle) return newString;

    const isOpening = (chars: string[], i: number): boolean => {
      if (i === 0) return true;
      const p = chars[i - 1];
      return (
        p === ' ' ||
        p === '\t' ||
        p === '\n' ||
        p === '\r' ||
        p === '(' ||
        p === '[' ||
        p === '{' ||
        p === '—' ||
        p === '–'
      );
    };

    const chars = [...newString];
    const out: string[] = [];
    for (let i = 0; i < chars.length; i++) {
      const c = chars[i];
      if (usesDouble && c === '"') {
        // Czech typography opens low („) and closes high (“).
        out.push(isOpening(chars, i) ? LEFT_DOUBLE_LOW : LEFT_DOUBLE);
      } else if (usesSingle && c === "'") {
        const prev = i > 0 ? chars[i - 1] : undefined;
        const next = i < chars.length - 1 ? chars[i + 1] : undefined;
        const contraction =
          prev !== undefined && next !== undefined && /\p{L}/u.test(prev) && /\p{L}/u.test(next);
        out.push(contraction ? RIGHT_SINGLE : isOpening(chars, i) ? LEFT_SINGLE_LOW : LEFT_SINGLE);
      } else {
        out.push(c);
      }
    }
    return out.join('');
  }
  return newString;
}

/**
 * Replace `search` with `replacement` (function form so `$&`/`$$`/`$1` in the
 * replacement are NOT treated as special `String.replace` patterns — a real
 * corruption bug when the model's new_string contains math/regex/price text).
 *
 * On deletion (replacement === "") where the match isn't followed by a
 * newline, consume the trailing newline too, so deleting a line doesn't leave
 * a stray blank line behind.
 */
function applyReplace(
  content: string,
  search: string,
  replacement: string,
  replaceAll: boolean
): string {
  const fn = () => replacement;
  if (replacement !== '') {
    return replaceAll ? content.replaceAll(search, fn) : content.replace(search, fn);
  }
  const stripTrailingNewline = !search.endsWith('\n') && content.includes(search + '\n');
  const target = stripTrailingNewline ? search + '\n' : search;
  return replaceAll ? content.replaceAll(target, fn) : content.replace(target, fn);
}

/**
 * Codex `seek_sequence` punctuation table
 * (codex-rs/apply-patch/src/seek_sequence.rs). Used only as jsdiff compareLine.
 */
function codexNormalise(s: string): string {
  let out = '';
  for (const c of s.trim()) {
    const code = c.charCodeAt(0);
    if (
      code === 0x2010 ||
      code === 0x2011 ||
      code === 0x2012 ||
      code === 0x2013 ||
      code === 0x2014 ||
      code === 0x2015 ||
      code === 0x2212
    ) {
      out += '-';
    } else if (code === 0x2018 || code === 0x2019 || code === 0x201a || code === 0x201b) {
      out += "'";
    } else if (code === 0x201c || code === 0x201d || code === 0x201e || code === 0x201f) {
      out += '"';
    } else if (
      code === 0x00a0 ||
      code === 0x2002 ||
      code === 0x2003 ||
      code === 0x2004 ||
      code === 0x2005 ||
      code === 0x2006 ||
      code === 0x2007 ||
      code === 0x2008 ||
      code === 0x2009 ||
      code === 0x200a ||
      code === 0x202f ||
      code === 0x205f ||
      code === 0x3000
    ) {
      out += ' ';
    } else {
      out += c;
    }
  }
  return out;
}

type LineEq = (fileLine: string, patternLine: string) => boolean;

const COMPARE_PASSES: LineEq[] = [
  (file, pat) => file === pat,
  (file, pat) => file.trimEnd() === pat.trimEnd(),
  (file, pat) => file.trim() === pat.trim(),
  (file, pat) => codexNormalise(file) === codexNormalise(pat),
];

function collectSeekHits(lines: string[], pattern: string[], eq: LineEq): number[] {
  if (pattern.length === 0 || pattern.length > lines.length) return [];
  const hits: number[] = [];
  const last = lines.length - pattern.length;
  for (let i = 0; i <= last; i++) {
    let ok = true;
    for (let j = 0; j < pattern.length; j++) {
      if (!eq(lines[i + j]!, pattern[j]!)) {
        ok = false;
        break;
      }
    }
    if (ok) hits.push(i);
  }
  return hits;
}

/** Whole-span hunk so jsdiff cannot drop shared lines as context. */
function searchReplacePatch(oldString: string, newString: string): StructuredPatch {
  const oldLines = oldString.split('\n');
  const newLines = newString.split('\n');
  return {
    oldFileName: 'old_string',
    newFileName: 'new_string',
    oldHeader: '',
    newHeader: '',
    hunks: [
      {
        oldStart: 1,
        oldLines: oldLines.length,
        newStart: 1,
        newLines: newLines.length,
        lines: [...oldLines.map((line) => `-${line}`), ...newLines.map((line) => `+${line}`)],
      },
    ],
  };
}

function applyWithCompare(
  content: string,
  oldString: string,
  newString: string
): { kind: 'ok'; content: string } | { kind: 'ambiguous'; count: number } | { kind: 'none' } {
  const lines = content.split('\n');
  const pattern = oldString.split('\n');
  for (const eq of COMPARE_PASSES) {
    const hits = collectSeekHits(lines, pattern, eq);
    if (hits.length > 1) return { kind: 'ambiguous', count: hits.length };
    if (hits.length !== 1) continue;
    const patched = applyPatch(content, searchReplacePatch(oldString, newString), {
      compareLine: (_n, line, _op, patchLine) => eq(line, patchLine),
    });
    if (patched === false) continue;
    return { kind: 'ok', content: patched };
  }
  return { kind: 'none' };
}

function diffArraysRatio(a: readonly string[], b: readonly string[]): number {
  if (a.length === 0 && b.length === 0) return 1;
  const parts = diffArrays([...a], [...b]) ?? [];
  let same = 0;
  for (const part of parts) {
    if (!part.added && !part.removed) same += part.value.length;
  }
  return (2 * same) / (a.length + b.length);
}

/** Aider window scan; similarity and excerpt come from jsdiff. */
function findSimilarExcerpt(search: string, content: string, threshold = 0.6): string {
  const searchLines = search.split('\n');
  const contentLines = content.split('\n');
  if (searchLines.length === 0 || searchLines.length > contentLines.length) return '';

  let bestRatio = 0;
  let bestMatch: string[] | null = null;
  let bestIndex = 0;
  const last = contentLines.length - searchLines.length;
  for (let i = 0; i <= last; i++) {
    const chunk = contentLines.slice(i, i + searchLines.length);
    const ratio = diffArraysRatio(searchLines, chunk);
    if (ratio > bestRatio) {
      bestRatio = ratio;
      bestMatch = chunk;
      bestIndex = i;
    }
  }
  if (bestRatio < threshold || !bestMatch) return '';

  const excerpt =
    bestMatch[0] === searchLines[0] &&
    bestMatch[bestMatch.length - 1] === searchLines[searchLines.length - 1]
      ? bestMatch.join('\n')
      : contentLines
          .slice(
            Math.max(0, bestIndex - 5),
            Math.min(contentLines.length, bestIndex + searchLines.length + 5)
          )
          .join('\n');
  return createTwoFilesPatch('old_string', 'document', search, excerpt);
}

function didYouMeanHint(content: string, oldString: string): string | null {
  const excerpt = findSimilarExcerpt(oldString, content);
  if (!excerpt) return null;
  return [
    'The old_string was not found in the document.',
    '',
    'Did you mean to match some of these actual lines from the document?',
    '',
    excerpt.trimEnd(),
    '',
    'Copy old_string exactly from the document side of the diff (or re-read that region with read_document).',
  ].join('\n');
}

export function findAndReplace(
  content: string,
  oldString: string,
  newString: string,
  replaceAll = false,
  expectedReplacements?: number
): FindAndReplaceResult {
  if (oldString === '') {
    return { success: false, error: 'old_string cannot be empty.' };
  }
  if (oldString === newString) {
    return {
      success: false,
      error: 'No changes to make: old_string and new_string are exactly the same.',
    };
  }

  // Layers 1 + 2: exact, else typography/entity-normalized.
  const actualOld = findActual(content, oldString);
  if (actualOld !== null) {
    const matchCount = content.split(actualOld).length - 1;

    // Declared-count guard: when the model says how many replacements it
    // expects, refuse to apply a different number. Catches replace_all
    // over/under-matching before it silently rewrites the wrong count.
    if (typeof expectedReplacements === 'number' && expectedReplacements !== matchCount) {
      return {
        success: false,
        error: `Expected ${expectedReplacements} occurrence(s) of old_string but found ${matchCount}. Re-read the document to confirm the current count, then adjust old_string (more/less context) or expected_replacements.`,
      };
    }

    if (!replaceAll && matchCount > 1) {
      return {
        success: false,
        error: `Found ${matchCount} matches for old_string. Set replace_all to true to replace all occurrences, or provide a more unique old_string.`,
      };
    }
    const styledNew = preserveQuoteStyle(oldString, actualOld, newString);
    return { success: true, newContent: applyReplace(content, actualOld, styledNew, replaceAll) };
  }

  // Layer 3: jsdiff applyPatch with Codex compareLine passes.
  if (!replaceAll) {
    const located = applyWithCompare(content, oldString, newString);
    if (located.kind === 'ambiguous') {
      return {
        success: false,
        error: `Found ${located.count} matches for old_string. Set replace_all to true to replace all occurrences, or provide a more unique old_string.`,
      };
    }
    if (located.kind === 'ok') {
      if (typeof expectedReplacements === 'number' && expectedReplacements !== 1) {
        return {
          success: false,
          error: `Expected ${expectedReplacements} occurrence(s) of old_string but found 1. Re-read the document to confirm the current count, then adjust old_string or expected_replacements.`,
        };
      }
      return { success: true, newContent: located.content };
    }
  }

  // Miss: whitespace-collapsed hint, else Aider "Did you mean?" excerpt.
  const trimLine = (s: string) =>
    s
      .split('\n')
      .map((l) => l.trimEnd())
      .join('\n');
  const normalize = (s: string) =>
    trimLine(s)
      .replace(/\n{3,}/g, '\n\n')
      .trim();

  if (normalize(content).indexOf(normalize(oldString)) !== -1) {
    return {
      success: false,
      error:
        'A whitespace-normalized match exists but not an exact one. Copy old_string exactly from the document, preserving all whitespace and blank lines.',
    };
  }

  return {
    success: false,
    error:
      didYouMeanHint(content, oldString) ??
      'The old_string was not found in the document. It must match the document EXACTLY, including every leading space and blank line. Copy it verbatim from the most recent read_document output. If read_document reported truncated=true, the match may be outside the returned window — re-read with offset/limit around the target region first.',
  };
}
