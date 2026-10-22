import {
  bodyStoryRoot,
  paragraphTextOf,
  storyParagraphs,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';

export interface NativeTextPoint {
  readonly paragraphId: string;
  readonly offset: number;
}

export interface NativeTextRange {
  readonly start: NativeTextPoint;
  readonly end: NativeTextPoint;
}

export type ProjectionSegment =
  | {
      readonly kind: 'text';
      readonly start: number;
      readonly end: number;
      readonly paragraphId: string;
      readonly rawStart: number;
      readonly rawEnd: number;
    }
  | {
      readonly kind: 'syntax';
      readonly start: number;
      readonly end: number;
      readonly role: string;
    };

export interface ProjectionParagraph {
  readonly id: string;
  readonly start: number;
  readonly end: number;
  readonly rawText: string;
}

/** A serializable read model. It contains no live editor or canonical node objects. */
export interface TextProjection {
  readonly revision: RevisionToken;
  readonly text: string;
  readonly paragraphs: readonly ProjectionParagraph[];
  readonly segments: readonly ProjectionSegment[];
}

export type TextProjectionResult =
  | { readonly success: true; readonly projection: TextProjection }
  | NativeEditFailure;

/**
 * Project raw canonical body text. Review views supply their derived part separately.
 * Paragraph separators are generated syntax; native breaks remain inside their paragraph.
 */
export function projectBodyText(part: OoxmlPart, revision: RevisionToken): TextProjectionResult {
  const root = bodyStoryRoot(part);
  if (!root)
    return { success: false, code: 'invalid-document', error: 'The document has no body.' };
  const paragraphs: ProjectionParagraph[] = [];
  const segments: ProjectionSegment[] = [];
  const chunks: string[] = [];
  let offset = 0;
  for (const node of storyParagraphs(root)) {
    if (paragraphs.length) {
      chunks.push('\n\n');
      segments.push(
        Object.freeze({
          kind: 'syntax',
          start: offset,
          end: offset + 2,
          role: 'paragraph-separator',
        })
      );
      offset += 2;
    }
    const rawText = paragraphTextOf(part, node.id);
    if (rawText === null)
      return {
        success: false,
        code: 'invalid-document',
        error: `Paragraph ${node.id} has no model text.`,
      };
    const end = offset + rawText.length;
    chunks.push(rawText);
    paragraphs.push(Object.freeze({ id: node.id, start: offset, end, rawText }));
    segments.push(
      Object.freeze({
        kind: 'text',
        start: offset,
        end,
        paragraphId: node.id,
        rawStart: 0,
        rawEnd: rawText.length,
      })
    );
    offset = end;
  }
  return {
    success: true,
    projection: Object.freeze({
      revision: Object.freeze({ ...revision }),
      text: chunks.join(''),
      paragraphs: Object.freeze(paragraphs),
      segments: Object.freeze(segments),
    }),
  };
}

function splitsSurrogate(text: string, offset: number): boolean {
  if (offset === 0 || offset === text.length) return false;
  const before = text.charCodeAt(offset - 1);
  const after = text.charCodeAt(offset);
  return before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff;
}

/** Generated syntax has no native caret position. Offsets use UTF-16 on both sides. */
export function nativePointAt(
  projection: TextProjection,
  offset: number,
  affinity: 'forward' | 'backward' = 'forward'
): NativeTextPoint | null {
  if (
    !Number.isInteger(offset) ||
    offset < 0 ||
    offset > projection.text.length ||
    splitsSurrogate(projection.text, offset)
  )
    return null;
  const candidates = projection.segments.filter(
    (segment) => segment.kind === 'text' && segment.start <= offset && offset <= segment.end
  );
  const segment = affinity === 'forward' ? candidates.at(-1) : candidates[0];
  if (
    !segment ||
    segment.kind !== 'text' ||
    segment.end - segment.start !== segment.rawEnd - segment.rawStart
  )
    return null;
  return { paragraphId: segment.paragraphId, offset: segment.rawStart + offset - segment.start };
}

export function nativeRangeAt(
  projection: TextProjection,
  start: number,
  end: number
): NativeTextRange | null {
  if (end < start) return null;
  const from = nativePointAt(projection, start, 'forward');
  const to = nativePointAt(projection, end, start === end ? 'forward' : 'backward');
  return from && to ? { start: from, end: to } : null;
}
