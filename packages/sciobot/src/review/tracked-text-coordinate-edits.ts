import { refuseWrite } from '../diff/fragment-seeds';
import type { CanonicalEdit } from '../diff/plan';
import type { StagedTextCoordinates } from './tracked-stage-coordinates';

type TextEdit = Extract<CanonicalEdit, { kind: 'text' }>;

/** Split visible deletion ranges around native tombstones; each source span stays independently guarded. */
export function trackedTextCoordinateEdits(
  edits: readonly TextEdit[],
  coordinates: StagedTextCoordinates
): readonly TextEdit[] {
  const original = coordinates.text();
  return edits.flatMap((edit) => {
    if (original.slice(edit.start, edit.end) !== edit.before)
      refuseWrite('Planned before-text does not match the proposed paragraph.', 'stale-revision');
    if (edit.start === edit.end) {
      const offset = coordinates.point(edit.start, edit.bias);
      return [{ ...edit, start: offset, end: offset }];
    }
    let consumed = 0;
    return coordinates.slices(edit.start, edit.end).map((range, index) => {
      const length = range.end - range.start;
      const before = edit.before.slice(consumed, consumed + length);
      consumed += length;
      return { ...edit, ...range, before, after: index === 0 ? edit.after : '' };
    });
  });
}
