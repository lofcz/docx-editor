import { segmentGraphemes } from '@docx-editor.dev/core/layout';
import {
  AUTHORABLE_RUN_PROPERTIES,
  authoredProperties,
  findNode,
  paragraphTextOf,
  propertyContainer,
  runPropertyEdits,
  runsCovering,
  validateTreeOp,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import type { CanonicalEdit, PlannedValue } from '../diff/plan';
import { runCharacterStyle } from '../native/inline-code';

/** Plan change name for the run's own `w:rStyle`, outside the accepted toggle vocabulary. */
const CHARACTER_STYLE = 'rStyle';

type FormatEdit = Extract<CanonicalEdit, { kind: 'format' }>;
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
function attributes(value: PlannedValue): Record<string, string> | null {
  if (
    !value.present ||
    value.value === null ||
    typeof value.value !== 'object' ||
    Array.isArray(value.value)
  )
    return null;
  return Object.values(value.value).every((entry) => typeof entry === 'string')
    ? (value.value as Record<string, string>)
    : null;
}
function sameAttributes(a: Readonly<Record<string, string>> | undefined, b: PlannedValue): boolean {
  if (!b.present) return a === undefined;
  const old = attributes(b);
  if (!a || !old || Object.keys(a).length !== Object.keys(old).length) return false;
  return Object.keys(a).every((key) => Object.hasOwn(old, key) && a[key] === old[key]);
}

/** Recheck lexical property presence and lower deltas over each run's complete native bag. */
export function trackedRunFormatOps(
  part: OoxmlPart,
  edits: readonly FormatEdit[],
  revision: { author: string; date: string }
): { success: true; ops: readonly TreeDocOp[] } | NativeEditFailure {
  const ops: TreeDocOp[] = [],
    ranges = new Map<string, { start: number; end: number }[]>();
  for (const edit of edits) {
    if (edit.partName !== part.name || edit.target.kind !== 'run')
      return fail('Only native body run formatting is supported.', 'unsupported-operation');
    const target = edit.target,
      paragraph = findNode(part, target.paragraphId),
      text = paragraphTextOf(part, target.paragraphId);
    if (
      paragraph?.kind !== 'paragraph' ||
      text === null ||
      target.start >= target.end ||
      target.end > text.length
    )
      return fail('Invalid native format range.');
    const boundaries = new Set([
      0,
      text.length,
      ...segmentGraphemes(text).flatMap((item) => [item.utf16From, item.utf16To]),
    ]);
    if (!boundaries.has(target.start) || !boundaries.has(target.end))
      return fail('A format range splits a native grapheme.', 'ambiguous-target');
    const previous = ranges.get(target.paragraphId) ?? [];
    if (
      previous.some(
        (range) => Math.max(range.start, target.start) < Math.min(range.end, target.end)
      )
    )
      return fail('Planned format ranges overlap.', 'ambiguous-target');
    previous.push(target);
    ranges.set(target.paragraphId, previous);
    for (const change of edit.changes) {
      if (change.name === CHARACTER_STYLE) {
        // The run's own character style (N18 inline code): one `val` attribute or absence.
        const style = [change.before, change.after].every(
          (value) =>
            !value.present ||
            (attributes(value) !== null &&
              Object.keys(attributes(value)!).join() === 'val' &&
              attributes(value)!.val!.length > 0)
        );
        if (!style) return fail('Invalid native character style delta.');
        continue;
      }
      if (
        !AUTHORABLE_RUN_PROPERTIES.has(change.name) ||
        (change.before.present && attributes(change.before) === null) ||
        (change.after.present && attributes(change.after) === null)
      )
        return fail('Invalid native run property delta.');
    }
    const native = runPropertyEdits(
      part,
      target.paragraphId,
      target.start,
      target.end,
      [],
      'all-markup'
    );
    const runs = runsCovering(part, target.paragraphId, target.start, target.end, 'all-markup');
    if (!native.length || native.length !== runs.length || native.some((item) => item.targetRunIds))
      return fail('Formatting requires unambiguous native run ownership.', 'unsupported-operation');
    for (let index = 0; index < native.length; index++) {
      const original = authoredProperties(
        propertyContainer(runs[index], 'runProperties', 'rPr'),
        AUTHORABLE_RUN_PROPERTIES
      );
      const bag = new Map(original.map((property) => [property.localName, property]));
      let characterStyle: string | null | undefined;
      for (const change of edit.changes) {
        if (change.name === CHARACTER_STYLE) {
          const current = runCharacterStyle(runs[index]!);
          if (!sameAttributes(current === undefined ? undefined : { val: current }, change.before))
            return fail(
              'Planned run properties no longer match their before-values.',
              'stale-revision'
            );
          characterStyle = change.after.present ? attributes(change.after)!.val! : null;
          continue;
        }
        const property = bag.get(change.name);
        if (!sameAttributes(property ? (property.attributes ?? {}) : undefined, change.before))
          return fail(
            'Planned run properties no longer match their before-values.',
            'stale-revision'
          );
        if (!change.after.present) bag.delete(change.name);
        else
          bag.set(change.name, { localName: change.name, attributes: attributes(change.after)! });
      }
      const range = native[index]!;
      const op: TreeDocOp = {
        op: 'setRunProperties',
        paragraphId: target.paragraphId,
        start: range.start,
        end: range.end,
        properties: [...bag.values()],
        ...(characterStyle !== undefined ? { characterStyle } : {}),
        revision,
      };
      const rejection = validateTreeOp(part, op);
      if (rejection) return fail(String(rejection), 'apply-failed');
      ops.push(op);
    }
  }
  return { success: true, ops };
}
