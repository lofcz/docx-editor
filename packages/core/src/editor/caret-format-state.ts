import { clampedToDocument } from './surface-selection-ops.ts';
import type { SemanticLayout } from '../layout/index.ts';
import type { SemanticPosition, SemanticSelection } from '../layout/semantic-interaction.ts';
import type { SurfaceProperty } from './surface-formatting.ts';
import { mergedRunWrite } from '../store/store/direct-properties.ts';
import type { TreeDocOp } from '../store/index.ts';

/** Capture the original face once, so later edits cannot change the armed typing style. */
export interface ArmedFormat {
  readonly properties: readonly SurfaceProperty[];
  readonly base: readonly SurfaceProperty[];
}
export type PendingFormat = ({ readonly position: SemanticPosition } & ArmedFormat) | null;

/** Character styles use the explicit operation field, outside replaceable run properties. */
export function pendingCharacterStyle(properties: readonly SurfaceProperty[]): {
  readonly characterStyle?: string | null;
} {
  const style = properties.find((property) => property.localName === 'rStyle');
  return style ? { characterStyle: style.attributes?.val || null } : {};
}

/** Style identity changes need local history even when their visible formatting agrees. */
export function pendingCharacterStyleChanged(before: PendingFormat, after: PendingFormat): boolean {
  return (
    JSON.stringify(pendingCharacterStyle(before?.properties ?? [])) !==
    JSON.stringify(pendingCharacterStyle(after?.properties ?? []))
  );
}

/** One insertion writes its captured face and explicit character style in the same transaction. */
export function pendingFormatOperation(
  armed: ArmedFormat,
  paragraphId: string,
  offset: number,
  length: number,
  complex: boolean
): TreeDocOp {
  return {
    op: 'setRunProperties',
    paragraphId,
    start: offset,
    end: offset + length,
    properties: mergedRunWrite(
      armed.base,
      armed.properties.filter((p) => p.localName !== 'rStyle'),
      complex
    ),
    ...pendingCharacterStyle(armed.properties),
  };
}

/** Only the caret that armed a format can consume it. */
export function armedFormatAt(
  pending: PendingFormat,
  selection: SemanticSelection
): ArmedFormat | null {
  if (!pending) return null;
  const at = pending.position;
  const same = (position: SemanticPosition) =>
    position.paragraphId === at.paragraphId && position.offset === at.offset;
  return same(selection.anchor) && same(selection.head) ? pending : null;
}

/** Undefined means a range selection cannot arm a caret format. */
export function armCaretFormat(
  selection: SemanticSelection,
  previous: PendingFormat,
  next: readonly SurfaceProperty[] | null,
  readBase: () => readonly SurfaceProperty[]
): PendingFormat | undefined {
  if (next === null || next.length === 0) return previous ? null : undefined;
  const { anchor, head } = selection;
  if (anchor.paragraphId !== head.paragraphId || anchor.offset !== head.offset) return undefined;
  return {
    position: head,
    properties: next,
    base: armedFormatAt(previous, selection)?.base ?? readBase(),
  };
}

export interface CaretFormatSnapshot {
  readonly selection: SemanticSelection;
  readonly format: PendingFormat;
}
export function restoreCaretFormatSnapshot(
  value: CaretFormatSnapshot,
  layout: SemanticLayout,
  order: readonly string[]
): CaretFormatSnapshot {
  const selection = clampedToDocument(layout, order, value.selection);
  const valid = order.includes(value.selection.head.paragraphId);
  return {
    selection,
    format: valid && value.format ? { ...value.format, position: selection.head } : null,
  };
}
