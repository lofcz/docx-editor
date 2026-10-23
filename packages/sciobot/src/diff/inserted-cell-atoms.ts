import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import type { TargetAstNode } from '../projection/parse-target';
import type { InsertedCellPayloadBinding } from './inserted-cell-fragments';
import type { InsertedInlineAtom } from './inserted-atoms';
import { isInsertedAtomBindings, planInsertedAtomBindings } from './inserted-atom-bindings';

/** Prefix every paragraph address with its physical cell stream. No preview IDs enter the plan. */
const roots = (cells: readonly (readonly TargetAstNode[])[]): readonly TargetAstNode[] =>
  cells.map((children) => ({ type: 'root', children: [...children] }));

export function planInsertedCellAtoms(
  pkg: OoxmlPackage,
  cells: readonly (readonly TargetAstNode[])[],
  payloads: readonly InsertedCellPayloadBinding[],
  allocate?: () => string
): { readonly success: true; readonly atoms: readonly InsertedInlineAtom[] } | NativeEditFailure {
  return planInsertedAtomBindings(pkg, roots(cells), payloads, allocate);
}

/** Pure reconstruction checks paths, offsets, source descriptors, payloads and identities. */
export function isInsertedCellAtoms(
  cells: readonly (readonly TargetAstNode[])[],
  payloads: readonly InsertedCellPayloadBinding[],
  atoms: unknown
): boolean {
  return isInsertedAtomBindings(roots(cells), payloads, atoms === undefined ? [] : atoms);
}

/** Strip the physical cell prefix only after execution allocates that cell's native identity. */
export function cellAtomBindings(
  atoms: readonly InsertedInlineAtom[] | undefined,
  cell: number
): readonly InsertedInlineAtom[] {
  return (atoms === undefined ? [] : atoms).flatMap((atom) =>
    atom.location.kind === 'new-paragraph' && atom.location.path[0] === cell
      ? [
          {
            ...atom,
            location: { kind: 'new-paragraph' as const, path: atom.location.path.slice(1) },
          },
        ]
      : []
  );
}

/** One atom inventory for D01 validation, executor resource preparation and shared admission. */
export function insertedAtomsOfEdit(
  edit: import('./plan').CanonicalEdit
): readonly InsertedInlineAtom[] {
  if (edit.kind === 'inline-atoms' || edit.kind === 'siblings') return edit.atoms ?? [];
  if (edit.kind === 'insert-row' || edit.kind === 'insert-column' || edit.kind === 'split-cell')
    return edit.intent.atoms ?? [];
  if (edit.kind === 'table-geometry' && edit.intent.kind === 'split')
    return edit.intent.atoms ?? [];
  return [];
}
