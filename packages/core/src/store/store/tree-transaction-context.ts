import type { OoxmlPackage } from '../package/ooxml-package.ts';
import type { TreeDocOp } from './tree-op-types.ts';
import type { SelectionMark } from './tree-store.ts';

/** What a transaction body is handed: the working tree, and the means to stage ops against it. */
export interface TransactionContext {
  /** Current intermediate package. A read; this does not publish or replace canonical state. */
  currentPackage(): OoxmlPackage;
  /** Stage one op against the STORY part. Returns false once the transaction has failed. */
  apply(op: TreeDocOp): boolean;
  /**
   * Stage one op against a named part.
   *
   * A comment body lives in `comments.xml` and is edited by the same ops that edit the story,
   * so this is `apply` with the target named rather than a second vocabulary.
   */
  applyTo(partName: string, op: TreeDocOp): boolean;
  /**
   * Stage a whole-package edit: a new part, a relationship, a content-type override.
   *
   * The edit is a pure function of the working package, so a rejected transaction discards it
   * with everything else. Returning the SAME package is a no-op, not a failure — a primitive
   * that finds nothing to do says so that way.
   *
   * Engine package edits refuse by throwing the internal `PackageEditRefusal`
   * (`package-edit-refusal.ts`): the transaction is refused with its reason, `package-changed`
   * by default, and nothing is applied. Any other throw escapes `transact` as a crash.
   */
  applyPackage(edit: (pkg: OoxmlPackage) => OoxmlPackage): boolean;
  /** The selection to restore when this entry is undone. */
  selectionBefore(selection: SelectionMark | null): void;
  /** The selection to restore when this entry is redone. */
  selectionAfter(selection: SelectionMark | null): void;
}
