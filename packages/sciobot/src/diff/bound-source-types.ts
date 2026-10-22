import type { NativeEditFailure } from '../contracts/document';
import type { ParsedTargetMarkdown, TargetSourceRange } from '../projection/parse-target';
import type { LogicalBlockDocument } from './logical-block-types';
import type { TargetSourceIndex, TargetSourceWindow } from './target-source-types';

export interface BoundSourceContribution {
  readonly ownerId: string;
  /** Index in that owner's original extents. No clipped native model range is inferred. */
  readonly extentIndex: number;
  readonly start: number;
  readonly end: number;
}
export interface BoundSourceInline {
  readonly controlId: string;
  readonly blockId: string;
  readonly start: number;
  readonly end: number;
  readonly native: { readonly start: number; readonly end: number };
}
export interface BoundSourceTable {
  readonly tableId: string;
  readonly rowId: string | null;
  readonly cellId: string | null;
  readonly gridColumn: number | null;
  readonly gridSpan: number | null;
  readonly role:
    | 'table'
    | 'row'
    | 'cell'
    | 'horizontal-placeholder'
    | 'vertical-placeholder'
    | 'missing-slot';
}
export interface BoundSourceBlock {
  /** Index path through logical sibling streams, distinct from the original AST path. */
  readonly logicalPath: readonly number[];
  readonly astPath: readonly number[];
  readonly source: TargetSourceRange | null;
  readonly contributions: readonly BoundSourceContribution[];
  /** Present only for one logical block and one actual native sibling, without competing claims. */
  readonly structuralOwnerId: string | null;
  readonly table: BoundSourceTable | null;
  readonly inline: readonly BoundSourceInline[];
  /** Native atomic policy wins even when a stored opaque tag resembles a known container. */
  readonly atomicControlId: string | null;
  readonly generatedTask: {
    readonly paragraphId: string;
    readonly controlId: string;
    readonly start: number;
    readonly end: number;
    readonly nativeInsertionOffset: number;
  } | null;
}
export interface BoundSourceWindow extends TargetSourceWindow {
  readonly members: readonly {
    readonly ownerId: string;
    readonly logicalPaths: readonly (readonly number[])[];
  }[];
}
export interface BoundLogicalSource {
  readonly index: TargetSourceIndex;
  readonly parsed: Pick<ParsedTargetMarkdown, 'sourceMarkdown' | 'markdown' | 'normalization'>;
  readonly logical: LogicalBlockDocument;
  readonly bindings: readonly BoundSourceBlock[];
  readonly windows: readonly BoundSourceWindow[];
  /** No missing AST block becomes a fabricated native node or text range. */
  readonly unrepresented: readonly {
    readonly ownerId: string;
    readonly reason:
      | 'generated'
      | 'atomic-fallback'
      | 'unprojected'
      | 'empty-extent'
      | 'not-parsed';
  }[];
}
export type BoundSourceResult =
  | { readonly success: true; readonly source: BoundLogicalSource }
  | NativeEditFailure;
