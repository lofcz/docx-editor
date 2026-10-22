import type { RevisionToken, NativeEditFailure } from '../contracts/document';
import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { ProjectedBlock } from '../projection/blocks';
import type { ProjectedChecklist } from '../projection/lists';
import type { MdxProjection, OpaqueMdxSpan } from '../projection/mdx';
import type { MarkdownSourceSlice } from '@docx-editor.dev/docx-to-markdown';

export interface TargetSourceInput {
  readonly revision: RevisionToken;
  readonly pkg: OoxmlPackage;
  /** Produce this from the same immutable package and native review view. */
  readonly projection: MdxProjection;
}
export interface TargetSourceExtent {
  readonly start: number;
  readonly end: number;
  /** Paragraph extents keep canonical model coordinates. Other constructs use null. */
  readonly native: { readonly start: number; readonly end: number } | null;
}
export interface TargetSourceOwner {
  readonly id: string;
  readonly kind: 'paragraph' | 'table' | 'block';
  readonly parentId: string;
  readonly childIndex: number;
  /** Self for a paragraph, enclosing paragraph for an inline control, otherwise null. */
  readonly paragraphId: string | null;
  /** Nearest enclosing cell; a nested table retains its outer cell here. */
  readonly cell: {
    readonly tableId: string;
    readonly rowId: string;
    readonly cellId: string;
  } | null;
  readonly visibility: 'projected' | 'generated' | 'atomic-fallback' | 'unprojected';
  readonly extents: readonly TargetSourceExtent[];
  readonly sources: readonly MarkdownSourceSlice[];
  readonly block: ProjectedBlock | null;
  readonly opaque: OpaqueMdxSpan | null;
  readonly checklist: ProjectedChecklist | null;
}
export interface TargetSourceWindow {
  readonly parentId: string;
  readonly startIndex: number;
  readonly ids: readonly string[];
  readonly beforeId: string | null;
  readonly afterId: string | null;
}
export interface TargetSourceIndex {
  readonly revision: RevisionToken;
  readonly partName: string;
  readonly bodyId: string;
  readonly owners: readonly TargetSourceOwner[];
  /** Actual contiguous native children. No flattened projection becomes a sibling window. */
  readonly windows: readonly TargetSourceWindow[];
}
export type TargetSourceResult =
  | { readonly success: true; readonly index: TargetSourceIndex }
  | NativeEditFailure;
