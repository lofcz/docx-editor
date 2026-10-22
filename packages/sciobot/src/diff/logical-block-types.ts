import type { JsonObject } from '../data/json';
import type { NativeEditFailure } from '../contracts/document';
import type { TargetAstNode, TargetAstRoot } from '../projection/parse-target';

export interface LogicalListContext {
  readonly marker: 'ordered' | 'bullet';
  readonly checked: boolean | null;
  readonly level: number;
  /** Paths address this input AST. They are not persistent or native list instance IDs. */
  readonly group: readonly number[];
  readonly item: readonly number[];
  readonly role: 'item' | 'continuation';
  /** Syntax data only. Generated source markers must not become native restart writes. */
  readonly start: number | null;
  /** The application recap rule removes numbering/indentation. Task state stays explicit. */
  readonly suppressed: boolean;
}
interface LogicalBase {
  readonly path: readonly number[];
  readonly list: LogicalListContext | null;
  /** Content matching hint only. List, property, and native ownership checks remain necessary. */
  readonly key: string;
}
export interface LogicalParagraph extends LogicalBase {
  readonly kind: 'paragraph';
  readonly ast: TargetAstNode;
  readonly synthetic: boolean;
  /** Child origins for synthetic paragraphs; each path addresses the untouched root AST. */
  readonly inlinePaths: readonly (readonly number[])[];
  readonly prefixRemoved: number;
  readonly html: JsonObject | null;
}
export interface LogicalRowGroup {
  readonly path: readonly number[];
  readonly header: JsonObject;
  readonly start: number;
  readonly end: number;
}
export interface LogicalGroup extends LogicalBase {
  readonly kind: 'container' | 'table' | 'row' | 'cell';
  /** Own AST fields. Structural children live below; the complete source remains in document.ast. */
  readonly header: JsonObject;
  readonly domainType: string | null;
  readonly children: readonly LogicalBlock[];
  readonly rowGroups: readonly LogicalRowGroup[];
}
export interface LogicalAtom extends LogicalBase {
  readonly kind: 'atomic' | 'block';
  /** Atomic nodes retain their entire subtree without interpreting descendants. */
  readonly ast: TargetAstNode;
}
export type LogicalBlock = LogicalParagraph | LogicalGroup | LogicalAtom;
export interface LogicalBlockDocument {
  readonly markdown: string;
  /** Full detached original AST, including unknown list/group metadata and diagnostic positions. */
  readonly ast: TargetAstRoot;
  readonly blocks: readonly LogicalBlock[];
}
export type LogicalBlockResult =
  | { readonly success: true; readonly document: LogicalBlockDocument }
  | NativeEditFailure;
