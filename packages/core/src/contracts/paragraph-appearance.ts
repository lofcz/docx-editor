import type { ContentControlMappedType } from '../layout/semantic-control-records.ts';
import type { RevisionDisplayMode } from '../layout/revision-projection.ts';

/**
 * Read-only inline-control facts for a paragraph appearance resolver.
 * Core applies revision visibility before creating these records.
 * Records contain no tree nodes, package maps, or mutation methods.
 * @internal Proposed host contribution; not wired or exported yet.
 */
export interface ParagraphAppearanceControl {
  readonly id: string;
  readonly tag?: string;
  readonly controlType: ContentControlMappedType;
  readonly start: number;
  readonly end: number;
  readonly checked?: boolean;
}

/** @internal Proposed host contribution; not wired or exported yet. */
export interface ParagraphAppearanceContext {
  readonly paragraphId: string;
  readonly displayMode: RevisionDisplayMode;
  /**
   * Effective controls in this view. All Markup uses the proposed control state
   * for paragraph appearance while retaining both revision halves in native content.
   */
  readonly controls: readonly ParagraphAppearanceControl[];
}

/**
 * Metric-neutral defaults. Authored character styles and direct run properties
 * retain their native precedence. Core validates colors through its Word resolver.
 * @internal Proposed host contribution; not wired or exported yet.
 */
export interface ParagraphAppearanceDefaults {
  readonly strike?: boolean;
  readonly color?: {
    readonly val: string;
    readonly themeColor?: string;
    readonly themeTint?: string;
  };
}

/**
 * A deterministic, DOM-free host contribution. The resolver reads only its input.
 * Core freezes bounded input records and validates the returned defaults.
 * This contribution creates no document operations or persisted appearance state.
 *
 * Proposed attachment points: EditorModule.paragraphAppearance and
 * OpenDocumentForExportOptions.paragraphAppearance. Both reach the shared
 * document-layout coordinator. Word materialization calls the same resolver
 * against a detached snapshot and leaves pending review paragraphs unchanged.
 * @internal Proposed host contribution; not wired or exported yet.
 */
export interface ParagraphAppearanceContribution {
  /** Stable implementation identity, included in layout cache keys. */
  readonly id: string;
  readonly resolve: (context: ParagraphAppearanceContext) => ParagraphAppearanceDefaults | null;
}
