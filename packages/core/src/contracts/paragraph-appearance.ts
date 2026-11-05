import type { ContentControlKind } from '@docx-editor.dev/core/store';
import type { RevisionDisplayMode } from '../layout/revision-projection.ts';

/**
 * Read-only inline-control facts for a paragraph appearance resolver.
 * Core applies revision visibility before creating these records.
 * Records contain no tree nodes, package maps, or mutation methods.
 * @public
 */
export interface ParagraphAppearanceControl {
  readonly id: string;
  readonly tag?: string;
  readonly controlType: ContentControlKind;
  readonly start: number;
  readonly end: number;
  readonly checked?: boolean;
}

/** Immutable facts supplied once per paragraph and revision view. @public */
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
 * @public
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
 * Register through EditorModule.paragraphAppearance or pass the contribution to
 * OpenDocumentForExportOptions.paragraphAppearance. Both use the same native
 * style cascade. DOCX saving retains authored styles; materialization belongs
 * to the host's detached export preparation.
 *
 * A registry permits one contribution. Repeating the same ID and callback is
 * allowed; conflicting registrations throw. Changed callbacks invalidate layout
 * caches, including when their IDs match. Resolver closures must remain immutable.
 * @public
 */
export interface ParagraphAppearanceContribution {
  /** Diagnostic identity: a nonempty string containing at most 256 characters. */
  readonly id: string;
  readonly resolve: (context: ParagraphAppearanceContext) => ParagraphAppearanceDefaults | null;
}
