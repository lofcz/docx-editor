import type { OoxmlProperty } from '../store/store/tree-op-types.ts';
import type { ResolvedRunStyle } from './run-style.ts';

/** Owning run formatting, including the resolved paragraph and character style cascade. @public */
export interface InlineDrawingFormatting {
  readonly style: ResolvedRunStyle;
  readonly props: readonly OoxmlProperty[];
}

export function inlineDrawingFormatting(formatting: InlineDrawingFormatting | undefined): {
  readonly formatting?: InlineDrawingFormatting;
} {
  return formatting
    ? { formatting: Object.freeze({ style: formatting.style, props: formatting.props }) }
    : {};
}
