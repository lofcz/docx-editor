import type { OoxmlProperty } from '../store/store/tree-op-types.ts';
import type { ResolvedRunStyle } from './run-style.ts';

/** Owning run formatting, including the resolved paragraph and character style cascade. @public */
export interface InlineDrawingFormatting {
  readonly style: ResolvedRunStyle;
  readonly props: readonly OoxmlProperty[];
  /** Hyperlink envelope outside the owning run or content control. */
  readonly link?: import('./semantic-records.ts').SpanLinkRecord;
}

export function inlineDrawingFormatting(formatting: InlineDrawingFormatting | undefined): {
  readonly formatting?: InlineDrawingFormatting;
} {
  return formatting
    ? {
        formatting: Object.freeze({
          style: formatting.style,
          props: formatting.props,
          ...(formatting.link ? { link: formatting.link } : {}),
        }),
      }
    : {};
}
