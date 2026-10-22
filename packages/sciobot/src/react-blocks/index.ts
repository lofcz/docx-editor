/** Optional React integration. Import document tools from @sciobot/docx without loading React. */
export * from './host';
export * from './block-ui';
export type {
  ReactBlockProviders,
  ReactBlockProviderScope,
  ReactBlockHostOptions,
} from './providers';
export * from './selection';
export * from './input-boundary';
export * from './clipboard';
export { BlockMeasurement, lastBodyParagraphIn } from './measurement';
export type { MeasureBlockElement, MeasuredFragment } from './measurement';
