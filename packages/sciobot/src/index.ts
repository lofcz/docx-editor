/** Private application integration. The native editor owns document state. */
export type { Editor } from '@docx-editor.dev/core/contracts/editor';
export {
  findAndReplace as findAndReplaceMarkdown,
  type FindAndReplaceResult,
} from './tools/find-replace';
export type * from './contracts/executor';
export { sameRevision } from './contracts/document';
export type * from './contracts/document';
export * from './diff/domain-strategies';
export * from './diff/domain-containers';
export * from './diff/domain-properties';
export * from './diff/atomic-blocks';
export * from './diff/plan';
export * from './diff/inline-target';
export { targetAstKey } from './diff/target-key';
export * from './diff/inserted-atoms';
export * from './diff/target-source';
export * from './diff/logical-blocks';
export * from './diff/bind-source';
export * from './diff/plan-target';
export * from './diff/paragraph-text';
export * from './diff/run-format';
export * from './diff/paragraph-format';
export * from './diff/sibling-blocks';
export * from './diff/lists';
export * from './diff/table-cells';
export * from './diff/table-rows';
export * from './diff/table-columns';
export * from './diff/table-merges';
export * from './diff/table-format';
export * from './diff/inserted-table-columns';
export * from './diff/inserted-table-rows';
export * from './projection/text';
export * from './projection/markdown';
export * from './projection/lists';
export * from './projection/quote-selection';
export * from './projection/tables';
export * from './projection/blocks';
export * from './projection/mdx';
export * from './projection/resolve-patch';
export * from './projection/protected-content';
export * from './projection/parse-target';
export { DOMAIN_BLOCK_TAGS, GENERATED_LABEL_TAG } from './projection/block-tags';
export type {
  BlockTagDefinition,
  BlockAttributeBinding,
  BlockPropertyPath,
  SerializedBlockProperty,
  SerializedBlockTag,
} from './projection/block-tags';
export * from './review/committed-view';
export * from './review/working-view';
export * from './review/property-projection';
export * from './review/property-record-schema';
export * from './review/property-record-store';
export * from './review/stored-property-proposals';
export * from './review/track-properties';
export * from './review/decide-properties';
export * from './review/track-atomic';
export * from './review/groups';
export * from './review/accept-group';
export * from './review/accept-all';
export * from './review/reject-all';
export * from './review/restore';
export * from './review/reconcile-external';
export * from './review/owned-proposal-plan';
export * from './review/replace-owned-proposals';
export * from './review/recompose';
export * from './review/owned-proposal-content';
export * from './review/reject-group';
export * from './review/group-record-schema';
export * from './review/group-record-store';
export * from './review/native-module';
export * from './review/track-text';
export * from './review/track-hyperlinks';
export * from './review/track-checklists';
export * from './review/track-format';
export * from './review/track-blocks';
export * from './review/track-table-cells';
export * from './ai/apply-direct';
export * from './ai/stream-insert';
export * from './ai/markdown-stream';
export * from './ai/change-ledger';
export * from './ai/direct-baseline';
export * from './ai/playback';
export * from './ai/acknowledge';
export type {
  ReviewDocumentView,
  ReviewViewOptions,
  ReviewViewResult,
} from './review/document-view';
export * from './projection/revision-cache';
export * from './tools/read';
export * from './tools/search';
export * from './tools/validate-target';
export * from './tools/edit-echo';
export * from './blocks/payload-schema';
export * from './blocks/payload-opc-codec';
export * from './blocks/payload-part';
export * from './blocks/block-index';
export * from './blocks/insert';
export * from './blocks/math-codec';
export * from './blocks/mention-codec';
export * from './blocks/mention-insert';
export * from './blocks/math-insert';
export * from './blocks/remove';
export * from './blocks/move';
export * from './native/inline-code';
export * from './native/prepare-document';
export * from './native/prepare-review-groups';
export * from './chrome/block-type';
export * from './chrome/block-styles';
export { insertBlock as insertTextBlock, type InsertBlockType } from './chrome/insert-block';
export * from './chrome/block-input';
export * from './chrome/block-toolbar';

export * from './review/group-allocation';

export type { TrackedPayloadOptions } from './review/tracked-stage-payloads';

export {
  prepareInlineImageResources,
  type PreparedInlineImages,
  type InlineImageSource,
  type InlineImageResolver,
} from './native/inline-image-resources';
