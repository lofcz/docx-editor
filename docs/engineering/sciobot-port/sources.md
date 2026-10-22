# Source map

Source references describe observed application behavior and native engine constraints. Primary task files describe proposed write locations.

Resolve `app:`, `plate:`, and `fork:` paths from these checkout roots:

- `fork:` [docx-editor](/GitHub/docx-editor)
- `app:` [sciobot-next](/run/media/lofcz/ssd_external/GitHub/sciobot-next)
- `plate:` [plate](/run/media/lofcz/ssd_external/GitHub/plate)

Git revisions identify reference points. Uncommitted changes also form part of the source baseline.

| Checkout    | Reference commit                           |
| ----------- | ------------------------------------------ |
| Application | `da9081adad1b54d17d9f24647adfd117bdffb942` |
| Plate fork  | `d81a8eee2cca9f38a0cb04d88a0019405bf109c6` |
| DOCX fork   | `9770e78f880ee92ecc6531f0ea747a01135b5400` |

C01 records content hashes for the actual fixture inputs. Do not treat a commit alone as the full dirty-checkout baseline.

## S01

**Production workspace kit.** All imported kits define the editor feature scope. Demo-only kits do not add requirements.

- [app:src/components/workspace/WorkspaceEditorKit.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/workspace/WorkspaceEditorKit.ts)

## S02

**Per-path agent bridge.** FileExecutor, working-view reads, read/search semantics, revision invalidation, edit/set routing, validation, and landed echoes.

- [app:src/hooks/chat/useWorkspaceEditBridge.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/hooks/chat/useWorkspaceEditBridge.ts)

## S03

**AI canonicalization and parity fixtures.** The seven before/after cases, symmetric normalization, asset reconciliation, direct baseline reuse, and atomic history.

- [app:src/lib/ai-review-diff.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/ai-review-diff.ts)
- [app:src/lib/ai-review-diff.test.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/ai-review-diff.test.ts)

## S04

**Committed and working suggestion views.** Pure accept/reject projections and old-property presence semantics.

- [app:src/lib/suggestion-views.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/suggestion-views.ts)
- [app:src/lib/suggestion-views.test.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/suggestion-views.test.ts)

## S05

**Agent tool wire contract.** read_document, search_document, edit_document, set_document, is_suggesion, direct default, and complete VFS export expectations.

- [app:supabase/functions/ai-chat_shared/tools/document-edit-tools.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/supabase/functions/ai-chat_shared/tools/document-edit-tools.ts)

## S06

**Exact replacement matching.** Exact, trimmed, and typography matching; count checks; literal replacements; ambiguous-target failures.

- [app:src/lib/agent-edit-find-replace.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/agent-edit-find-replace.ts)
- [app:src/lib/agent-edit-find-replace.test.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/agent-edit-find-replace.test.ts)

## S07

**Review grouping and scroll behavior.** Nested groups, replacement pairing, manual suggestion mode, and SuggestionScrollOrchestrator.

- [app:src/lib/suggestion-groups.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/suggestion-groups.ts)
- [app:src/components/suggestion-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/suggestion-kit.tsx)

## S08

**Review session and decisions.** Fresh committed baselines, multi-wave recomposition, per-group and bulk history, reload recovery, cancellation, and document switches.

- [app:src/components/ai-review-context.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/ai-review-context.tsx)
- [app:src/components/ai-review-context-value.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/ai-review-context-value.ts)

## S09

**Lesson plan structures and identities.** Title fields, recap fields, phase types, activity properties, resource references, and atomic visuals/listening.

- [app:src/components/lesson-plan/types.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/lesson-plan/types.ts)
- [app:src/components/lesson-plan/diff-strategy.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/lesson-plan/diff-strategy.ts)
- [app:src/components/lesson-plan/lesson-plan-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/lesson-plan/lesson-plan-kit.tsx)
- [app:src/components/lesson-plan/md-rules.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/lesson-plan/md-rules.ts)

## S10

**Source chips and excerpts.** Atomic chips, citation identity, provenance actions, trust, and native excerpt children.

- [app:src/components/source/types.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/source/types.ts)
- [app:src/components/source/diff-strategy.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/source/diff-strategy.ts)
- [app:src/components/source/md-rules.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/source/md-rules.ts)
- [app:src/components/ui/source-node.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/ui/source-node.tsx)
- [app:src/components/ui/source-excerpt-node.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/ui/source-excerpt-node.tsx)

## S11

**Markdown and shared MDX.** Complete AST preservation, rule precedence, allowed attributes, literal tags, data-only tags, and preview dispatch.

- [app:src/components/markdown-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/markdown-kit.tsx)
- [app:src/components/mdx/rules.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/mdx/rules.ts)
- [app:src/components/mdx/nodes.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/mdx/nodes.tsx)
- [app:src/components/chat/mdx-tags.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/chat/mdx-tags.ts)

## S12

**Listening and media assets.** Audio references, speaker and line markup, generated media retention, and cache hydration.

- [app:src/components/ui/listening-node.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/ui/listening-node.tsx)
- [app:src/components/lesson-plan/listening-md-rules.test.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/lesson-plan/listening-md-rules.test.ts)
- [app:src/lib/reconcileListeningNodes.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/reconcileListeningNodes.ts)
- [app:supabase/functions/\_shared/activities/audio-exercise/markup.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/supabase/functions/_shared/activities/audio-exercise/markup.ts)

## S13

**Native and media kit definitions.** Start with actual options and UI actions. Expand imports when capturing N01 and C06.

- [app:src/components/basic-blocks-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/basic-blocks-kit.tsx)
- [app:src/components/basic-marks-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/basic-marks-kit.tsx)
- [app:src/components/font-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/font-kit.tsx)
- [app:src/components/list-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/list-kit.tsx)
- [app:src/components/align-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/align-kit.tsx)
- [app:src/components/line-height-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/line-height-kit.tsx)
- [app:src/components/table-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/table-kit.tsx)
- [app:src/components/link-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/link-kit.tsx)
- [app:src/components/code-block-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/code-block-kit.tsx)
- [app:src/components/toggle-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/toggle-kit.tsx)
- [app:src/components/toc-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/toc-kit.tsx)
- [app:src/components/column-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/column-kit.tsx)
- [app:src/components/date-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/date-kit.tsx)
- [app:src/components/mention-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/mention-kit.tsx)
- [app:src/components/math-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/math-kit.tsx)
- [app:src/components/smiles-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/smiles-kit.tsx)
- [app:src/components/media-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/media-kit.tsx)

## S14

**Lesson syntax and schema validation.** Syntax failures, incremental shape warnings, and stricter full-set validation.

- [app:src/lib/lesson-plan-validation.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/lesson-plan-validation.ts)
- [app:src/lib/lesson-plan-validation.test.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/lesson-plan-validation.test.ts)

## S15

**Lesson timing and visual replacement actions.** Computed durations, retiming rules, image ownership, and pristine visual references.

- [app:src/components/lesson-plan/sum-phase-durations.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/lesson-plan/sum-phase-durations.ts)
- [app:src/components/lesson-plan/retime-document.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/lesson-plan/retime-document.ts)
- [app:src/components/lesson-plan/replace-visual.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/lesson-plan/replace-visual.ts)
- [app:src/components/lesson-plan/replace-visual.test.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/lesson-plan/replace-visual.test.ts)

## S16

**Callout, story, and solution blocks.** Container attributes, native child content, reveal controls, and static fallback behavior.

- [app:src/components/callout-md-rules.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/callout-md-rules.ts)
- [app:src/components/story/md-rules.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/story/md-rules.ts)
- [app:src/components/solution/md-rules.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/solution/md-rules.ts)
- [app:src/components/ui/story-node.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/ui/story-node.tsx)
- [app:src/components/ui/solution-node.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/ui/solution-node.tsx)

## S17

**Production editor interaction controls.** Follow each production command to its implementation and capture keyboard bindings before changing them.

- [app:src/components/floating-toolbar-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/floating-toolbar-kit.tsx)
- [app:src/components/block-menu-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/block-menu-kit.tsx)
- [app:src/components/slash-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/slash-kit.tsx)
- [app:src/components/autoformat-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/autoformat-kit.tsx)
- [app:src/components/dnd-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/dnd-kit.tsx)
- [app:src/components/emoji-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/emoji-kit.tsx)
- [app:src/components/exit-break-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/exit-break-kit.tsx)
- [app:src/components/block-placeholder-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/block-placeholder-kit.tsx)
- [app:src/components/cursor-overlay-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/cursor-overlay-kit.tsx)
- [app:src/components/ui/ai-review-bar.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/ui/ai-review-bar.tsx)
- [app:src/components/ui/ai-suggestion-toolbar.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/ui/ai-suggestion-toolbar.tsx)

## S18

**Plate fork AI and suggestion engine.** One direct commit, presentation cancellation, session ledger, 4000 ms cap, paired block diffs, and symmetric word hints.

- [plate:packages/ai/src/react/direct-edit/applyAnimatedAIEdit.ts](/run/media/lofcz/ssd_external/GitHub/plate/packages/ai/src/react/direct-edit/applyAnimatedAIEdit.ts)
- [plate:packages/ai/src/react/direct-edit/planAIEdit.ts](/run/media/lofcz/ssd_external/GitHub/plate/packages/ai/src/react/direct-edit/planAIEdit.ts)
- [plate:packages/ai/src/react/direct-edit/aiChanges.ts](/run/media/lofcz/ssd_external/GitHub/plate/packages/ai/src/react/direct-edit/aiChanges.ts)
- [plate:packages/ai/src/lib/transforms/withAIBatch.ts](/run/media/lofcz/ssd_external/GitHub/plate/packages/ai/src/lib/transforms/withAIBatch.ts)
- [plate:packages/suggestion/src/lib/diffToSuggestions.ts](/run/media/lofcz/ssd_external/GitHub/plate/packages/suggestion/src/lib/diffToSuggestions.ts)
- [plate:packages/diff/src/internal/transforms/pairBlocksWithWordHints.ts](/run/media/lofcz/ssd_external/GitHub/plate/packages/diff/src/internal/transforms/pairBlocksWithWordHints.ts)

## S19

**Selection AI menu and stream hooks.** Menu modes, selection scope, mod+j shortcut, stream insertion anchor, final flush, and request lifecycle.

- [app:src/components/ai-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/ai-kit.tsx)
- [app:src/components/use-chat.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/use-chat.ts)

## S20

**Import, autosave, and stored document state.** Original DOCX assets, existing JSON snapshots, drafts, save status, and immutable export expectations.

- [app:src/services/attachmentWorkspaceApi.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/services/attachmentWorkspaceApi.ts)
- [app:src/lib/autosave.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/autosave.ts)
- [app:src/stores/workspaceStore.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/stores/workspaceStore.ts)

## S21

**Workspace editor lifecycle.** Editor instance ownership, snapshots, generation hydration, tab switches, attachment import, and restoration races.

- [app:src/components/workspace/WorkspacePanel.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/workspace/WorkspacePanel.tsx)
- [app:src/components/workspace/LinkedLessonEditor.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/workspace/LinkedLessonEditor.tsx)

## S22

**Shared and static consumers.** Static and legacy rendering, including stored comments and suggestions. This does not require a new live comment UI.

- [app:src/components/shared-material/SharedLessonDocView.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/shared-material/SharedLessonDocView.tsx)
- [app:src/components/editor-base-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/editor-base-kit.tsx)
- [app:src/components/lesson-plan/lesson-plan-export-kit.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/lesson-plan/lesson-plan-export-kit.tsx)

## S23

**Workspace export contracts.** Immutable revisions, full exports, batch archives, text export defaults, and native modal workflows.

- [app:src/lib/workspaceExportSnapshot.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/workspaceExportSnapshot.ts)
- [app:src/lib/workspaceExportArchive.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/workspaceExportArchive.ts)
- [app:src/lib/workspaceExportArchive.pipeline.test.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/workspaceExportArchive.pipeline.test.ts)
- [app:src/lib/documentTextExport.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/documentTextExport.ts)
- [app:src/components/workspace/WorkspaceExportModal.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/workspace/WorkspaceExportModal.tsx)

## S24

**React block proof.** Proof evidence: nested native content, arbitrary widgets, metadata, suggestions, page movement, and two real collaboration replicas.

- [fork:examples/react-blocks/README.md](../../../examples/react-blocks/README.md)
- [fork:examples/react-blocks/proof-results.json](../../../examples/react-blocks/proof-results.json)
- [fork:examples/react-blocks/src/react-block-layer.tsx](../../../examples/react-blocks/src/react-block-layer.tsx)
- [fork:examples/react-blocks/src/bridge.ts](../../../examples/react-blocks/src/bridge.ts)
- [fork:examples/react-blocks/test/collaboration.test.ts](../../../examples/react-blocks/test/collaboration.test.ts)
- [fork:e2e/react-blocks.spec.ts](../../../e2e/react-blocks.spec.ts)

## S25

**Native engine constraints and entry points.** Canonical store authority, native transaction and tracked-operation seams, lossless parts, layout boundaries, and collaboration verification.

- [fork:AGENTS.md](../../../AGENTS.md)
- [fork:docs/architecture/production-engine-packages.md](../../../docs/architecture/production-engine-packages.md)
- [fork:packages/core/src/editor/surface-paint-observer.ts](../../../packages/core/src/editor/surface-paint-observer.ts)
- [fork:packages/core/src/automation/proposals.ts](../../../packages/core/src/automation/proposals.ts)
- [fork:packages/core/src/automation/text-projection.ts](../../../packages/core/src/automation/text-projection.ts)
- [fork:packages/core/src/automation/content-controls.ts](../../../packages/core/src/automation/content-controls.ts)
- [fork:packages/core/src/store/store/review-items.ts](../../../packages/core/src/store/store/review-items.ts)
- [fork:packages/core/src/store/store/tracked-replacement-landing.ts](../../../packages/core/src/store/store/tracked-replacement-landing.ts)
- [fork:packages/core/src/store/store/table-property-attributes.ts](../../../packages/core/src/store/store/table-property-attributes.ts)
- [fork:packages/core/src/store/store/tree-op-table-cell-properties.ts](../../../packages/core/src/store/store/tree-op-table-cell-properties.ts)
- [fork:packages/core/src/store/**tests**/table-property-fidelity.test.ts](../../../packages/core/src/store/__tests__/table-property-fidelity.test.ts)
- [fork:packages/core/src/store/store/review-export-derivation.ts](../../../packages/core/src/store/store/review-export-derivation.ts)

## S26

**Native Markdown source mapping.** Shared semantic translation, sparse source boundaries, and canonical layout fixtures.

- [Native translator](../../../packages/docx-to-markdown/src/markdown.ts)
- [Inline mapping](../../../packages/docx-to-markdown/src/markdown-inline.ts)
- [Logical table rendering](../../../packages/docx-to-markdown/src/markdown-logical.ts)
- [Public source-map types](../../../packages/docx-to-markdown/src/markdown-source-types.ts)
- [Mapped Markdown operations](../../../packages/docx-to-markdown/src/markdown-source-map.ts)
- [Canonical layout fixture](../../../packages/docx-to-markdown/test/canonical-layout.ts)

## S27

**Selection quotes and chat focus.** The production floating toolbar supplies Add to chat. Working-view quote lines must match agent reads. Quote clicks activate or hydrate the document before native focus.

- [app:src/components/ui/floating-toolbar-buttons.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/ui/floating-toolbar-buttons.tsx)
- [app:src/lib/quote-focus.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/quote-focus.ts)
- [app:src/lib/attemptAddToChat.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/lib/attemptAddToChat.ts)
- [app:src/stores/quoteStore.ts](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/stores/quoteStore.ts)
- [app:src/components/chat/renderer/UserMessageWithInlineMarkers.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/chat/renderer/UserMessageWithInlineMarkers.tsx)
- [app:src/components/chat/chat-input/EdixTextarea.tsx](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/chat/chat-input/EdixTextarea.tsx)

## S28

**Cell merge and unmerge semantics.** The application uses its table toolbar merge/split actions. Plate restores original grid slots and keeps combined content in the first cell. Native equal-width splits can add grid cuts.

- [Application table toolbar](/run/media/lofcz/ssd_external/GitHub/sciobot-next/src/components/ui/table-toolbar-button.tsx)
- [Plate merge transform](/run/media/lofcz/ssd_external/GitHub/plate/packages/table/src/lib/merge/mergeTableCells.ts)
- [Plate split transform](/run/media/lofcz/ssd_external/GitHub/plate/packages/table/src/lib/merge/splitTableCell.ts)
- [Native split implementation](../../../packages/core/src/store/store/tree-op-table-split.ts)
- [Native merge implementation](../../../packages/core/src/store/store/tree-op-table-merge.ts)

- [Native node and parent indexes](../../../packages/core/src/store/package/ooxml-edit.ts)
- [Native index oracle](../../../packages/core/src/store/__tests__/ooxml-edit-index.test.ts)
