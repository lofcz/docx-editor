# Direct AI edits and streaming lifecycle

Direct edits commit before animation. Presentation never owns document state.

Read the [dispatch and completion rules](../../../SCIOBOT_PORT_TODO.md) before claiming a leaf.

Each named test or artifact is a deliverable. These specifications do not claim those checks already exist.

## A01

- [x] Apply a direct plan in one transaction

**Depends:** [D15](04-edit-planner.md#d15), [C07](01-contracts.md#c07), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/ai/apply-direct.ts`; `sibling apply-direct.test.ts`

**Read:** [S03](sources.md#s03), [S18](sources.md#s18), [S25](sources.md#s25)

**Change:** Apply a validated plan with one history boundary and one committed notification. Roll back completely if any operation fails.

**Close when:** C01 pairs each create one undo entry and no transient review items. Undo and redo restore exact semantic results.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `packages/sciobot/src/ai/apply-direct.ts` adds `applyDirectPlan(editor, plan, { documentId, epoch })`, exported from the private entry. It decodes a detached copy of the plan, checks the mount revision, and refuses view and suggestion modes without changing them (direct edits never disable tracking). It rechecks every old value against the live canonical package: paragraph text slices, run property bags (shared `trackedRunFormatOps` lowering without attribution), paragraph properties, hyperlink targets/tooltips, exact sibling gaps, list source fingerprints (re-read through `planListMembership`), table row anchors, and payload values at their JSON paths. Existing-ID work lowers to native operations; identity-dependent work (links and list membership of new paragraphs, new rows/columns with their cell paragraphs, payload records, list instances via `direct-lists.ts`) runs in the existing after-allocation package-edit phase. A complete immutable preview, including relationship mints and completions, runs before the live write, so a refused plan leaves the document, history, and review state untouched. The live write is one `applyAutomationOps` call with `separateHistory`, so it publishes once and creates one undo unit. Focused run: `packages/sciobot/src/ai/apply-direct.test.ts`, 19 tests, all passing. All seven C01 pairs, inside the imported document with header, footer, notes, media, and unknown parts: one publication, revision +1, no review items, empty replanning against the target, live native Markdown equal to the immutable oracle, one undo restoring the exact canonical fingerprint with `canUndo()` false, redo restoring the landed fingerprint, unchanged header/footer/notes/unknown parts versus an unedited save, identical media bytes, no `w:ins`, and exact reopen. Six further plan kinds (run marks, heading style, new bullet list, checklist creation/state, inserted ordered list with a new instance, inserted linked paragraph) pass the same history and reopen checks. Refusals: wrong epoch, invalid context, suggestion mode, view mode, forged same-length old text, and a plan whose completion fails in preview; each leaves fingerprint, history, and review items unchanged. An empty plan commits nothing. Two real sessions (both receivers) run direct C01 complete rewrite, tables/links (relationship mint plus new row), and nested MDX (payload record) beside an offline peer edit: both converge on the target plus peer text, no review items, remote caret geometry follows the peer paragraph, undo removes only the direct transaction and keeps the peer text, redo restores it, deleting the peer paragraph during concurrent editing converges, and saved semantic digests match. Full private suite during this leaf: 1060 pass; the one failure was another worker's spike test that was deleted mid-run (ENOENT). Private typecheck has no errors in these files (remaining errors belong to the concurrent R08 worker); private build verifies 112 modules with the core external; scoped oxlint and prettier pass; compatibility passes with `.collaboration/changes/sciobot-direct-plan-execution.json`. Limits: inserted rich blocks, registered blocks in new rows/columns, nested inserted list paths, and tooltip removal refuse here and belong to D27-D30. A02-A09 keep request-session ledgers and presentation.

## A02

- [x] Record direct AI changes by request session

**Depends:** [A01](06-direct-ai.md#a01)

**Primary files:** `fork:packages/sciobot/src/ai/change-ledger.ts`; `sibling change-ledger.test.ts`

**Read:** [S18](sources.md#s18)

**Change:** Track changed ranges relative to the request baseline. Merge repeated edits and settle previous highlights when a new request starts.

**Close when:** Two edits to one activity form one change. A new session settles earlier highlights without accepting unrelated Word review items.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `packages/sciobot/src/ai/change-ledger.ts` adds `DirectChangeLedger`, the native counterpart of the Plate `AIChangeLedger`. Callers run `applyDirectPlan` inside `ledger.mute(...)` and then `ledger.record(partBefore, { session })`. Changes are recomputed against the request-session baseline from canonical paragraph fingerprints and grouped by the nearest registered block control, otherwise by top-level body block, so repeated edits merge. Statuses, pending playback (`reveal`), `acknowledge` with fade, `settle`, and `settleAll` follow the Plate contract. Any unmuted commit (typing, undo, review decision, remote revision) seals the baseline and drops only the changes it touched. A new session settles earlier changes. The ledger only reads fingerprints; it never writes the document or decides review items. Focused run: `change-ledger.test.ts`, 4 tests, 33 assertions. Two direct edits to two paragraphs of one C01 activity in one request form one change owned by `activity`, with a stable ID; a third edit of the same paragraph keeps it. A second request settles the first request's highlight while an imported tracked insertion keeps the same review item, also after `settleAll`. Removal-only edits produce no change; inserted paragraphs are `insert` changes; pending reveal by node and for all; acknowledgement fades and folds content into the baseline so a later edit highlights only new text. User typing drops only the touched change and keeps the manual text through a later direct edit; undo of the direct edit clears its change. Private typecheck shows no errors in these files; oxlint and prettier pass; compatibility decision `.collaboration/changes/sciobot-direct-change-ledger.json`. Limit: highlight painting and acknowledgement wiring belong to A03/A04/U leaves.

## A03

- [x] Acknowledge highlights after user edits

**Depends:** [A02](06-direct-ai.md#a02), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/ai/acknowledge.ts`; `sibling acknowledge.test.ts`

**Read:** [S18](sources.md#s18)

**Change:** Map ledger references through native operations. Settle touched changes and preserve unaffected highlights.

**Close when:** Typing in a highlighted block settles that change only. Undo, deletion, and remote edits never leave dangling highlight ranges.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Native IDs are stable through operations, so ledger references need no positional remapping. A02's ledger keys changes by paragraph ID and fingerprint and drops exactly the changes an unmuted commit touched. `packages/sciobot/src/ai/acknowledge.ts` adds `directHighlights(editor, ledger)`: it flushes queued typing first (that commit settles touched changes) and resolves visible changes to whole live paragraph ranges, skipping pending playback and any vanished node. Focused run: `acknowledge.test.ts`, 3 tests. Real `surface.type` in one of two highlighted paragraphs settles only that change; the other keeps its change ID and every range addresses a live paragraph at its full length. Deleting the remaining highlighted paragraph clears it. Undo of that deletion does not resurrect a stale highlight. A later direct edit highlights only its new paragraph, and undo/redo of it leaves no dangling range. A reopened document has no highlights and keeps the text. Two real sessions (both receivers): a remote insertion in a highlighted paragraph settles that change on the author's peer while preserving the other; a remote deletion clears the last; saved digests match. Private typecheck shows no errors in these files; oxlint passes; compatibility decision `.collaboration/changes/sciobot-direct-highlights.json`. Limit: painting the ranges and fade styling belong to A04/U leaves.

## A04

- [x] Render bounded AI edit playback

**Depends:** [A02](06-direct-ai.md#a02), [B06](02-react-blocks.md#b06), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/ai/playback.ts`; `sibling playback browser test`

**Read:** [S18](sources.md#s18), [S24](sources.md#s24)

**Change:** Animate committed changed ranges, reveal targets, and show the AI caret. Keep playback within the existing 4000 ms cap.

**Close when:** The full target is readable before playback finishes. Cancelling playback retains the edit. Deletions create no stale highlight.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `packages/sciobot/src/ai/playback.ts` adds `playDirectChanges(editor, ledger, options)`. Callers commit through `applyDirectPlan` inside `ledger.mute`, record with `pending: true`, and then start playback. The edit is already canonical and painted, so overlays only tint changed paragraphs and never hide text. Changes play in document order. Each step reveals its first paragraph through `surface.revealParagraph` (falling back to the painted fragment when the host lacks `.docx-editor__scroll-container`), sweeps a fixed-position overlay across the paragraph's painted line rectangles, moves a labelled AI caret along them, and then reveals that change in the ledger. Each frame re-reads the canonical part, so a paragraph that the user, undo, or a peer deletes loses its overlay at once; a change whose nodes vanished is revealed without painting. The budget is `clamp(duration ?? 2400, 0, 4000)`. A newer playback on the same editor cancels the older one and inherits its start, so a burst shares one deadline. Cancel and abort remove every overlay, reveal the rest, and never write the document; reduced motion reveals at once. Chromium (`e2e/sciobot-playback.spec.ts` with `e2e/sciobot-playback.config.ts`, harness `packages/sciobot/test/browser/playback.tsx`, two real rooms): 3 cases pass. (1) A projection-derived target edits two paragraphs 33 paragraphs apart and deletes one. Before any step completes, both new texts are in the canonical part and painted, the deleted text is gone, one sweep and the `✦ ScioBot` caret show, and changes are pending. A requested 10 s budget ends between 1 and 4.4 s. Afterwards no overlay remains, nothing is pending, both edited paragraphs are highlighted and the deleted one is not, and the far target was scrolled into view (pane `scrollTop` 1052). (2) Real clicking and typing into a changed paragraph during playback, then cancel: no overlays, the typed and committed text stays, the typed change settles, and the other change stays highlighted. (3) The peer deletes a played paragraph mid-playback: no overlay or highlight remains, both peers hold the same paragraphs, and saved digests match. Bun (`src/ai/playback.test.ts`): 4 cases, 43 assertions. A 60 s request plus a second burst commit ends inside `4000 + 500` ms with all three changes revealed. Cancel keeps the package revision. Reduced motion plays nothing. Undo plus abort during playback leaves no highlight for the undone text. Two real sessions in both receiver orders: a peer edit and a peer deletion during playback converge (`Alpha changed.`, `Gamma. Bob`), only the surviving change is highlighted, Alice's undo keeps the peer's work, and saved digests match. Scoped Prettier/oxlint pass. Private typecheck reports no errors in these files; its remaining errors are in the R worker's in-progress `track-table-format.ts` and core `tree-op-tracked-table-properties.ts`. Compatibility: `.collaboration/changes/sciobot-direct-playback.json`. Limits: playback tints whole paragraphs and does not animate removed blocks as ghosts; A05 owns document-switch and detach cancellation evidence.

## A05

- [x] Honor reduced motion and playback cancellation

**Depends:** [A04](06-direct-ai.md#a04), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/ai/playback-lifecycle.ts`; `sibling playback-lifecycle browser test`

**Read:** [S18](sources.md#s18)

**Change:** Honor reduced motion and abort signals. Cancel on document switch, detach, or newer playback; clean overlays and timers.

**Close when:** Reduced motion leaves no animation delay. Abort and document switch leave no overlay or canonical edit rollback.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Engine baseline `9770e78f` plus local work. `watchPlaybackLifecycle` observes editor document/selection/host events, native host disposal, and system motion preference changes. Playback stops on detach, document replacement, destruction, explicit abort, a newer playback, or reduced motion. Cleanup releases event listeners, MutationObserver, preference listener, animation frame, timer, sweep and caret nodes. Reduced motion reveals immediately; cancellation preserves the committed edit and history. Document replacement also forgets old presentation highlights, so matching paragraph IDs cannot tint an unrelated file. The root private package exports `playDirectChanges`. `bun test --preload ./scripts/test/preload-dom.ts --timeout 30000 packages/sciobot/src/ai/playback.test.ts`: 9 pass, 78 assertions; `/tmp/codex-docx-playback-lifecycle-final.log`. The final switch/highlight regression also passes separately (1 pass, 6 assertions; `/tmp/codex-docx-playback-switch.log`). Tests prove immediate completion through microtasks instead of waiting out the timer. Two real sessions cover both reconnect receivers, detach beside offline peer typing, independent undo/redo, shared saved convergence, and deletion during playback. `bunx playwright test --config e2e/sciobot-playback.config.ts`: 6 pass; `/tmp/codex-docx-playback-browser.log`. Chromium covers system reduced motion with real native typing, real Abort/Detach/Attach buttons, changing motion preference during playback, the 4000 ms cap, cancellation during typing, and peer deletion. Private typecheck/build, scoped formatting/lint, and compatibility policy pass. No canonical write or review decision occurs in lifecycle cleanup.

## A06

- [ ] Coalesce streamed suggestion targets safely

**Depends:** [R19](05-suggestions.md#r19), [C07](01-contracts.md#c07), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/ai/review-stream.ts`; `sibling review-stream.test.ts`

**Read:** [S08](sources.md#s08), [S02](sources.md#s02)

**Change:** Keep only the latest complete target per animation frame. Flush the final target and group one stream into one AI history batch.

**Close when:** Three same-frame targets apply only the last target. A final target applies once. Undo returns to the pre-stream document.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## A07

- [ ] Cancel stale suggestion callbacks

**Depends:** [A06](06-direct-ai.md#a06), [A01](06-direct-ai.md#a01), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/ai/session-lifecycle.ts`; `sibling session-lifecycle.test.ts`

**Read:** [S08](sources.md#s08), [S02](sources.md#s02)

**Change:** Guard queued callbacks by document identity and session generation. Cancel queued review work before a direct edit or document switch.

**Close when:** A delayed review frame cannot overwrite a newer direct edit. Ending a session flushes or cancels its target exactly once.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## A08

- [ ] Insert generated content at a stable AI anchor

**Depends:** [A10](06-direct-ai.md#a10), [D05](04-edit-planner.md#d05), [A01](06-direct-ai.md#a01), [B11](02-react-blocks.md#b11), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/ai/stream-insert.ts`; `sibling stream-insert.test.ts`

**Read:** [S19](sources.md#s19)

**Change:** Insert generated content after the selected blocks or caret scope. Keep the anchor outside serialized document content.

**Close when:** Generate inserts after the selected activity. Cancel and finish remove the anchor. One undo removes the generated insertion.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Implementation in progress from `69466941`. `packages/sciobot/src/ai/stream-anchor.ts` adds an internal request-owned native scope. Selected sibling order follows the document, not argument order. Registered scopes resolve through persisted block IDs; paragraphs resolve through Word paraIds. The anchor follows peer activity moves and clears after source deletion, explicit release, or native surface replacement/destruction. No placeholder enters canonical content, history, or saved Word output. Mixed-parent and duplicate selections refuse before allocating an observer. Paint observation reads the current model without flushing pending input.

Command: `bun test --preload ./scripts/test/preload-dom.ts packages/sciobot/src/ai/stream-anchor.test.ts`; three cases and 46 assertions pass in `/tmp/a08-anchor-final.log`. Both reconnect receivers cover concurrent activity movement and independent peer text, undo/redo, deletion, and exact save/reopen digest convergence. The native move/removal regression command passes 38 cases and 709 assertions across three files in `/tmp/a08-native.log`. Private typecheck passes 2044 files in `/tmp/a08-types-last.log`. Root formatting, scoped formatting, root lint, diff checking, backlog validation, and compatibility policy pass. Logs: `/tmp/a08-root-format.log`, `/tmp/a08-format-check.log`, `/tmp/a08-lint.log`, `/tmp/a08-backlog.log`, `/tmp/a08-policy.log`. The first test attempt used the wrong removal option; the corrected test uses the existing native delete mode. Decision: `.collaboration/changes/sciobot-stream-anchor.json`; no shared format fields change.

A08 stays unchecked. The `ef388d37` follow-up implements `stream-insert.ts` and exports `beginStreamInsertion` from the private document entry. Completed D05 fragments append at the live scope through `applyDirectPlan` and native history frames. Rich text and hyperlink content survive full Word reopen. Finish and cancel release the request; already committed output stays undoable. Stable paragraph identities and registered control tags retain peer text inside generated output. Both reconnect receivers preserve concurrent activity text, move following native geometry equally, place the peer caret consistently, and converge after save/reopen. One undo restores the exact independently captured peer-only package. Redo restores the complete landed package. Scope deletion during peer editing rejects late chunks without writes.

Focused verification passes eight anchor/insertion cases and 121 assertions in `/tmp/a08-stream-final.log`. The later split-outcome regression passes six insertion cases and 84 assertions in `/tmp/a08-split.log`. The adapter reports native history outcomes explicitly and ends on a split. An intervening local gesture produces such a split; its test verifies that undo preserves the teacher's intervening text. This is recorded behavior, not proof of a single undo across that boundary. Types pass 2046 files; the private build verifies 212 modules with external core. Logs: `/tmp/a08-insert-types.log`, `/tmp/a08-insert-build.log`. Root formatting, scoped formatting, root lint, diff checking, and compatibility policy pass. Logs: `/tmp/a08-insert-root-format.log`, `/tmp/a08-insert-scoped-format.log`, `/tmp/a08-insert-lint.log`, `/tmp/a08-insert-policy.log`. The completed affected run covers thirty cases across the direct executor, anchors, and insertion. It records 29 passes and one 5-second timeout: `/tmp/a08-insert-affected.log`. The unchanged receiver-0 case passes in isolation at the same limit, with twenty assertions in 1.22 seconds: `/tmp/a08-insert-receiver0.log`. No expected result or deadline changed. This combines separate verification results; the thirty-case process itself exits nonzero. Decision: `.collaboration/changes/sciobot-stream-insert.json`; existing operations, history tokens, and shared formats remain unchanged.

The `e2b38fbf` follow-up adds `beginMarkdownInsertion`, shared generated-fragment policy, and structural scopes. Decoded text deltas retain unfinished syntax outside the document. Completed Markdown blocks append progressively. Finish and stop flush a valid final block without requiring a terminal newline. Received content stays undoable after stop. Split UTF-16 surrogate pairs retain the exact character; oversized deltas refuse before writes. The stream parser uses D05 normalization, registered payload/list bindings, and quote styles. It shares the extracted policy in `diff/inserted-fragment-policy.ts` with `planTargetDocument`, rather than creating a competing native conversion path. Native canonical identities now support table scopes, caret scopes inside tables, and unregistered native controls. Those identities end on document replacement instead of guessing from text.

Six structural-scope cases and 102 assertions pass: `/tmp/a08-struct.log`. Both reconnect receivers cover peer text, row insertion, exact placement, complete peer-only undo, redo, deletion during peer editing, and saved convergence. Six raw-stream cases pass in `/tmp/a08-final-repro.log`: incremental rich output, final flushing, stop, malformed input, byte limits, surrogate pairs, and both plain-text reconnect cases. Two additional rich-list reconnect cases intentionally remain red. Their explicit modeled-numbering-part assertion fails after shared undo. A10 owns that prerequisite. Final private types pass 2051 files in `/tmp/a08-parser-types-final.log`. The private build verifies 215 modules with external core: `/tmp/a08-parser-build.log`. Root format, scoped format, root lint, diff checking, backlog validation, and compatibility policy pass. Logs: `/tmp/a08-parser-root-format.log`, `/tmp/a08-parser-scoped-format.log`, `/tmp/a08-parser-lint.log`, `/tmp/a08-parser-backlog.log`, `/tmp/a08-parser-policy.log`. The sequential affected-source command passes fifty cases and 639 assertions across six files, including original D05 block, quote, and target-planning regressions plus anchor, insertion, and structural scopes: `/tmp/a08-parser-affected.log`. No red raw-stream case was skipped in its full eight-case reproduction.

The first local rich-stream test incorrectly expected byte-identical undo with no numbering resources. The existing native shell policy intentionally retains numbering allocations. The corrected oracle independently allocates the exact allowed bullet resource and relationship slot, removes the transient hyperlink, and compares the complete expected semantic package, content types, and external targets. Shared undo has a different explicit policy: keep an empty numbering directory/root, type, and relationship while undoing the author's definitions. It instead saves malformed empty `abstractNum` and `num` shells, and reopen omits the modeled numbering part. The exact schema and reproduction remain under A10; no acceptance oracle was weakened to permit malformed output.

A08 stays unchecked until A10 passes. Application chat transport, AI menu wiring, and document-switch request cleanup remain U13 and the integration leaves. The new native API is not a claim of production Generate cutover. Changes to generated root ownership still refuse before further writes; local history boundaries report their native split explicitly.

## A09

- [x] Reuse canonical direct targets until document changes

**Depends:** [A01](06-direct-ai.md#a01), [P09](03-agent-projection.md#p09), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/ai/direct-baseline.ts`; `sibling direct-baseline.test.ts`

**Read:** [S03](sources.md#s03), [S02](sources.md#s02)

**Change:** Reuse the landed normalized target across consecutive direct calls. Invalidate on manual edits, undo, review decisions, and remote revisions.

**Close when:** The existing consecutive-tool-call fixture passes. A manual sentence between calls remains present after the next direct edit.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `packages/sciobot/src/ai/direct-baseline.ts` adds `DirectTargetBaseline`. Native planning needs the document's real source map, so the reusable baseline is the working projection a direct edit landed, not the requested Markdown (native projection normalizes it). `remember` reads that projection once after a commit, and P13 echoes also need it. `prepare` reuses it only while the editor still holds the identical canonical package at the same revision token, as Plate compares tree identity. Typing, undo, redo, review decisions, and remote revisions all replace the package or advance the revision, so the next call reads a fresh projection. Focused run: `direct-baseline.test.ts`, 4 tests, 35 assertions. Port of the application consecutive-tool-call fixture (`Opening/Middle/Closing`): the second call reuses the landed projection object, whose Markdown equals a fresh read of a document holding the intermediate target, and lands `Closing, revised.`. A manual sentence between calls invalidates the landing; the next direct edit plans against the live document and the sentence remains. Undo and redo invalidate, save/reopen keeps the text, and accepting an imported review item invalidates. Two real sessions (both receivers): a remote peer edit invalidates Alice's landing; her next direct call keeps Bob's text on both peers; undo/redo of her transaction keeps it; saved digests match. Private typecheck shows no errors in these files; oxlint and prettier pass; compatibility decision `.collaboration/changes/sciobot-direct-baseline.json`.

## A10

- [ ] Export valid numbering infrastructure after shared undo

**Depends:** [V01](11-verification.md#v01)

**Primary files:** `fork:packages/pro/src/collaboration/document/history-undone.ts`; `fork:packages/pro/src/collaboration/document/materialize-part-members.ts`; `fork:packages/sciobot/src/ai/markdown-stream.test.ts`

**Read:** [S25](sources.md#s25); `reference/ecma-376/part1/schemas/wml.xsd` types `CT_AbstractNum`, `CT_Num`, and `CT_Numbering`; existing provider `document/undo-numbering.ts`

**Change:** Keep the shared numbering directory entry, content type, and relationship after undo, as the provider requires. Do not export undone definition shells without their required IDs. Preserve peer definitions, live references, redo identity, and pre-existing opaque import data. Use the existing undo ownership model before adding shared fields.

**Close when:** Undo of a streamed generated activity with a nested list saves a valid empty numbering root when no definitions remain. Both replicas save the same complete expected package. Redo restores the complete generated content and valid definitions. A peer's independent list and shared references remain resolvable. No new malformed or dangling numbering resource appears.

**Check:** Retain the exact numbering fingerprint and explicit modeled-part assertion in the two `rich=true` raw stream cases. Cover both reconnect receivers, concurrent peer list creation/reference use, redo, and saved Word reopen. Run the affected numbering adoption/history tests and the collaboration compatibility check.

**Claim:** unclaimed. **Evidence:** Reproduced while implementing A08 from `e2b38fbf`. Provider `undo-numbering.ts` deliberately retains numbering infrastructure. After undo, the live package instead shows `<w:numbering><w:abstractNum/><w:num/></w:numbering>`. The pinned schema requires `abstractNumId` on `CT_AbstractNum`, and `numId` plus an `abstractNumId` child on `CT_Num`. Saved bytes retain those malformed shells, but the saved content types also omit the numbering override. Save/reopen therefore does not yield a modeled numbering part. See [saved-state evidence](evidence/a10-numbering-undo.json), captured by `/tmp/a10-word-repro.log`. The existing semantic digest omits this invalid part, so it is insufficient alone. A08's exact modeled-part/fingerprint assertions retain the failure. The provider's member adoption and undo-listing handling are the next inspection points. This is a test-transport defect under the EigenPal Pro License; production dependency boundaries stay unchanged.
