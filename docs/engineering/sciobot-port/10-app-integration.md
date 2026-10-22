# Application persistence and editor replacement

Work in isolated application branches. Keep unrelated dirty files intact. Preserve existing per-path tools and export workflows.

Read the [dispatch and completion rules](../../../SCIOBOT_PORT_TODO.md) before claiming a leaf.

Each named test or artifact is a deliverable. These specifications do not claim those checks already exist.

## I01

- [x] Mount the native editor behind the workspace flag

**Depends:** [C03](01-contracts.md#c03), [B04](02-react-blocks.md#b04), [B05](02-react-blocks.md#b05), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/DocxWorkspaceEditor.tsx`; `workspace selector and package dependency`

**Read:** [S01](sources.md#s01), [S21](sources.md#s21), [S25](sources.md#s25)

**Change:** Add a lazy native editor route using packaged fonts, styles, and private adapter state. Keep instance ownership explicit.

**Close when:** The flagged workspace opens a DOCX and accepts native typing. Unmount releases its instance and React roots.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, app branch `claude/docx-native-editor`. **Evidence:** Application worktree `/run/media/lofcz/ssd_external/GitHub/sciobot-next-wts/docx-native`, local commit `59701be2a` on master `9dec0ffb3`; not pushed. Engine dependency: `scripts/docx-native/vendor.mjs` snapshots this fork outside the shared checkout, builds core, fonts, i18n, docx-to-markdown, and `@sciobot/docx`, then packs content-addressed tarballs into `vendor/docx-native`. `file:` dependencies plus `overrides` keep one core copy. `SOURCE.json` records fork ref `refs/backup/codex-wip-2026-10-20` (`9c91ec4ff`); refresh it from a consistent fork revision as later leaves land. `WorkspacePanel` delegates to `WorkspaceEditorSelector`. That selector renders the lazy `DocxWorkspacePanel` only with the `docx_native_editor` grant or the development-only `localStorage["sciobot:docx-native"]` opt-in. The flag key is in both catalogs, with migration `20261021170018_docx_native_editor_feature.sql` from the helper. `DocxWorkspaceEditor` owns one `createDocxEditor` instance and one `mountReactBlockHost` per opened document. It loads packaged default fonts once and uses `docx-paginated-surface`/`docx-editor` host scoping. Release order is claim, host dispose, editor destroy, then the DOM container. `docxWorkspaceInstance` keeps the active claim, so a stale release cannot drop a newer instance; legacy Plate refs stay untouched. Opening a Word file loads its original bytes into a new session without Plate conversion. Browser harness: `scripts/docx-native/harness`; `bun test --timeout 600000 scripts/docx-native/harness/workspace.test.ts` passes 5 cases and 37 assertions in real Chromium with StrictMode. They cover: no engine requests without the flag; real file-input open; typing, Enter, and Czech text; Ctrl+Z/Ctrl+Y; reopening saved bytes with an identical semantic digest; prior-instance destruction; and unmount releasing the editor, claim, painted pages, and every block root (mounted = unmounted), with a fresh root on remount. Two-session cases in both client orders (73/41, 41/73) use two app editors on real rooms. They type with the keyboard, reconnect after offline concurrent edits, run peer-preserving undo/redo, and unmount the peer while the other session types; the peer then rejoins. Offline deletion of a paragraph being extended converges to identical text and saved digests, and saved bytes reopen. `rstest` ownership cases pass 3 cases. Full app `bun check -p tsconfig.json` passes with no errors in 6,484 files; harness `bun check -p scripts/docx-native/tsconfig.json` passes 452 files. Production `rsbuild build` passes: the engine is only in async chunks, and fonts and wasm are emitted. Scoped oxlint shows no new warnings; `check:migrations` passes. Limits: no persistence, autosave, or generation hydration (I07-I09, I18); no FileExecutor (I02+); no app toolbar wiring (U leaves). The evaluated collaboration transport stays test-only under the EigenPal Pro License, and no collaboration geometry changes. See [implementation notes](evidence/implementation-notes.md#application-workspace-mount).

## I02

- [x] Connect read and search through FileExecutor

**Depends:** [I01](10-app-integration.md#i01), [P07](03-agent-projection.md#p07), [P08](03-agent-projection.md#p08), [P09](03-agent-projection.md#p09), [R02](05-suggestions.md#r02), [V01](11-verification.md#v01)

**Primary files:** `app:src/hooks/chat/docxFileExecutor.ts`; `sibling docxFileExecutor.test.ts`

**Read:** [S02](sources.md#s02), [S05](sources.md#s05)

**Change:** Implement the read and search portions of the existing per-path executor. Read the live working view.

**Close when:** C02 read and search outputs match through the bridge. Pending suggestions appear in reads without resolving them.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, application branch `claude/docx-native-editor`. **Evidence:** Application baseline `f32fefc83`, preserved preceding worktree in `refs/backup/codex-docx-app-r38-final` (`72b1b266b14f047d5bee36555e197ed7edb800fe`); native runtime remains snapshot `a463b2b0697c0ad04c9bf1c3a51ded5f56639c09`. `docxFileExecutor.ts` reads the live native working view asynchronously. It flushes user input, checks canonical revision and owned document/session identity before publishing, keeps no Markdown accumulator, and never resolves proposals. The registry copies document/path/kind and allocates a new mount epoch on replacement. The route separates document sessions and keeps legacy `lesson` identity stable when transient draft metadata clears. The bridge waits for matching native ownership and never falls back to stale Plate refs. Shared wire normalizers retain read defaults and search caps. Native calls submit once; stream cleanup cannot end an unrelated Plate review or replace its dirty baseline. The executor loads only on demand. `src/hooks/chat/__fixtures__/docx-read-search.json` exactly matches the frozen C02 read/search records. Executor and mounted-hook cases assert all eight C02 examples. Invalid-regex suffixes retain the existing host RegExp message: Bun and V8 differ. Command: `bunx rstest run src/hooks/chat/docxFileExecutor.test.ts src/hooks/chat/useWorkspaceEditBridge.native.test.tsx src/components/workspace/docx/docxWorkspaceInstance.test.ts src/hooks/chat/useWorkspaceEditBridge.test.tsx src/hooks/chat/useWorkspaceEditBridge.focus.test.tsx src/hooks/chat/useWorkspaceEditBridge.games.test.tsx src/hooks/chat/useWorkspaceEditBridge.variants.test.tsx`; 51 pass, seven files, `/tmp/codex-i02-unit-complete.log`. This includes legacy document/game/worksheet routes, async replacement/typing refusals, owner readiness, a 30-line search cap, and native-turn cleanup isolation. Chromium command: `DOCX_NATIVE_SNAPSHOT=/home/lofcz/.cache/sciobot-docx-native/fork-n20 DOCX_HARNESS_PORT=8098 DOCX_HARNESS_URL=http://127.0.0.1:8098 bun test --timeout 600000 scripts/docx-native/harness/file-executor.test.ts`; eight pass, 80 assertions, `/tmp/codex-i02-browser-complete.log`. Real native input, review reads without mutation, user/decision undo/redo, pending DOCX reopen, epoch replacement, path switches, unflagged bridge chunk isolation, and legacy metadata retention pass. Four real-session cases cover both client orders and reconnect receivers, concurrent peer text/proposals, remote deletion during local input, deletion undo, both active readers, and saved convergence with both rooms ready. Existing workspace browser cases also pass all five scenarios in `/tmp/codex-i02-browser-final.log`. That earlier combined run had one new test-oracle failure: the unmodified Plate serializer retains its terminal LF. Byte-identical baseline read/serialization functions confirmed the correct expected result; the corrected case passes in both `/tmp/codex-i02-browser-legacy-final.log` and the clean eight-case run. Full application typecheck passes 6620 files; final harness typecheck passes 2186. `bun run build` passes with the native executor only in async chunk `dist/static/js/async/y7.86c52b3299.js`; all four initial scripts exclude it. Build retains the unrelated MapLibre expression warning. Scoped formatting/lint and diff checks pass; lint retains the existing unawaited submission warning in the legacy bridge. Compatibility policy passes 93 decisions. Decision: `.collaboration/changes/sciobot-file-executor-reads.json`; no public API, shared format, provider, database, or native runtime changes. The evaluated provider remains test-only under the EigenPal Pro License. Native direct edits close under I03; suggestion/set routing remain I04-I05; persistence and initial material hydration remain I07-I09.

## I03

- [x] Connect direct edit_document calls

**Depends:** [I02](10-app-integration.md#i02), [C04](01-contracts.md#c04), [P10](03-agent-projection.md#p10), [P11](03-agent-projection.md#p11), [P13](03-agent-projection.md#p13), [A01](06-direct-ai.md#a01), [A09](06-direct-ai.md#a09), [V01](11-verification.md#v01)

**Primary files:** `app:src/hooks/chat/docxEditExecutor.ts`; `sibling docxEditExecutor.test.ts`

**Read:** [S02](sources.md#s02), [S05](sources.md#s05), [S06](sources.md#s06)

**Change:** Route default edit calls through validation, targeted planning, direct application, and landed-result echoes.

**Close when:** Two chained C02 edits succeed using the first echo. Count mismatch, invalid syntax, and stale targets leave the document unchanged.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, application branch `claude/docx-native-editor`. **Evidence:** Application baseline `f32fefc83` plus preceding preserved snapshot `refs/backup/codex-docx-app-i02-final` (`3bd1955168aab11ee32491cc03d741432c928d09`). Native runtime remains `a463b2b0697c0ad04c9bf1c3a51ded5f56639c09`; no engine source or vendor package changes. `docxEditExecutor` uses C04 matching, MDX normalization, P11 material-specific validation, canonical target planning, and A01's gated single-transaction executor. A per-session `DirectTargetBaseline` reuses only exact canonical landings. Ownership/revision checks precede writes; manual input, remote changes, and history invalidate stale targets. P13 echoes use the actual landed projection. A post-commit read failure returns success with a re-read instruction rather than asking the model to repeat an applied write. A shared queue preserves received native read/write order and isolates new document sessions. The existing bridge formatter retains success strings, schema warnings, failure hints, assets, and direct-default wire behavior. Explicit suggestions stay separate; direct edits never disable native tracking. Six copied C02 records match the frozen source exactly and assert complete results. Executor tests also cover chained echoes, cached session ownership, stale sources/plans with zero writes, and honest post-commit fallback. Queue cases cover write/read order, recovery after failure, and separate sessions. Command: `bunx rstest run src/hooks/chat/docxEditExecutor.test.ts src/hooks/chat/docxFileRouting.test.ts src/hooks/chat/docxFileExecutor.test.ts src/hooks/chat/useWorkspaceEditBridge.native.test.tsx src/components/workspace/docx/docxWorkspaceInstance.test.ts src/hooks/chat/useWorkspaceEditBridge.test.tsx src/hooks/chat/useWorkspaceEditBridge.focus.test.tsx src/hooks/chat/useWorkspaceEditBridge.games.test.tsx src/hooks/chat/useWorkspaceEditBridge.variants.test.tsx`; 65 pass, nine files, `/tmp/codex-i03-unit-complete.log`. Chromium command: `DOCX_NATIVE_SNAPSHOT=/home/lofcz/.cache/sciobot-docx-native/fork-n20 DOCX_HARNESS_PORT=8098 DOCX_HARNESS_URL=http://127.0.0.1:8098 bun test --timeout 600000 scripts/docx-native/harness/direct-edits.test.ts`; eight pass, 89 assertions, `/tmp/codex-i03-browser-initial.log`. The C02 source document receives two calls using the first native landed echo. A teacher sentence between them survives. Native escaping is asserted explicitly. Count mismatch, malformed syntax, and missing targets retain the complete native state and history. Separate undo steps, redo, pending-proposal preservation, tracking-mode refusal, same-message edit/read ordering, and saved reopen pass. Four real-session cases cover both client orders and reconnect receivers, peer text, remote selection/caret geometry, independent undo/redo, direct editing during peer deletion, deletion undo, ready states, and saved convergence. I02's eight browser regressions also pass, 80 assertions, `/tmp/codex-i03-read-regressions.log`. Full application typecheck passes 6621 files; final harness typecheck passes 2187 with its test-only session declarations corrected to actual public core exports. `bun run build` passes; the direct executor remains only in `dist/static/js/async/z2f.fc16dcbd37.js`, and read remains only in `dist/static/js/async/y7.86c52b3299.js`. Four initial scripts exclude both. Scoped formatting/lint and diff checks pass; the established legacy unawaited-submission warning and unrelated MapLibre build warning remain. Compatibility policy passes 94 decisions. Decision: `.collaboration/changes/sciobot-file-executor-direct.json`. No public API, shared format, provider, database, or runtime package changes occur. The evaluated provider stays test-only under the EigenPal Pro License. Suggestion and set routing remain I04-I05; persistence remains I07-I09.

## I04

- [ ] Connect explicit suggestion edit_document calls

**Depends:** [I03](10-app-integration.md#i03), [R19](05-suggestions.md#r19), [A06](06-direct-ai.md#a06), [A07](06-direct-ai.md#a07), [V01](11-verification.md#v01)

**Primary files:** `app:src/hooks/chat/docxSuggestionExecutor.ts`; `sibling docxSuggestionExecutor.test.ts`

**Read:** [S02](sources.md#s02), [S05](sources.md#s05), [S08](sources.md#s08)

**Change:** Map explicit is_suggesion=true to the suggestion pipeline. Preserve the public tool spelling and direct default.

**Close when:** The same patch yields pending review only with the explicit flag. A later wave can edit its proposed text.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I05

- [ ] Connect set_document with material-specific validation

**Depends:** [I03](10-app-integration.md#i03), [I04](10-app-integration.md#i04), [D15](04-edit-planner.md#d15), [D26](04-edit-planner.md#d26), [D27](04-edit-planner.md#d27), [D28](04-edit-planner.md#d28), [D29](04-edit-planner.md#d29), [D30](04-edit-planner.md#d30), [D31](04-edit-planner.md#d31), [D32](04-edit-planner.md#d32), [D33](04-edit-planner.md#d33), [D34](04-edit-planner.md#d34), [D35](04-edit-planner.md#d35), [D36](04-edit-planner.md#d36), [D37](04-edit-planner.md#d37), [D38](04-edit-planner.md#d38), [D39](04-edit-planner.md#d39), [P23](03-agent-projection.md#p23), [P11](03-agent-projection.md#p11), [V01](11-verification.md#v01)

**Primary files:** `app:src/hooks/chat/docxSetExecutor.ts`; `sibling docxSetExecutor.test.ts`

**Read:** [S02](sources.md#s02), [S05](sources.md#s05), [S14](sources.md#s14)

**Change:** Apply full targets through the planner in the requested mode. Enforce lesson structure only for lesson materials.

**Close when:** A generic document accepts ordinary Markdown. A malformed lesson set fails. The C01 complete rewrite preserves untouched DOCX parts.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I06

- [ ] Route tools across tabs and document identities

**Depends:** [I02](10-app-integration.md#i02), [I03](10-app-integration.md#i03), [I04](10-app-integration.md#i04), [I05](10-app-integration.md#i05), [A07](06-direct-ai.md#a07), [V01](11-verification.md#v01)

**Primary files:** `app:src/hooks/chat/useWorkspaceEditBridge.ts`; `app:src/lib/workspaceClientTools.ts`; `app:src/stores/workspaceStore.ts`; `focused routing tests`

**Read:** [S02](sources.md#s02), [S21](sources.md#s21)

**Change:** Register the native executor by workspace path. Retain other file executors and reject edits to read-only paths. Preserve parked calls for their original material and supersede unrun calls after a newer message. Reuse the application call-run ledger when mutations resume across remounts, reloads, and tabs. Replay stored output instead of applying an edit twice. Wait for the matching registered store editor to finish loading before mutating games, worksheets, activities, or print materials.

**Close when:** Switch tabs while an AI callback waits. Only its original live document can receive the edit; another path remains unchanged. Resume a parked call only for its original material. Repeated mutation IDs return prior output without another canonical edit or undo unit. A superseded unrun call reports no mutation. A store refetch cannot overwrite an edit already reported as applied; defer mutations until the correct editor is hydrated. Tool-driven focus records the assistant origin; manual focus retains the user origin. Identical provider call IDs in different conversations remain independent. Invalid variants on non-worksheet materials report validation errors without waiting for hydration.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** Pending. The R15 inventory gate detected an external reference bridge edit after its source tests. Reviewed `useWorkspaceEditBridge.ts` now imports the separate `workspaceClientTools.ts` helper. That helper stores call-run results and waiting turns outside document content. This required routing behavior has no native parity proof yet. Keep I06 open and reuse the application service when integrating it. A later external bridge edit also requires matching registered store editors to finish hydration before mutation. This prevents a refetch from replacing an acknowledged edit. The R16 audit also captures explicit assistant-origin focus calls. The clean reference workspace store already supports that origin. These routing changes remain pending and have no native parity claim. The R15/R16 engine suites do not verify these external changes.

## I07

- [ ] Define versioned document snapshot storage

**Depends:** [B02](02-react-blocks.md#b02), [R21](05-suggestions.md#r21)

**Primary files:** `app:src/lib/docx/snapshot.ts`; `sibling snapshot.test.ts`

**Read:** [S20](sources.md#s20), [S21](sources.md#s21), [S23](sources.md#s23)

**Change:** Define a versioned JSON envelope with immutable DOCX asset reference, document identity, revision, and derived Markdown. Include bytes in local drafts.

**Close when:** Decode a saved native snapshot and a legacy Slate snapshot distinctly. An unknown version preserves the raw data and reports a typed failure.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** unclaimed. **Evidence:** pending.

## I08

- [ ] Save native documents through existing autosave

**Depends:** [I01](10-app-integration.md#i01), [I07](10-app-integration.md#i07), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/autosave.ts`; `existing autosave adapter tests`

**Read:** [S20](sources.md#s20), [S21](sources.md#s21)

**Change:** Serialize one canonical revision and upload its DOCX before saving the snapshot reference. Reuse existing save status and retry behavior.

**Close when:** A failed upload never replaces the saved reference. Successful save restores identical text, payloads, and pending review after reload.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I09

- [ ] Restore drafts without overwriting newer content

**Depends:** [I08](10-app-integration.md#i08), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/restore.ts`; `sibling restore.test.ts`

**Read:** [S20](sources.md#s20), [S21](sources.md#s21)

**Change:** Choose saved snapshots and local drafts using existing revision rules. Guard delayed loads by material identity and live revision.

**Close when:** A newer local draft wins. A delayed server load cannot overwrite user typing or a newly selected document.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I10

- [ ] Import original DOCX bytes without a Plate conversion

**Depends:** [I01](10-app-integration.md#i01), [I07](10-app-integration.md#i07), [V01](11-verification.md#v01)

**Primary files:** `app:src/services/docxAttachmentImport.ts`; `attachmentWorkspaceApi integration test`

**Read:** [S20](sources.md#s20), [S21](sources.md#s21), [S25](sources.md#s25)

**Change:** Open the original uploaded DOCX and retain its asset reference. Derive agent Markdown separately without reconstructing imported OOXML.

**Close when:** Import, edit one sentence, and export a document containing headers, fields, drawings, and unknown XML. Untouched fingerprints remain equal.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I11

- [ ] Convert legacy paragraphs, marks, and links

**Depends:** [I07](10-app-integration.md#i07), [N02](07-native-content.md#n02), [N03](07-native-content.md#n03), [N08](07-native-content.md#n08), [N09](07-native-content.md#n09), [N10](07-native-content.md#n10), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/migrate-native-slate.ts`; `sibling migrate-native-slate.test.ts`

**Read:** [S01](sources.md#s01), [S20](sources.md#s20), [S22](sources.md#s22)

**Change:** Convert saved paragraphs, headings, marks, and links into native Word content. Retain the original snapshot for fallback.

**Close when:** Captured legacy fixtures preserve visible text, heading levels, links, and marks. Repeating migration does not duplicate content.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I12

- [ ] Convert legacy application block payloads

**Depends:** [I11](10-app-integration.md#i11), [B02](02-react-blocks.md#b02), [X09](08-domain-content.md#x09), [X19](08-domain-content.md#x19), [X23](08-domain-content.md#x23), [I22](10-app-integration.md#i22), [I23](10-app-integration.md#i23), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/migrate-domain-slate.ts`; `sibling migrate-domain-slate.test.ts`

**Read:** [S09](sources.md#s09), [S10](sources.md#s10), [S11](sources.md#s11), [S12](sources.md#s12), [S16](sources.md#s16)

**Change:** Convert registered domain nodes and opaque MDX ASTs to payloads and native fallbacks. Preserve assets and stable application identities.

**Close when:** One fixture per saved domain type retains every meaningful attribute. Unrecognized nodes retain raw payload and readable fallback.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I13

- [ ] Convert legacy pending text suggestions

**Depends:** [I12](10-app-integration.md#i12), [R03](05-suggestions.md#r03), [R14](05-suggestions.md#r14), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/migrate-slate-review.ts`; `sibling migrate-slate-review.test.ts`

**Read:** [S04](sources.md#s04), [S07](sources.md#s07), [S08](sources.md#s08), [S22](sources.md#s22)

**Change:** Convert pending text insertion, deletion, and replacement suggestions without resolving them. Preserve authors and replacement pairing.

**Close when:** Text committed and working views match before and after migration. Accept, reject, and undo produce equivalent text.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I14

- [ ] Export native DOCX through the workspace workflow

**Depends:** [I08](10-app-integration.md#i08), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/export-docx.ts`; `workspace export pipeline test`

**Read:** [S23](sources.md#s23), [S20](sources.md#s20)

**Change:** Capture one immutable revision and return native DOCX bytes through the existing export modal and archive flow.

**Close when:** Single and batch exports contain their selected revisions, pending native review, payload parts, and assets. Editing later cannot change an export snapshot.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I15

- [ ] Export Markdown and plain text from the correct view

**Depends:** [P02](03-agent-projection.md#p02), [P04](03-agent-projection.md#p04), [P05](03-agent-projection.md#p05), [R01](05-suggestions.md#r01), [R02](05-suggestions.md#r02), [I07](10-app-integration.md#i07)

**Primary files:** `app:src/lib/docx/export-text.ts`; `documentTextExport integration tests`

**Read:** [S23](sources.md#s23), [S11](sources.md#s11)

**Change:** Adapt Markdown and plain text exports with explicit committed or working view selection. Keep existing user-visible defaults.

**Close when:** Captured export fixtures match existing defaults. Hidden placeholders, AI anchors, and review preview text never enter exported content.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** unclaimed. **Evidence:** pending.

## I16

- [ ] Adapt native HTML and PDF export entry points

**Depends:** [I14](10-app-integration.md#i14), [I17](10-app-integration.md#i17), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/export-rendered.ts`; `existing export pipeline integration tests`

**Read:** [S23](sources.md#s23), [S22](sources.md#s22)

**Change:** Keep existing export modal actions and batch behavior. Feed immutable native snapshots to the selected print or rendering path.

**Close when:** An exported lesson includes custom block fallbacks, captions, page order, and complete listening text. No export path reconstructs imported DOCX through Markdown.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I17

- [ ] Render shared documents without live Plate state

**Depends:** [I07](10-app-integration.md#i07), [I12](10-app-integration.md#i12), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/shared-material/SharedLessonDocView.tsx`; `native static renderer and focused tests`

**Read:** [S22](sources.md#s22), [S23](sources.md#s23)

**Change:** Add static rendering for native snapshots using registered fallbacks. Keep legacy views readable, including saved comments and suggestions.

**Close when:** Shared native and legacy fixtures retain content and metadata. The native path does not mount an editable Plate provider.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I18

- [ ] Connect generated lesson draft hydration

**Depends:** [I09](10-app-integration.md#i09), [I12](10-app-integration.md#i12), [A08](06-direct-ai.md#a08), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/LinkedLessonEditor.tsx`; `native hydration helper and tests`

**Read:** [S21](sources.md#s21)

**Change:** Convert streamed generation drafts into native content. Apply final persisted content only when it does not replace newer local edits.

**Close when:** A final lesson hydrates once. Typing during generation survives delayed completion. Switching material cancels the old hydration.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I19

- [ ] Publish complete native snapshots to VFS consumers

**Depends:** [I07](10-app-integration.md#i07), [I14](10-app-integration.md#i14), [I15](10-app-integration.md#i15), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/vfs-snapshot.ts`; `VFS/export integration tests`

**Read:** [S05](sources.md#s05), [S20](sources.md#s20), [S23](sources.md#s23)

**Change:** Adapt complete persisted JSON and asset references for existing VFS readers. Keep read_document pagination independent from full-file export.

**Close when:** A paginated agent read still exports the complete document. A VFS snapshot resolves the exact immutable DOCX revision and required assets.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I20

- [ ] Add reversible workspace editor selection

**Depends:** [I09](10-app-integration.md#i09), [I11](10-app-integration.md#i11), [I12](10-app-integration.md#i12), [I13](10-app-integration.md#i13), [I24](10-app-integration.md#i24), [I25](10-app-integration.md#i25), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/editor-selection.ts`; `focused selection tests`

**Read:** [S20](sources.md#s20), [S21](sources.md#s21)

**Change:** Select native or legacy rendering by snapshot format and feature flag. Preserve legacy source snapshots during conversion failures.

**Close when:** Disabling the native route opens preserved legacy data. A failed conversion never saves an empty replacement or overwrites the source.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I21

- [ ] Isolate retained Plate consumers from the native editor

**Depends:** [I06](10-app-integration.md#i06), [I17](10-app-integration.md#i17), [I18](10-app-integration.md#i18), [I20](10-app-integration.md#i20)

**Primary files:** `app:workspace import boundaries and lazy component entry points`; `bundle evidence`

**Read:** [S01](sources.md#s01), [S02](sources.md#s02), [S21](sources.md#s21), [S22](sources.md#s22)

**Change:** Remove eager Plate imports from the native editor route. Inventory remaining static and unrelated consumers before dependency removal.

**Close when:** The native route bundle graph contains no Plate editor runtime. Every retained consumer has an explicit source and reason.

**Check:** Run scoped build and typecheck checks. Save the inspected dependency or bundle evidence.

**Claim:** unclaimed. **Evidence:** pending.

## I22

- [ ] Convert legacy list structure

**Depends:** [I11](10-app-integration.md#i11), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/migrate-slate-lists.ts`; `sibling migrate-slate-lists.test.ts`

**Read:** [S13](sources.md#s13), [S20](sources.md#s20)

**Change:** Convert saved list styles, starts, nesting, and checklist state to native numbering and list content.

**Close when:** Nested and restarted list fixtures retain order and check state. Conversion does not flatten lists into plain text.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I23

- [ ] Convert legacy tables and spans

**Depends:** [I11](10-app-integration.md#i11), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/migrate-slate-tables.ts`; `sibling migrate-slate-tables.test.ts`

**Read:** [S13](sources.md#s13), [S20](sources.md#s20)

**Change:** Convert saved rows, cells, widths, headers, borders, and merge spans using native table operations.

**Close when:** Merged-cell fixtures preserve text, spans, and row order. Invalid saved geometry retains the source and reports a conversion failure.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I24

- [ ] Convert legacy pending block suggestions

**Depends:** [I12](10-app-integration.md#i12), [I13](10-app-integration.md#i13), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/migrate-slate-block-review.ts`; `sibling migrate-slate-block-review.test.ts`

**Read:** [S07](sources.md#s07), [S08](sources.md#s08), [S22](sources.md#s22)

**Change:** Convert nested block insertion, deletion, and paired replacements while preserving group membership.

**Close when:** A pending activity replacement remains one group. Committed and working native views match the saved legacy views.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## I25

- [ ] Convert legacy pending property suggestions

**Depends:** [I12](10-app-integration.md#i12), [I13](10-app-integration.md#i13), [R03](05-suggestions.md#r03), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/migrate-slate-property-review.ts`; `sibling migrate-slate-property-review.test.ts`

**Read:** [S04](sources.md#s04), [S09](sources.md#s09)

**Change:** Convert old and proposed wrapper properties, including absent old values, to native payload proposals.

**Close when:** The existing suggestion-views cases preserve working and committed properties after migration. Reject removes a newly added property.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.
