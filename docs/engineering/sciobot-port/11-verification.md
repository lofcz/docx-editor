# Cross-feature evidence and cutover gates

These leaves collect required evidence. Passing one gate does not replace focused checks on each feature.

Read the [dispatch and completion rules](../../../SCIOBOT_PORT_TODO.md) before claiming a leaf.

Each named test or artifact is a deliverable. These specifications do not claim those checks already exist.

## V01

- [x] Generalize the two-editor collaboration fixture

**Depends:** None.

**Primary files:** `fork:e2e/support/sciobot-collaboration.ts`; `smoke fixture using examples/react-blocks/test/collaboration.test.ts`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Provide two real editor sessions, pause/reconnect, semantic digests, and presence assertions. Keep the evaluated transport in test-only dependencies.

**Close when:** The smoke case covers concurrent edits, undo/redo, deletion during peer editing, remote selections, reconnect, and save/reopen convergence.

**Check:** Run the two-real-session smoke scenario and save its semantic convergence assertions.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `collaboration-harness.test.ts` passes one real two-session scenario and 18 assertions: live synchronization, offline concurrent typing, peer-preserving undo/redo, remote selections, offline deletion during peer editing, reconnect, DOM repaint, and equal save/reopen digests. The reusable fixture owns test-only transport dependencies in `packages/sciobot/test/support`; e2e re-exports it. Package typecheck, scoped Prettier, and oxlint pass. Compatibility decision: tests only; no shared format or production runtime change.

## V02

- [x] Compare both AI modes against the seven Plate fixtures

**Depends:** [C01](01-contracts.md#c01), [D15](04-edit-planner.md#d15), [A01](06-direct-ai.md#a01), [R17](05-suggestions.md#r17), [R18](05-suggestions.md#r18), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/__tests__/ai-parity.test.ts`

**Read:** [S03](sources.md#s03), [S04](sources.md#s04)

**Change:** Run every C01 pair through direct edits and accepted suggestions. Compare independently captured normalized semantics, not only the two implementations.

**Close when:** All seven pairs match the Plate oracle. Reject restores the baseline. Direct undo/redo and review decision undo/redo restore exact states.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** The seven C01 pairs run through the native target planner in `target-c01-matrix.test.ts`, preserving independent non-body parts and exact empty replanning. The recomposition path then executes accepted suggestion waves, replaces earlier owned proposals, preserves accepted/rejected/manual content, and reopens saved documents without nested revisions. The focused batch passes 37 cases and 547 assertions (`/tmp/v02-c01-ai-batch.log`). The final unchanged C01/composer shared gate passes 333 cases and 16,071 assertions with two-session histories, peer edits, deletion, reconnect, remote geometry, decisions, undo/redo, and saved convergence (`docs/engineering/sciobot-port/evidence/composer-shared-history-repairs.md`). Direct execution is covered by A01’s seven captured pairs and its shared matrix. No fixture or oracle was recaptured.

## V03

- [x] Exercise mixed review waves through the application

**Depends:** [I06](10-app-integration.md#i06), [U09](09-editor-ui.md#u09), [U10](09-editor-ui.md#u10), [U11](09-editor-ui.md#u11), [U12](09-editor-ui.md#u12), [R20](05-suggestions.md#r20), [V01](11-verification.md#v01)

**Primary files:** `app:workspace native editor end-to-end tests and fixtures`

**Read:** [S02](sources.md#s02), [S07](sources.md#s07), [S08](sources.md#s08)

**Change:** Run two suggestion waves with a manual edit, accepted and rejected groups, and a property change. Include a direct edit between waves.

**Close when:** Agent reads show the working view. Decisions preserve user text. Reload restores pending groups with stable identities.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, app branch `claude/docx-native-cutover` commit `53d7c6f0b`. **Evidence:** Chromium `scripts/docx-native/harness/review-waves.test.ts` drives the production tool executor (`edit_document`), the review bar, and two real sessions (`file-executor` scenario, activity document): wave one proposes an activity name change and a text change; `read_document` returns the working view; Bob types in the tail; a direct edit (`is_suggesion` false) rewrites that text and keeps the pending groups; wave two changes the text from wave one and proposes a duration change. Each session's saved DOCX, read back without an editor, holds the same pending group identities. Alice accepts the first group and rejects the next in the review bar; both sessions converge with Bob's and the direct-edit text kept, and saved digests match before and after the decisions. 2 pass (both client orders); with `direct-edits.test.ts` 25 pass. Tab routing across documents (I06) is not part of this scenario.

## V04

- [x] Verify cross-feature collaboration convergence

**Depends:** [V02](11-verification.md#v02), [V03](11-verification.md#v03), [B10](02-react-blocks.md#b10), [B13](02-react-blocks.md#b13), [X05](08-domain-content.md#x05), [X10](08-domain-content.md#x10), [I08](10-app-integration.md#i08), [V01](11-verification.md#v01)

**Primary files:** `fork:e2e/sciobot-parity-collaboration.spec.ts`; `application integration fixture`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Exercise nested block moves, property proposals, text suggestions, and listening metadata through two real sessions.

**Close when:** Concurrent edits, peer deletion, undo/redo, presence, and reconnect converge. Both saved documents reopen with equal semantic digests.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, app branch `claude/docx-native-cutover` commit `ae9a6a390`. **Evidence:** Chromium `scripts/docx-native/harness/cross-feature.test.ts` runs two real sessions over one document with an activity, a listening exercise that carries generated-audio and opaque future fields, and a tail paragraph (`crossFeatureDocument`). Bob's remote caret shows in Alice's pane. Offline, Alice moves the activity below the listening block (`moveBlock`, as a block drag does) while Bob types inside the activity and deletes the tail paragraph, which is the move anchor; after reconnect both sessions hold the same Markdown with the moved order and Bob's text. One AI suggestion wave then proposes an activity property, activity text, and listening metadata (title); both sessions show the same pending group identities and working view; keyboard undo and redo in Alice's session converge in both; accepting all leaves equal documents, equal saved digests, and a reopened copy with the same working view. 2 pass (both client orders); with `listening-edits.test.ts` 10 pass. The block move is between body blocks; moves into or out of containers are covered by `block-drag.test.ts`.

## V05

- [x] Measure large-document interaction and retention

**Depends:** [B16](02-react-blocks.md#b16), [I01](10-app-integration.md#i01)

**Primary files:** `fork:e2e/sciobot-performance.spec.ts; `docs/engineering/sciobot-port/evidence/performance.json`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Measure open, typing, review application, page movement, and retained roots on a fixed 100-page fixture. Record environment and raw samples.

**Close when:** Publish p50/p95 timings and memory observations against the same native baseline. File explicit leaves for regressions and keep readiness red until resolved.

**Check:** Save the repeatable command, environment, fixture hash, raw samples, and regression decision.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** The repeatable native benchmark `scripts/sciobot-port/measure-v05.ts` generates a 100-page, 1,300-paragraph DOCX and measures open, typing, review application, page movement, retained roots, and process RSS. Two matched runs use the same fixture and revision. The baseline reports open p95 4,417 ms, typing p50/p95 117/283 ms, review p50/p95 98/130 ms, movement p50/p95 0.116/0.786 ms, and 1,300 retained paragraphs. RSS rises 423 MiB in the baseline and 430 MiB in the repeat. Raw samples, summaries, environment, comparison, revision, and script hash are stored in [performance.json](evidence/performance.json). No regression threshold is weakened; deployment hardware remains a release-readiness follow-up.

## V07

- [x] Check native runtime package boundaries

**Depends:** [C03](01-contracts.md#c03)

**Primary files:** `fork:private adapter dependency check; `app:production bundle report`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Inspect source imports and built runtime graphs. Exclude Pro, editor-api, and Plate editor runtime from the replacement path.

**Close when:** The graph contains one core copy and application-owned adapters. Test-only compatibility dependencies do not appear in the shipped route.

**Check:** Run scoped build and typecheck checks. Save the inspected dependency or bundle evidence.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** The private adapter manifest keeps `@docx-editor.dev/core` as a peer dependency and has no Pro, editor-api, Plate, or core runtime dependency. `packages/sciobot/src/__tests__/package-boundary.test.ts` passes its one boundary case. The private build bundles 331 adapter modules with native core external and passes `packages/sciobot/scripts/check-build.mjs` (`/tmp/v07-private-build.log`). Private typecheck passes 2,423 files (`/tmp/v07-private-types.log`). The generated graph contains document, React-block, and collaboration entries; the document entry excludes React and Yjs, and all entries exclude Pro, editor-api, and Plate runtimes. The test-only collaboration transport remains outside the shipped route. `check:lane-boundaries` still reports the repository’s existing store project DOM-library configuration errors; it is separate from the adapter boundary build and does not identify a shipped-route dependency. No published runtime dependency or shared format changes.

## V08

- [x] Close the production coverage ledger

**Depends:** Every other leaf in this backlog.

**Primary files:** `fork:docs/engineering/sciobot-port/coverage.md`; `evidence index`

**Read:** [S01](sources.md#s01), [S21](sources.md#s21), [S22](sources.md#s22), [S23](sources.md#s23)

**Change:** Match every production capability to completed leaves and reproducible evidence. Include migration and the reversible editor selection path.

**Close when:** No required coverage row remains pending or refused. Every failed or unverified requirement has an open leaf and prevents cutover.

**Check:** Validate every referenced file, task ID, and required coverage row.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** All 280 required leaves are closed after the user's desktop verification exclusions. All 50 coverage rows have completed required leaves. Application commit `3326720a2` installs runtime `7d5131205`; [native column controls](evidence/native-column-controls.md) records the final application checks. The production audit covers 37 workspace members, 26 static kits, 13 print overrides, and 202 source hashes. Its verification fields now follow each required leaf's checkbox. Before X20 closure, only ColumnKit and BaseColumnKit report pending verification. After closure, all 76 entries report completed required leaves. `SCIOBOT_REFERENCE_APP=<cutover worktree> node docs/engineering/sciobot-port/check.mjs` checks source hashes, links, dependencies, and closure evidence. The existing migration and reversible editor selection paths retain their completed leaf evidence. Desktop Word and operating-system input-method checks remain unverified exclusions. The recorded repository baseline failures remain separate from this scope closure.

## Scope exclusions

### V06

Desktop Microsoft Word round-trip verification is excluded from the required TODO list at the user's request on November 5, 2026. This check remains unverified. It does not block the coverage ledger.
