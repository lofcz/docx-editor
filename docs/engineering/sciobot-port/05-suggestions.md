# Suggestion state and review operations

Native revisions and application property proposals must form one review model. Required cases cannot close through refusal alone.

Read the [dispatch and completion rules](../../../SCIOBOT_PORT_TODO.md) before claiming a leaf.

Each named test or artifact is a deliverable. These specifications do not claim those checks already exist.

## R01

- [x] Read a pure committed document view

**Depends:** [C07](01-contracts.md#c07), [P01](03-agent-projection.md#p01)

**Primary files:** `fork:packages/sciobot/src/review/committed-view.ts`; `sibling committed-view.test.ts`

**Read:** [S04](sources.md#s04), [S07](sources.md#s07), [S25](sources.md#s25)

**Change:** Derive the view with pending suggestions rejected. Do not accept, reject, or mutate the live document during reads.

**Close when:** Insertions disappear, deletions return, and original formatting returns. Document revision and history remain unchanged.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `review/committed-view.test.ts` passes three real-editor cases and 25 assertions. Native Original projection hides insertions, restores deleted text, restores old run formatting and paragraph outline formatting, and preserves live paragraph IDs. Repeated reads leave canonical fingerprints, pending revision sites, package revisions, and history unchanged. One user undo still removes the preceding user edit. Stale, racing, and aborted reads fail without review decisions. The full private suite passes 95 tests and 621 assertions. Package typecheck, external-runtime build, scoped format, and oxlint pass. The reader uses Apache core export records and disposes its temporary session. R03 still owns application property proposals.

## R02

- [x] Read a pure working document view

**Depends:** [R01](05-suggestions.md#r01)

**Primary files:** `fork:packages/sciobot/src/review/working-view.ts`; `sibling working-view.test.ts`

**Read:** [S04](sources.md#s04), [S02](sources.md#s02), [S25](sources.md#s25)

**Change:** Derive the view with pending suggestions accepted. Preserve the underlying proposals and live review state.

**Close when:** Reads include pending replacements exactly once. The next agent patch can match pending text without resolving its review item.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `review/working-view.test.ts` passes two real-editor cases and 21 assertions. Concurrent committed/working reads return independent Original/Proposed text and formatting without changing live display mode or canonical fingerprints. Pending New text occurs once and matches a later replacement. Its Markdown range maps to the original live paragraph and model offsets. User undo/redo remains intact. Saved DOCX reopens with all native proposals pending and both views correct. The full private suite passes 97 tests and 642 assertions. Package typecheck, external-runtime build, scoped Prettier, and oxlint pass. Native view reads add no mutation. R03 retains application property projection work, and R04/R16 retain actual tracked patch execution.

## R03

- [x] Project pending application property changes

**Depends:** [B02](02-react-blocks.md#b02), [R01](05-suggestions.md#r01), [R02](05-suggestions.md#r02), [D13](04-edit-planner.md#d13)

**Primary files:** `fork:packages/sciobot/src/review/property-projection.ts`; `sibling property-projection.test.ts`

**Read:** [S04](sources.md#s04), [S09](sources.md#s09)

**Change:** Store old and proposed values with explicit property presence. Apply reject and accept semantics to pure views.

**Close when:** Committed activity duration is old; working duration is proposed. Rejection of an added property removes that property.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `property-projection.ts` owns bounded frozen copies of nonempty D13 attribute-only plans. Explicit old/proposed presence keeps null, zero, false, and absence distinct. Pure committed views select earliest old values; working views select latest proposed values. Ordered steps retain native container rules and original property ownership. Divergent current values, broken chains, duplicate IDs, conflicting owners, foreign documents, unsupported fields, structural changes, and mixed edit plans refuse without writes. Eighteen property cases include all eight source-executed application presence fixtures, independent blocks, native identity/opaque metadata retention, live proposed values, and asynchronous input snapshots. Native review reads combine domain attributes and Original/Proposed text with mapped native offsets. Save/reopen plus separately supplied serialized pending input retains the same read. Pending metadata persistence remains R12. Display mode, native proposals, package bytes, revision, and user history remain unchanged. Focused tests pass 99 cases and 1438 assertions across six files. Full private suite: 805 pass and the same two R29 failures, 807 tests, 70 files, 16817 assertions. Private typecheck passes 1628 files. Build passes with 96 private modules, one external core peer, and a React-free main. Scoped formatting/lint, root lint, compatibility, source coverage, backlog, and diff checks pass. Compatibility: `.collaboration/changes/sciobot-property-projection.json`, 43 decisions. No native/shared formats change. This pure private view requires no collaboration mutation scenario, published changeset, or public API extraction. All 196 application source hashes and revision `9693712d` remain unchanged. R12 retains stored metadata, fallback revisions, and live decisions. R14-R16 retain persistent groups and atomic group decisions. See [implementation notes](evidence/implementation-notes.md).

## R04

- [x] Execute tracked paragraph text replacements

**Depends:** [D02](04-edit-planner.md#d02), [R02](05-suggestions.md#r02), [V01](11-verification.md#v01), [P19](03-agent-projection.md#p19)

**Primary files:** `fork:packages/sciobot/src/review/track-text.ts`; `sibling track-text.test.ts`

**Read:** [S25](sources.md#s25), [S18](sources.md#s18)

**Change:** Apply text plans through existing native tracked operations. Use consistent author, time, and proposal identity.

**Close when:** The native review queue contains one replacement pair. Accept yields the target; reject yields the original; undo restores pending state.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `packages/sciobot/src/review/track-text.test.ts` and `track-text-collaboration.test.ts` pass 34 tests and 269 assertions through the package test preload. The application-owned module uses public Apache core reads. The executor validates current mount/revision, attribution, old text, native grapheme boundaries, and nonoverlapping body text plans before one guarded transaction. A contiguous replacement yields one native facade pair with distinct Word IDs and consistent author/date. Accept yields the target; reject yields the original; undo restores pending state; redo restores the decision. Native addresses survive save/reopen. Unicode, tabs, manual breaks, hyphens, unchanged run IDs/unknown XML, explicit AI attribution, owned insertion retraction, foreign pending insertions, multi-paragraph undo, no-ops, and atomic refusals pass. Two real sessions prove decisions, remote caret geometry, offline disjoint edits, reconnect, peer-preserving undo/redo, concurrent paragraph deletion, and saved semantic convergence. Full private suite: 262 tests, 3050 assertions, 28 files. Private typecheck, external-runtime build with 32 private modules, scoped Prettier, root lint, diff check, and `bun run collaboration:check --base HEAD` pass. Compatibility record: `.collaboration/changes/sciobot-tracked-text-execution.json`; no shared format changes. Read contract details in [implementation notes](evidence/implementation-notes.md). R14-R25 retain application groups, ownership, and later-wave work. The evaluated transport remains test-only under the EigenPal Pro License.

## R05

- [x] Execute tracked sibling block changes

**Depends:** [D05](04-edit-planner.md#d05), [R04](05-suggestions.md#r04), [V01](11-verification.md#v01), [R35](05-suggestions.md#r35), [R36](05-suggestions.md#r36), [P22](03-agent-projection.md#p22), [R37](05-suggestions.md#r37), [R38](05-suggestions.md#r38)

**Primary files:** `fork:packages/sciobot/src/review/track-blocks.ts`; `sibling track-blocks.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Track paragraph and section insertion or deletion without marking unchanged neighboring content.

**Close when:** C01 section insertion and deletion fixtures accept and reject exactly. Untouched sibling IDs remain stable.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof` (continuing Claude). **Evidence:** Native baseline `9770e78` plus current private worktree. Reviewed the existing guarded `trackSiblingPlan` after R38 closes the final required fragment prerequisite. C01 section insertion creates its heading and two body paragraphs; C01 deletion removes only the requested paragraph. Exact committed/working ASTs, accept/reject results, native revision receipts, creation/decision undo/redo, and pending/resolved save/reopen pass. Unchanged sibling objects and IDs remain stable. Additional cases cover first/middle/terminal insertion/deletion, replacing several paragraphs, sole-paragraph fallback, native links and marks, same-author neighboring proposals, atomic rollback, stale plans, and protected section-mark refusal. The 29 local and 26 real-session sibling cases pass, plus eleven R38 cases: `bun test --preload ./scripts/test/preload-dom.ts --timeout 60000 packages/sciobot/src/review/track-blocks-inline-code.test.ts packages/sciobot/src/review/track-blocks.test.ts packages/sciobot/src/review/track-blocks-collaboration.test.ts`; 66 pass, 1327 assertions, `/tmp/codex-r38-siblings-final.log`. Sixteen C01 session cases cover both client orders, reconnect receivers, accept/reject, independent peer text, history, remote geometry, and saved convergence. Eight middle-join cases retain unseen peer text during an offline decision. Two further cases retain peer text during terminal owner deletion. R38 adds code-specific geometry, peer edit/deletion, and history. Full private suite: 1355 pass, zero failures, 27635 assertions, 117 files, `/tmp/codex-r38-private-full.log`. Private typecheck/build, scoped format/lint, root lint, diff, and compatibility pass. Existing published APIs are unchanged and match their checked snapshots. Decision: `.collaboration/changes/sciobot-tracked-sibling-execution.json`. This closes paragraph and Markdown section execution; R13 retains atomic/resource-backed fragments, R41 retains rich rows/columns, and I03-I05 retain bridge integration. Native Word section-break edits retain their separate controls. No production EigenPal Pro License dependency or provider change.

## R35

- [x] Insert native sibling paragraphs without rewriting anchors

**Depends:** [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/store/tree-op-insert-paragraphs.ts`; `fork:packages/core/src/store/__tests__/insert-paragraphs.test.ts`; `fork:packages/sciobot/src/review/insert-paragraphs.test.ts`; `fork:packages/sciobot/src/review/insert-paragraphs-collaboration.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Add one bounded native sibling paragraph insertion operation. Preserve existing paragraph identities and properties. Track new content and paragraph marks through native revisions. Remove empty terminal proposed paragraphs when their own marks resolve away. Retain required paragraphs and peer content visible before the decision. R36 separately owns peer content that arrives after an offline rejection.

**Close when:** Direct and tracked insertions preserve untouched siblings. Accept retains inserted paragraphs. Reject restores original siblings exactly. Empty and terminal paragraphs, native history, protected ancestry, save/reopen, and two real sessions pass explicit assertions.

**Check:** Run focused native and real-session regressions. Check concurrency, reconnect, undo/redo, peer editing during deletion, remote geometry, and saved convergence. Run core/private typecheck, build, API extraction/check, scoped formatting/lint, compatibility, and backlog checks. Add the native changeset.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. The bounded native `insertParagraphs` operation clones fresh canonical/Word paragraph identities and preserves existing sibling objects. Tracked content and paragraph marks share one fresh native address. Terminal mark resolution removes its emptied paragraph while retaining required structure, section properties, opaque inline markers, and peer runs visible before the decision. The native insertion file supplies 27 explicit cases. The five-file native insertion/journal/reach/protection subset passes 225 tests and 980 assertions. The existing native revision subset passes 620 tests and 3154 assertions across 23 files. Three mode cases verify editing, suggesting attribution, viewing refusal, and one-unit history. Twelve real-session cases pass 652 assertions across both fixed client orders and reconnect receivers. They verify independent proposals, stable anchors/addresses, remote selection geometry, peer text, proposal/decision undo/redo, synchronized peer suffix decisions, explicit anchor deletion during insertion, and saved/reopened convergence. Full private suite: 879 pass, six fail, 885 tests, 77 files, 18526 assertions. Two failures remain R29; four unchanged expected-content regressions belong to R36. R05 requires R36 and remains open. Core/private typecheck pass 1457/1644 files. Core build, private build with 102 modules and one external core peer, scoped formatting/lint, root lint, API extraction/check, compatibility, source coverage, backlog, and diff checks pass. API checks scan 15 entries with no errors or drift; no forgotten-export baseline changes occur. Existing line limits remain intact. Changeset: `.changeset/native-sibling-paragraph-insertion.md`. Compatibility: `.collaboration/changes/native-sibling-paragraph-insertion.json`; no shared fields change. All 196 application hashes and revision `9693712d` remain unchanged. No Pro, application, or Plate implementation changes. See [implementation notes](evidence/implementation-notes.md#native-sibling-paragraph-insertion).

## R36

- [x] Retain offline peer text when rejecting an inserted terminal paragraph

**Depends:** [R35](05-suggestions.md#r35), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/package/retired-paragraph.ts`; `fork:packages/core/src/store/store/revision-paragraph-merge.ts`; `fork:packages/core/src/store/__tests__/retired-paragraph.test.ts`; `fork:packages/core/src/binding/__tests__/retired-paragraph-binding.test.ts`; `fork:packages/sciobot/src/review/terminal-paragraph-conflict.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Keep a peer's independent text visible when it arrives after another session rejects the paragraph's proposed insertion. Preserve explicit structural-owner deletion behavior. Use Apache core or application-owned collaboration policy; do not modify the evaluated transport implementation.

**Close when:** All four existing conflict cases retain `Peer` in both native views after reconnect. Decision undo restores `NewPeer` only in the working view. Redo retains `Peer`. Both fixed client orders and reconnect receivers converge and save/reopen identically. No assertion becomes weaker and no test becomes skipped.

**Check:** Run the preserved four-case regression, native paragraph resolution tests, ordinary structural deletion tests, and R35's passing session cases. Check history, geometry, stale decisions, and saved convergence. Record any shared-format compatibility decision and required changeset.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Actor-bound native mark resolution retains an empty terminal paragraph identity with versioned generic XML metadata. Story, layout, and binding projections hide it while empty; independent peer text makes the same identity visible. Native section marks, pending proposals, opaque inline content, and required paragraphs remain available. Word export removes recognized metadata from current properties and native formatting history without live mutation. Splits, joins, and fresh paragraph clones clear prior retirement state when creating new structure. `bun test --preload ./scripts/test/preload-dom.ts packages/core/src/store/__tests__/retired-paragraph.test.ts packages/core/src/binding/__tests__/retired-paragraph-binding.test.ts packages/sciobot/src/review/terminal-paragraph-conflict.test.ts` passes 32 cases and 290 assertions: 25 native/binding cases and seven real-session cases. The original four `Peer`/`NewPeer` expectations pass unchanged across both fixed client orders and reconnect receivers. Cases verify local caret clamping, remote geometry, stale decisions, save-before-peer-arrival, decision undo/redo, explicit owner deletion, ordinary editing after recovery, peer-preserving history, and saved/reopened convergence. The expanded native regression command passes 922 tests and 4477 assertions across 33 files. `bun run --filter @sciobot/docx test` reports 886 pass and the same two R29 failures, 888 tests, 77 files, 18690 assertions. No original content assertion becomes weaker or skipped. Core/private typecheck pass 1458/1645 files. Core/private build, scoped formatting/lint, root lint, API extraction/check, compatibility, source coverage, backlog, and diff checks pass. API checks scan 15 entries with no errors, drift, or forgotten-export baseline changes. Existing line limits and the private 102-module single-core boundary remain intact. Changeset: `.changeset/retired-paragraph-peer-recovery.md`. Compatibility: `.collaboration/changes/retired-paragraph-peer-recovery.json`; existing generic XML and canonical primitives carry the metadata without shared schema changes. All 196 application source hashes and revision `9693712d` remain unchanged. No Pro, application, or Plate implementation changes. R05 still owns application sibling execution and complete target parity. See the [format](evidence/retired-paragraph-format.md) and [implementation notes](evidence/implementation-notes.md#terminal-peer-recovery).

## R06

- [x] Execute tracked formatting and list changes

**Depends:** [D03](04-edit-planner.md#d03), [D04](04-edit-planner.md#d04), [D06](04-edit-planner.md#d06), [R04](05-suggestions.md#r04), [R30](05-suggestions.md#r30), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/track-format.ts`; `sibling track-format.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Apply native run, paragraph, and list property revisions. Retain prior values for rejection. Reuse R30 for checklist control changes.

**Close when:** Reject restores text marks, heading properties, and list membership. Accept matches the corresponding direct plan.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`, integrated runtime `40af5106`. **Evidence:** `trackFormatPlan` validates and copies native run, paragraph, and list plans before one owned-proposal transaction. Ordinary lists retain their previous numbering properties in native revisions. Checkbox conversions reuse R30 control snapshots. Controlled numbering starts and checkbox seeds receive fresh destination identities. Native code formatting uses D31 target planning. Production uses Apache core and the private adapter.

The six-file focused command passes **71 cases and 1240 assertions**. It covers seven marks, heading properties, existing/new numbering, starts, nesting, removal, checkbox conversion, malformed input, history, and saved decisions. Run rejection uses independently authored native run partitions; undo requires the original unsplit tree. Thirty-two real-session cases cover both client orders, both reconnect receivers, peer suffixes and deletion, author/reviewer histories, remote caret geometry, complete saved convergence, DOCX reopening, and room reopening. Four complete-package cases compare every modeled part, opaque byte, content type, relationship, and external target with an independently authored ledger. They pin numbering templates, nesting, start overrides, and retained review records. Only the validated random review identity varies.

The frozen integration run covers 231 private files: **2931 pass, 14 fail, zero skipped**. All 4124 recorded source hashes remain unchanged. Twelve failed table assertions expose a pre-existing nested paragraph defect. Native solo decisions now remove an unnecessary resolved empty cell/control sibling; shared decisions retain its identity for late peer text. The unchanged six-file native/private command passes 140 cases and 2176 assertions. Three peer/binding files pass another 20 cases and 834 assertions. The package-path failure comes from the isolated dependency symlink; the actual primary boundary passes one case and eight assertions. The remaining failure reaches the original 30000 ms capacity deadline. Complete primary reruns of all five failed review files pass **88 cases and 1963 assertions** without changing deadlines or expectations. All 4131 integration input hashes remain unchanged.

Private types pass 2066 integrated files; the test-only peer additions also pass types across 2068 files. Core types pass 1583 files. The primary private build verifies 218 modules and one external core; core build passes after its required i18n declarations. Root/scoped formatting, root lint, compatibility, and the 199-source-hash audit pass. Compatibility records zero new format fields. The first nested-fix formatter exits 143; its complete cached retry exits zero. The full integration run itself remains recorded as failing; the complete affected-file reruns resolve every reported failure. Commands, revisions, counts, logs, and source-manifest hashes are in [tracked formatting evidence](evidence/tracked-format.json). Application installation, controls, and later-wave orchestration retain their separate leaves.

## R07

- [x] Execute tracked table cell edits

**Depends:** [D07](04-edit-planner.md#d07), [R04](05-suggestions.md#r04), [V01](11-verification.md#v01), [R28](05-suggestions.md#r28), [R29](05-suggestions.md#r29)

**Primary files:** `fork:packages/sciobot/src/review/track-table-cells.ts`; `sibling track-table-cells.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Track text and mark edits within one cell. Keep native table boundaries and neighboring content unchanged.

**Close when:** The C01 changed cell accepts and rejects exactly. A replacement never joins text across cells.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Local tracked cell implementation and nine native regressions pass. The complete private suite reports 318 passing and two failing tests. Deterministic same-cell concurrency fails for fixed clients 73/41 in both reconnect delivery orders; R29 owns the missing behavior. R28 fixes the independent review-after-redo routing defect. Keep this leaf unchecked until all collaboration checks pass. See [implementation notes](evidence/implementation-notes.md).

**Closed by:** claude, branch `codex/react-block-proof`. R29 closed the remaining same-cell concurrency requirement. `packages/sciobot/src/review/track-table-cells.test.ts` passes 9 cases and 95 assertions. They cover C01 accept/reject with native neighbors, pending undo/redo, combined text/marks in one history entry matching direct native acceptance, mark-only receipts, pending save/reopen addresses, multi-paragraph cells and manual breaks that never join neighboring cells, editing/suggesting mode with AI attribution, and atomic refusal. The two-session scenario is `track-table-cells-collaboration.test.ts`, with 27 cases. It covers geometry, both fixed client orders and receivers, peer suffixes, review decisions, history, reopen, saved digests, and row deletion. The R29 evidence lists them, together with the complete private suite: 1039 pass, 0 fail.

## R08

- [x] Execute tracked table row changes

**Depends:** [D08](04-edit-planner.md#d08), [D16](04-edit-planner.md#d16), [R31](05-suggestions.md#r31), [R07](05-suggestions.md#r07), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/track-table-rows.ts`; `sibling track-table-rows.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Track inserted and deleted rows with native row revisions and matching history boundaries.

**Close when:** Accepting the added C01 row matches direct output. Rejection and undo restore the exact original row structure.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `trackTableRowPlan` (`packages/sciobot/src/review/track-table-rows.ts`) applies D08 `table` row operations (blank and valued insertion, deletion) and D16 `insert-row` intents as native Word row revisions. It validates context, mount revision, view mode, row ownership, control protection, duplicate anchors, and earlier pending row proposals. A complete immutable preview, including both review decisions, runs before one gated `applyAutomationOps` call with a separate undo unit. The after-allocation hook fills new cells through `inserted-row-content.ts`, which uses the same `insertParagraphs`/`deleteBlock` lowering as A01 `applyDirectPlan`. New content belongs to the row insertion revision. Receipts contain actual native row addresses per planned row. The imported-row requirement exposed a core defect: rejecting a tracked deletion left empty `w:trPr`/`w:tcPr` holders. Deletion-created holders now use mark-family IDs, and resolution removes holders that held only resolved markers (`tree-op-tables.ts`, `revision-marker-content.ts`). Source holders, including empty ones, stay. Reopened files and peer replicas keep an emptied holder, which converges. `track-table-rows.test.ts` passes 9 cases: C01 Three/Four accept equals the real A01 direct output for the same plan (serialized part, fresh `w14:paraId` excluded) and cell text equals the independent D16 oracle; reject restores the exact original fingerprint; decision and proposal undo/redo restore exact fingerprints with one proposal undo entry; deletion of an imported row without holders rejects exactly and accepts to the A01 output; imported holders survive rejection; valued and blank rows; a rich row (bold, link with tooltip, manual break, two paragraphs, heading, 500-twip height) accepts to A01 output and keeps its row address and content through pending save/reopen; mixed insertion/deletion plans form two decisions in one undo entry; code, nested tables, stale revisions, invalid context, duplicate anchors, non-row edits, pending rows, and a missing review module refuse atomically. `track-table-rows-collaboration.test.ts` passes 18 real-session cases: clients 41/73 and 73/41, both receivers, accept and reject. They cover an offline row proposal beside peer text, remote selection and caret geometry inside the new row, proposal and decision undo/redo, pending and final saved digests, and room reopen. They also cover tracked deletion while the peer edits that row, with decision history, and rejection while the proposer types inside the row. Core regression `tracked-row-deletion-holders.test.ts` passes 4 cases; 53 core revision/table/review files pass 1539 cases. The focused row/A01/D16 set passes 61 cases. Full private suite: 1089 pass; 6 load timeouts at the 5 s default in three other table collaboration files pass on rerun with `--timeout 60000`. Core typecheck/build and the scoped core API check pass with no snapshot drift (15 entries, 2443 warnings, 380 forgotten exports). Private build verifies 113 modules. Private typecheck reports no errors in these files; one error belongs to the concurrent D28 worker's `fragment-seeds.ts`. Scoped Prettier/oxlint pass. Changeset: `.changeset/tracked-row-deletion-holders.md` (core patch). Compatibility: `.collaboration/changes/tracked-row-deletion-holders.json` and `.collaboration/changes/sciobot-tracked-row-execution.json`; the policy passes 62 decisions. Limits: application blocks, tables, code, and inline atoms in new rows refuse here, as in A01. [R41](05-suggestions.md#r41) owns them after the D28-D30 materializer. The evaluated transport remains test-only under the EigenPal Pro License.

## R09

- [x] Execute tracked table column changes

**Depends:** [D09](04-edit-planner.md#d09), [D17](04-edit-planner.md#d17), [R07](05-suggestions.md#r07), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/track-table-columns.ts`; `sibling track-table-columns.test.ts`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Track column insertion and deletion, including grid and affected cell spans.

**Close when:** Accept and reject a middle-column edit with merged neighbors. Both outputs have valid grids and expected cell text.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Native column revisions did not exist, so this leaf adds them in core. `insertTableColumn`/`deleteTableColumn` accept `revision` (and `widthTwips` for a tracked insertion) through `tree-op-tracked-table-columns.ts`. An insertion runs the untracked native applier (and native fixed-grid sizing), then records `w:cellIns` on new cells, `w:tcPrChange` on every other cell, `w:tblGridChange`, and `w:tblPrChange` when the table width changes; existing Word decision semantics then reject to the exact original. Created holders use mark-family IDs; `heldOnlyResolvedMarks` also treats restored histories as resolved. A deletion records `w:cellDel` on the column's cells; accepting it follows Word width transfer and grid compaction. Deletions through a horizontal span and tables with pending structure/property history refuse (`unsupported`). `isTrackedEdit` reports the tracked forms so the review-module gate applies. Two native view gaps were closed: `table-cell-revision-view.ts` omits `w:cellDel` cells in Proposed (span to the previous survivor, else the next) and `w:cellIns` cells in Original, and `revision-formatting-projection.ts` now restores `tblPr`/`tblPrEx`/`tblGrid`/`trPr`/`tcPr` records in Original while keeping row/cell structure markers live. Private `trackTableColumnPlan` (`packages/sciobot/src/review/track-table-columns.ts`) executes D09 operations and D17 intents in one gated, separately undoable transaction, fills new cells through the shared `materializeNewCells` lowering, and returns per-table receipts (grid history has an author-less address). Shared executor gates moved to `tracked-structure.ts`; R08 uses them unchanged. `tracked-table-columns.test.ts` passes 10 core cases (regular and merged insertion both sides, holderless cells, sized insertion, last- and middle-column deletion, refusals). `table-cell-revision-view.test.ts` passes 2 layout cases. `track-table-columns.test.ts` passes 6 cases: a rich middle column (bold link, manual break, 1800-twip width) beside a horizontal span and a vertical merge accepts to the real A01 `applyDirectPlan` output with a valid grid and expected text, and rejects to the exact original fingerprint; Original view equals the pre-proposal view; deletion of the middle column between both merged neighbours hides it in Proposed, accepts to a valid three-column grid with expected text, and rejects exactly; one proposal undo entry and decision undo/redo; pending save/reopen keeps all addresses and rejects exactly; refusals are atomic. `track-table-columns-collaboration.test.ts` passes 16 real-session cases: clients 41/73 and 73/41, both receivers, accept and reject, with peer text in a surviving cell and a deleted cell, remote caret geometry in a new cell, proposal and decision undo/redo, saved digests, and room reopen. Core store/editor regressions: 107 files, 2497 pass. Layout revision/table regressions: 138 files, 1848 pass; `revision-memo-retention` passes alone (GC timing in the batch) and the corpus attribution test exceeds its fixed 30 s timeout at load average 36 both with and without these layout changes. Full private suite: 1148 pass; one 5 s load timeout in `table-format-collaboration.test.ts` passes on rerun. Core typecheck/build, private typecheck/build (121 modules), scoped Prettier/oxlint, and compatibility (70 decisions) pass. API snapshot updated with `bun run api:extract` (store `TreeDocOp` column fields; other drift in the snapshot belongs to concurrent workers). Changeset: `.changeset/tracked-table-columns.md` (core minor). Decisions: `.collaboration/changes/tracked-table-columns.json`, `.collaboration/changes/sciobot-tracked-column-execution.json`. Docs: feature matrix `review.cell-revisions` and the tracked-changes guide. Limits: deleting through a horizontally merged cell refuses; per-card partial decisions follow Word semantics; rich blocks in new cells belong to [R41](05-suggestions.md#r41). The evaluated transport remains test-only under the EigenPal Pro License.

## R10

- [x] Execute tracked table merge and split changes

**Depends:** [D10](04-edit-planner.md#d10), [R07](05-suggestions.md#r07), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/track-table-merges.ts`; `sibling track-table-merges.test.ts`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Represent merge and split proposals with complete reversible span and cell changes.

**Close when:** Accept matches direct merged output. Reject restores original cells, content, and geometry. No partial merge survives failure.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Native merges and splits had no tracked form, so core adds `revision` to `mergeTableCells`/`splitTableCell` (`tree-op-tracked-table-merges.ts`). A split reuses the R09 decorator, extended to new rows: new rows get a row insertion, new cells `w:cellIns`, survivors `w:tcPrChange`, plus grid/table histories. A merge keeps the original cells: cells right of each row's first cell get `w:cellDel`, a multi-row merge marks the first cells with `w:cellMerge` (`rest` on top, `cont` below), and the content the direct merge gathers is inserted into the owner as a tracked paragraph insertion (empty owner paragraphs the merge drops get a deleted mark). The landing is read from a scratch direct merge run under `runWithoutJournalCapture`, so peers receive only the tracked markup; the collaboration test caught the earlier journal leak. Gathered tables, controls, or non-empty dropped paragraphs refuse (`unsupported`). Suggesting mode now admits attributed merges/splits. Resolved views read pending `w:cellMerge` (`resolvedVMerge` in `table-cell-revision-view.ts`). Private `trackTableMergePlan` executes D10 plans through the shared `tracked-structure.ts` gates. `tracked-table-merges.test.ts` passes 10 core cases: horizontal (including an empty owner), vertical, rectangle, three-row merges, and two-column, three-row, and 2x2 splits all accept to the direct command (fresh `w14:paraId` and the explicit `continue` spelling excluded) and reject to the exact original fingerprint; an existing-grid split of a merged cell restores the merge on reject; stacked histories, gathered nested tables, and invalid attribution refuse. `track-table-merges.test.ts` passes 11 cases: rectangle, vertical, and 2x2 split accept to the real A01 `applyDirectPlan` output and reject exactly, with Original equal to the pre-proposal view and Proposed showing the merge; one proposal undo entry and decision undo/redo; a merge covering every row accepts to equal cell widths, merges, and content with Word grid compaction (documented limit); pending save/reopen keeps addresses and rejects exactly; editing and suggesting modes keep their mode and AI attribution; refusals are atomic. `track-table-merges-collaboration.test.ts` passes 16 real-session cases: clients 41/73 and 73/41, both receivers, accept and reject, with peer text in the owner and outside, remote caret geometry in the owner, proposal and decision undo/redo, peer text in a merged source cell (kept on reject, removed with the deleted cell on accept), and an offline split beside peer text, with saved digests. Core store/editor regressions: 113 files, 2527 pass. Full private suite: 1192 pass, 0 fail. Core/private typecheck, core build, scoped Prettier/oxlint, and compatibility (74 decisions) pass; the API snapshot was re-extracted. Changeset: `.changeset/tracked-table-merges.md` (core minor). Decisions: `.collaboration/changes/tracked-table-merges.json`, `.collaboration/changes/sciobot-tracked-merge-execution.json`. Docs: feature matrix and tracked-changes guide. The evaluated transport remains test-only under the EigenPal Pro License.

## R11

- [x] Execute tracked table property changes

**Depends:** [D11](04-edit-planner.md#d11), [R07](05-suggestions.md#r07), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/track-table-format.ts`; `sibling track-table-format.test.ts`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Track table and cell formatting independently from text changes. Use native property revisions when available.

**Close when:** Reject restores widths, backgrounds, borders, and header flags. Accept retains unrelated cell text.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Core gains the write half of Word table property history: the seven table property commands accept optional `revision` (`tree-op-tracked-table-properties.ts`). The untracked command runs first; each changed container keeps its new values and records the original in `w:tblPrChange`/`w:trPrChange`/`w:tcPrChange`/`w:tblGridChange`. Because rejecting any row record resets that row's unrecorded cells to defaults, a changed row records every cell and a grid change every row. A later proposal keeps the original record; landing back on it drops this author's records and keeps another author's. Holders the proposal creates take mark-family ids and `heldOnlyResolvedMarks` now removes them when a restored record empties them, so rejection is exact. `trackTableFormatPlan` (private) lowers D11 plans in one gated, separately undoable transaction, refuses when formatting tracking is disabled, keeps unrelated pending revisions, and returns per-table Word addresses. Core: 16 cases (fill, cleared fill, fill on a cell without properties, borders, vertical alignment, header rows, column widths, divider drag, right edge, row height on a row without properties; exact reject, direct-equal accept, save/reopen both decisions, original-record retention, foreign-record retention, revert removing created holders, atomic refusals). Private: 23 native cases (each target accepts to the A01 direct output and rejects to the exact original; cell text unchanged in both views; one-step undo/redo; combined fill+width save/reopen with stable addresses; later proposal; refusals; mode preservation) and 10 two-session cases (both client orders and reconnect receivers, accept/reject, peer text in the reformatted cell, remote caret geometry, proposal and decision undo/redo, row deletion during the offline proposal, saved digests). Reopened files keep an emptied created holder as `<w:tcPr/>`/`<w:trPr/>`, as R08 documents. Full private suite: 1308 pass, 0 fail, 115 files. Core review/revision/tracked regressions: 1955 pass, one load timeout (corpus attribution, 49 s at load ~35) across 105 files; hyperlink/table/tracked subset 188 pass. Core and private typecheck pass in these files; build/API extraction not rerun to avoid deleting the shared `dist` under concurrent workers, so the core API snapshot refresh for the new optional `revision` fields remains for the next build. Changeset `.changeset/tracked-table-properties.md`; decisions `.collaboration/changes/tracked-table-properties.json`, `.collaboration/changes/sciobot-tracked-table-format.json`; compatibility policy passes 83 decisions. Docs: tracked-changes guide and table feature row. See [implementation notes](evidence/implementation-notes.md).

## R12

- [x] Execute tracked application property updates

**Depends:** [R03](05-suggestions.md#r03), [D13](04-edit-planner.md#d13), [R04](05-suggestions.md#r04), [V01](11-verification.md#v01), [R32](05-suggestions.md#r32), [R34](05-suggestions.md#r34)

**Primary files:** `fork:packages/sciobot/src/review/track-properties.ts`; `sibling track-properties.test.ts`

**Read:** [S04](sources.md#s04), [S09](sources.md#s09)

**Change:** Write pending property metadata and its native Word fallback revisions in one transaction.

**Close when:** Activity name and duration appear as property changes. One accept or reject updates both metadata and fallback without replacing the phase.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `trackPropertyPlan` executes D13 deltas and exact application-supplied generated labels through one guarded native transaction. R34 records actual allocated Word revision addresses before publication. Canonical attributes retain committed values while R03 projects pending values. `decidePropertyProposal` validates independent host identity, current revision, exact old-value presence, control ownership, and retained native receipts. Acceptance commits selected attributes and fallback text; rejection preserves original attributes and restores fallback text. One undo restores the pending record and native revisions. Twenty-nine live cases cover all eight captured application presence fixtures, metadata-only updates, Unicode/tabs/breaks/empty labels, native save/reopen, phase/body identity, opaque metadata, control locks, stale inputs, and atomic refusals. Twelve real-session cases cover both fixed client orders and reconnect receivers, independent offline proposals, peer body text, proposal/decision history, remote geometry, owner deletion during peer proposals, peer suffixes inside the fallback, and saved convergence. Focused tests pass 101 cases and 1323 assertions across five files. Full private suite: 864 pass and the same two R29 failures, 866 tests, 74 files, 17843 assertions. Private typecheck passes 1640 files; build verifies 102 private modules, one external core peer, and a React-free main. Core tests/typecheck/build, scoped API extraction/check, formatting/lint, root lint, compatibility, source coverage, backlog, and diff checks pass. Compatibility: `.collaboration/changes/sciobot-tracked-property-execution.json`; 47 decisions, with the earlier native format-field change unchanged. This private executor needs no published changeset or public API extraction. All 196 application source hashes and revision `9693712d` remain unchanged. R14/R19 retain group composition and later-wave replacement for overlapping pending blocks. R21 retains review UI restoration. R22 retains external Word reconciliation; missing or externally resolved receipts refuse instead of silently deciding metadata. No Pro implementation changes. See [implementation notes](evidence/implementation-notes.md).

## R13

- [x] Execute tracked atomic block replacements

**Depends:** [D14](04-edit-planner.md#d14), [R05](05-suggestions.md#r05), [B02](02-react-blocks.md#b02), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/track-atomic.ts`; `sibling track-atomic.test.ts`

**Read:** [S09](sources.md#s09), [S10](sources.md#s10), [S12](sources.md#s12)

**Change:** Track complete atomic block replacement with payload and asset references. Keep the old payload until review resolves it.

**Close when:** Accept selects the proposed visual or listening block. Reject restores the original asset references and fallback.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, current private working tree. `trackAtomicPlan` retains complete old/proposed payload pairs, opaque metadata, asset references, attribution, and exact native fallback revision receipts in one guarded transaction. The original payload remains canonical until acceptance. The shared bounded review projection selects complete old/proposed payloads and rebinds stable block IDs after reopening. `decideAtomicProposal` accepts the proposed visual/listening payload or rejects it while restoring the original fallback. Each proposal and decision has one undo unit. Existing property execution retains its own route and cannot interpret atomic pairs as attribute writes. Twenty-five native cases cover combined visual/listening decisions, all changed captured listening asset fixtures, unhydrated and hydrated audio, script invalidation, opaque MDX, no-ops, unchanged text fallbacks, inserted widgets with generated labels and empty caret paragraphs, save/reopen, stable owners, exact rollback, locks, view mode, missing review support, stale identities/values, changed payloads, incomplete receipts, and wrong executor routing. Fourteen real-session cases cover both client orders, reconnect receivers, independent offline visual/audio proposals, native text, proposal/decision history, remote fallback geometry, owner deletion during a peer proposal, peer suffixes, conflicting atomic ownership without data loss, and saved convergence. `bun test --preload ./scripts/test/preload-dom.ts --timeout 60000` with the two atomic files and five property regression files passes 116 cases and 2041 assertions; `/tmp/codex-r13-focused-final.log`. Full private suite passes 1392 cases, zero failures, 28274 assertions, 119 files; `/tmp/codex-r13-private-full.log`. The later focused run adds two concurrent-ownership cases and fallback geometry assertions. Private typecheck passes 1738 files; build verifies 133 private modules and one external core peer. Scoped format/lint, root lint, compatibility, and source coverage pass. Root lint retains existing unrelated warnings. Decision: `.collaboration/changes/sciobot-tracked-atomic-execution.json`; 95 decisions, with the earlier native format-field change unchanged. No published API, provider, transport, or shared-format implementation changes occur. This private executor needs no published changeset. Native text fallbacks use one paragraph; generated labels can retain empty caret paragraphs. Inline chips, richer imported fallbacks, production media renderers/services, group composition, later-wave replacement, and external Word decisions retain their dedicated work. R14/R19 own overlapping group decisions; R22 owns external Word reconciliation. All 196 audited application source hashes and reference revision `071352581` remain unchanged. The app pins runtime `a28acec2b32387d9508ee3e89bca7675377df8b7` through `refs/backup/codex-r13-runtime` and private archive `sciobot-docx-7c31f5cd308e.tgz`. Four published archives remain byte-identical. Installed-bundle atomic accept/reject, pending `editor.save()` reopening, assets, and history pass. Application regressions pass 53 unit cases and 16 browser cases with 169 assertions. Application/harness typechecks pass 6622/2188 files; production build passes with the existing MapLibre warning. Archive hashes, manifest, lock, CRLF-aware formatting, and diff checks pass. I04 retains native suggestion routing. See [implementation notes](evidence/implementation-notes.md).

## R14

- [x] Persist replacement and review group identities

**Depends:** [R04](05-suggestions.md#r04), [R05](05-suggestions.md#r05), [R12](05-suggestions.md#r12), [R13](05-suggestions.md#r13), [R27](05-suggestions.md#r27), [R45](05-suggestions.md#r45), [V01](11-verification.md#v01), [R44](05-suggestions.md#r44)

**Primary files:** `fork:packages/sciobot/src/review/groups.ts`; `sibling groups.test.ts`

**Read:** [S07](sources.md#s07), [S18](sources.md#s18)

**Change:** Pair old and proposed content into document-ordered groups at any nesting depth. Persist stable group identifiers.

**Close when:** A replacement has one group after save/reopen. Recomputing groups preserves unrelated group IDs and does not mutate content.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, private working tree. Native-only groups use immutable independent custom XML records with compact stable IDs and exact native part/revision membership. Property/atomic groups use proposal and fallback block identity. Canonical node IDs, paths, and offsets remain disposable read projections. New Word revisions carry no application group attributes or namespaces. Original native XML and compatibility rules stay intact. R44 initializes the carrier before room creation. R45 reserves active, inactive, and removed native IDs before allocation, preserving records for either actor's decision undo. A new neighbor cannot merge earlier group identities. Exact structural and hyperlink receipts form complete groups. Incomplete membership stays visible and read-only; conflicts and future versions refuse. Imported reused addresses retain separate visible read-only groups for R21/R22 reconciliation. Twenty local and seventeen real-session group cases pass. Four allocation cases and ten additional real-session cases prove same-paragraph later proposals with the same author/date, both client orders and reconnect directions, native input, peer-preserving cross-actor undo/redo, deletion recovery, remote geometry, room/DOCX reopening, and separate rapid proposal undo units. The strict original table XML save/reopen comparison passes unchanged, as do all 23 table-format cases. Focused group/allocation checks pass 66 cases and 1051 assertions across five files, including fifteen native allocator cases; `/tmp/codex-r45-focused-final.log`. Full private suite: **1460 pass, zero failures**, 29728 assertions, 125 files; `/tmp/codex-r45-private-full.log`. Native regressions pass 124 cases with 663 assertions; module/isolation checks pass 24 with 65 assertions, and lane guards pass 36 with 145 assertions. Private/core typechecks pass 1754/1471 files. The private build verifies 140 modules, one external core, and a React-free main. Core build/declarations, scoped formatting, root lint, API extraction/check, compatibility, source coverage, and backlog pass. Existing unrelated lint and declaration warnings remain. Public native reservation infrastructure uses `.changeset/retained-native-revision-ids.md` and its separate compatibility decision. The private grouping code adds no public document-model method or evaluated implementation. Compatibility passes 98 decisions with the earlier one-field format change unchanged. All 196 audited source hashes remain identical. The reference HEAD advanced to `89d7a134`; independent hashing and coverage comparison prove only the audit revision changed. The app now pins N21 runtime `f8a2d15e1a686fd504d8ae6255cab3d5dfca68a1`, with core archive `docx-editor.dev-core-0247acf90551.tgz` and adapter `sciobot-docx-fdf6552c585c.tgz`. Installed-bundle group/history/reopen checks, 53 app unit cases, 16 browser cases, both app typechecks, and production build pass. See N21 for the initialization regression and artifact provenance. Group decisions, later waves, external Word reconciliation, and suggestion bridge routing retain R15/R16/R19/R22/I04. See [implementation notes](evidence/implementation-notes.md).

## R15

- [x] Accept one review group atomically

**Depends:** [R14](05-suggestions.md#r14), [R46](05-suggestions.md#r46), [R47](05-suggestions.md#r47), [R48](05-suggestions.md#r48), [R49](05-suggestions.md#r49), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/accept-group.ts`; `sibling accept-group.test.ts`

**Read:** [S07](sources.md#s07), [S08](sources.md#s08)

**Change:** Resolve one complete group, including native revisions and application properties, in a fresh history transaction.

**Close when:** Accepting group two retains groups one and three. One undo returns exactly group two to pending state.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** R14/R46/R47/R48/R49/R50/V01 are complete. Acceptance resolves the exact selected native sites together per part, including coupled merge/split members. Native/property/atomic groups retain immutable data, neighboring pending groups, fresh revision checks, one native review transaction, and one undo unit. Sixteen permanent real-session merge/split cases preserve original Word views, peer text, geometry, both actors' undo, complete redo, and saved reopening. The complete private command passes **1642 cases, zero failures, 46828 assertions, 142 files**; `/tmp/codex-r15-private-final.log`. Installed native structural/capacity checks pass **32 cases, 1122 assertions**; `/tmp/codex-r15-final-installed-native-verified.log`. Installed baseline, solo, and actual shared main/collaboration checks pass stable IDs, immutable labels, disposed scope refusal, complete native history, and both capacity limits. The immutable runtime is `refs/backup/codex-r15-runtime-final`, commit `a96fbb7cf5b5e283e5cd65c8fc00f8745f77865d`, tree `0d211925a3fb42270df19a3da8e6edf8a7230c84`. Its private archive is `sciobot-docx-4f7af82bd49f.tgz`. All four published archives retain exact N21 bytes. The reuse guard passes unchanged build inputs and rejects changed published inputs before mutation. All five current archive hashes, provenance, dependency/override/lock references, and line endings pass. App and adapter resolve one physical native core and one Yjs. Seven app unit files pass **53 cases**. Chromium passes **16 cases, 169 assertions** against the controlled new-snapshot harness. Frontend and harness typechecks pass **6636/2202 files**. Production build and source/app formatting pass. Root lint and 105 compatibility decisions pass with baseline warnings. The app backup is `refs/backup/codex-docx-app-r15-final`, commit `71e98c789d22215b0188b42b243176359734904d`, tree `a99de2d471f25309183d141f6bf6d3da1def1bd7`. Its exact five changes are package.json, bun.lock, SOURCE, replaced adapter archive, and new adapter archive. All app source files, branch HEADs, and real indices stay unchanged. All 197 reviewed reference source hashes match; externally added call-run behavior remains pending under I06. The controlled server stops while the older harness stays alive. See the implementation notes for exact commands and logs. Test transport uses the EigenPal Pro License through test-only fixtures. No native core, evaluated implementation, published API, or protocol changes accompany this leaf.

## R16

- [x] Reject one review group atomically

**Depends:** [R14](05-suggestions.md#r14), [R15](05-suggestions.md#r15), [R46](05-suggestions.md#r46), [R47](05-suggestions.md#r47), [R48](05-suggestions.md#r48), [R49](05-suggestions.md#r49), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/reject-group.ts`; `fork:packages/sciobot/src/review/decide-group.ts`; `sibling reject-group.test.ts`

**Read:** [S07](sources.md#s07), [S08](sources.md#s08)

**Change:** Reject one complete group without dropping unaffected suggestions. Restore absent properties correctly.

**Close when:** Rejecting group two restores its old content only. Undo returns the same pending group and original group identity.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** R14/R15/R46/R47/R48/R49/V01 are complete. Rejection shares the R15 decision transaction, bounded inputs, fresh revision check, native gate, exact selected sites, complete payload previews, immutable receipts, scoped history, and capacity preflight. Native/property/atomic rejection restores original content and property presence. It retains neighboring groups, immutable labels/assets, retired IDs, and one undo unit. Both actors' decisions, deletion recovery, body/header groups, imported identities, peer text, geometry, Word export, and room/DOCX reopening pass. R36 retains its native empty paragraph identities and existing pure Word exporter. No revision filter or repair changes. The complete current private suite passes **1713 cases, zero failures, zero skips, 147 files** through the official process-per-file runner scoped to private sources, with one worker and unchanged deadlines; `/tmp/codex-r16-private-processes.log`. The earlier serial command had six timeouts across two existing files. Both entire files pass unchanged deadlines on rerun: 27 cases and 839 assertions. That failed command remains documented separately. Private typecheck passes 1793 files. Build verifies 153 private modules with external core and a React-free main. Source formatting and root lint pass with baseline warnings. Installed main/collaboration entries pass acceptance, rejection, disposed-scope refusal, both actors' history, saved convergence, and both capacity limits. Seven installed regression files pass **95 cases, 4803 assertions**; `/tmp/codex-r16-installed-reject-matrix-verified.log`. Runtime `refs/backup/codex-r16-runtime` names commit `394169a11bab4e7b8f1d5b50b03cbd015764c18c`, tree `1f4ac85286bcfe583e8652bdc56e84bccb60dc87`. Its private archive is `sciobot-docx-e4825a213bb2.tgz`. All four published archives retain exact N21/R15 bytes. Reuse guards, five hashes, provenance, dependency/override/lock references, installed file bytes, and line endings pass. App and adapter resolve one physical core and one Yjs. Seven app unit files pass **53 cases**. Chromium passes **16 cases, 169 assertions**. Frontend/harness typechecks pass **6638/2204 files**. Production build passes with the established MapLibre warning. The controlled R16 server stops after testing. App backup `refs/backup/codex-docx-app-r16-final` names commit `65afe7d196f13d3cd6f53b2a82dbb13da9378bb6`, tree `7c148ee93188d19dfd3033d138b535f41fe24cb7`. Exactly five artifact paths differ from R15. All app sources, branch HEADs, and real indices stay unchanged. All 198 reviewed reference hashes match; external call-run and focus-origin behavior remains pending under I06. Test transport uses the EigenPal Pro License through test-only fixtures. No native core, evaluated implementation, published document-model API, or protocol changes accompany this leaf. See [implementation notes](evidence/implementation-notes.md) for commands, setup corrections, timeout evidence, and immutable backups.

## R17

- [x] Accept all review groups in one user action

**Depends:** [R15](05-suggestions.md#r15), [R16](05-suggestions.md#r16), [R51](05-suggestions.md#r51), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/accept-all.ts`; `sibling accept-all.test.ts`

**Read:** [S08](sources.md#s08)

**Change:** Batch bulk acceptance into one history entry. Guard against repeated invocation while application runs.

**Close when:** One undo restores every pending group. A double invocation does not create duplicate history or resolve an unrelated document.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** R15/R16/R51/V01 are complete. `acceptAllReviewGroups` selects every complete current group and resolves its exact combined native sites per Word part. Pure previews validate all property/atomic members and immutable receipts before one native transaction. Each transient stage contributes to the unchanged R50 removal budget and R51 field budget. One undo restores all pending group IDs, native markup, and payloads. Empty/repeated calls add no history. The per-editor guard refuses synchronous reentry, including one-group calls. Fresh revision, document ID, epoch, and surface checks refuse calls for a replaced document. A late read-only or locked member refuses the whole batch before writes. Buffered input, view mode, and missing review capability retain their guards.

R51 fixes the strict concurrent geometry prerequisite without changing the R17 oracles. The final stable source suite passes **1874 cases, zero failures, zero skipped, 156 files**: `node /tmp/codex-r51-private-runner.mjs --jobs 2`; `/tmp/codex-r51-stable-full-private.log`. Actual installed bundles pass the 12-case structural bulk matrix, 12-case exact/insufficient-capacity matrix, and 12-case payload/concurrent-group matrix. Both client orders, reconnect receivers, actor histories, unseen peer text, remote geometry, room snapshots, and DOCX reopening pass. Native comparisons retain independent exact Word fingerprints. Installed local checks additionally pass **10 cases, 85 assertions**: `bun test --tsconfig-override /tmp/codex-r51-installed-tsconfig.json --timeout 30000 --preload /tmp/codex-r51-installed-preload.ts ./src/review/accept-all.test.ts`; `/tmp/codex-r17-installed-local.log`. They explicitly prove reentry refusal, duplicate-call history, complete undo, and equal-revision document switches. The other installed logs are `/tmp/codex-r51-installed-native-full.log` and `/tmp/codex-r51-installed-bulk-payload.log`.

The verified runtime stays `refs/backup/codex-r51-runtime`, commit `415bf8cf0cb86c4cd0385cc2210802e913bbebe6`. It already contains this reviewed implementation. R51 source types/build/lint/format/API checks remain applicable because this final review changes only evidence and the R17 compatibility reason. No new shared format or published API changes occur. Compatibility decision: `.collaboration/changes/sciobot-accept-all-review-groups.json`. R51 separately records its native core changeset and private protocol upgrade. App artifact backup `refs/backup/codex-docx-app-r51-final` retains one physical core/Yjs, all five verified hashes, frozen install, 53 app units, 16 browser cases, both typechecks, and production build. Its source tree remains unchanged. All 198 reference source hashes, both branch HEADs, and real indexes remain unchanged. Test-only transport uses the EigenPal Pro License. Host admission, review UI wiring, later-wave state, and cutover gates remain separate. See [implementation notes](evidence/implementation-notes.md).

## R18

- [x] Reject all review groups in one user action

**Depends:** [R16](05-suggestions.md#r16), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/reject-all.ts`; `sibling reject-all.test.ts`

**Read:** [S08](sources.md#s08)

**Change:** Reject against the live committed view. Preserve intervening user edits and batch the action into one history entry.

**Close when:** A manual sentence added between AI waves survives reject-all. One undo restores all pending groups.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** R16/V01 are complete. R17/R51 supply verified atomic bulk dispatch and owned history. `rejectAllReviewGroups` resolves every complete current native/property/atomic group against the live canonical package in one native history frame. It reuses exact native membership, immutable decisions, fresh document/epoch/revision checks, and complete staged removal/field preflight. It restores no document snapshot. Actual manual input adds `Manual sentence. ` between the first and later proposal waves. Rejection, one undo, redo, and saved reopening retain that sentence. One undo restores all pending group IDs, markup, and payloads. Empty/repeated actions add no history. Synchronous reentry and document switches refuse. Late locked/read-only groups refuse atomically.

Permanent tests cover mixed native/property/atomic and metadata-only groups, body/header imported groups, both client orders, both reconnect receivers, concurrent decisions, a later unrelated group, owner deletion, unseen peer typing, both actor histories, remote geometry, room snapshots, and DOCX reopening. Structural width/merge/split cases retain independent exact Word fingerprints. Exact/insufficient-capacity cases preserve vectors, payloads, pending groups, and prior redo on refusal. Full flag ledgers still admit metadata-only decisions. The source focused command passes **46 cases, zero failures, 2133 assertions**; `/tmp/codex-r18-first-focused.log`. The complete stable command `node /tmp/codex-r18-private-runner.mjs --jobs 2` passes **1920 cases, zero failures, zero skipped, 160 files**; `/tmp/codex-r18-full-private.log`. The runner retains repository memory, pressure, deadline, and tally checks. Private types pass 1814 files. Build verifies 159 private modules, external native core, and a React-free main. Scoped formatting and root lint pass. Compatibility passes 109 decisions with the earlier native field change unchanged; `.collaboration/changes/sciobot-reject-all-review-groups.json`. This private command adds no published API, shared format, or provider implementation change.

Runtime `refs/backup/codex-r18-runtime` is commit `4f85f1889db04026f871cfd3a102f1456c4ac3a8`, tree `79233dd4ff36f0999c62c96e882f92df766a5079`. Immutable private build `fork-r18` produces `sciobot-docx-298a1c39aa69.tgz`, SHA256 `298a1c39aa69f69b32123960c5336cac3c81ac4cd933543299cbd3fb764fa6d7`. Public input guards prove all package/build/style/lock inputs match the R51 public build. All four reused public archives retain exact verified bytes. Frozen install, five hashes, direct SDK dependency, four public overrides, lock/source records, installed main/collaboration/source bytes, and one physical core/Yjs pass. Actual installed rejection checks pass **46 cases, 2133 assertions**; `/tmp/codex-r18-installed-rejection.log`. App checks pass **53 units**, **16 browser cases**, frontend/harness types across **6641/2207 files**, and production build. The unit runner exits before the build starts. The owned port-8122 server stops after browser testing. App backup `refs/backup/codex-docx-app-r18-final` is commit `1ffe35fc47a8ca579a116622f84b178146022422`, tree `8287fddc71145c4b6dd58352c69e55549058c6cb`. Only package/lock/source records and the added adapter archive differ from verified R51. Both HEADs and real indexes remain unchanged. All 198 reference source hashes remain unchanged. Native tests use the EigenPal Pro License through original test-only fixtures. Private history upgrade, host admission, review UI integration, and cutover gates remain separate. See [implementation notes](evidence/implementation-notes.md).

## R19

- [ ] Replace earlier pending proposals in a later AI wave

**Depends:** [R25](05-suggestions.md#r25), [R15](05-suggestions.md#r15), [R16](05-suggestions.md#r16), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/recompose.ts`; `sibling recompose.test.ts`

**Read:** [S02](sources.md#s02), [S08](sources.md#s08), [S25](sources.md#s25)

**Change:** Orchestrate fresh committed-to-target planning through the owned-proposal replacement transaction. Re-read after each intervening document revision.

**Close when:** Wave two changes text introduced by wave one. Mixed accept, reject, and manual edits survive. No duplicated or nested owned revisions remain.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Verified engine/app R25 backups retain the starting state. `recomposeOwnedProposals` plans a complete Markdown target from a fresh ownership-relative baseline. It retains resolved ownership for undo and explicit host ownership after reload. One staged native transaction supports body and retained-cell text, run and paragraph formatting, existing-link retargets, and paragraph sibling gaps. Inserted paragraphs include marks, links, and inline code. Dependent formatting maps past tracked deletion halves. Exact native receipts keep unrelated formatting groups independent. Content proofs cover saved/current property bags and complete marked paragraphs. Later manual changes refuse before writes. Unselected native revisions retain their content and group membership. Hyperlink destinations retain earlier tooltip deltas in either field order. The mixed files pass 47 cases: 19 local and 28 shared. Eight complete affected files pass 130 cases and 3514 assertions. A complete 175-file run passes 2161 cases before the final tooltip correction. Its final affected-file rerun adds eight cases. Composite verification covers 2169 distinct cases across all 175 files, with exact source manifests and original deadlines. Private types pass 1849 files; formatting and root lint pass. An immutable source build verifies 430 package files and 176 runtime modules with core external. Source backup `refs/backup/codex-r19-mixed-verified-source` retains the verified runtime. Compatibility retains 116 decisions and one inherited format-field change. R19 remains unchecked: tracked lists, remaining table and rich structural execution, host-supplied property/atomic fallback inputs, changed retained atomic target planning, and installed-runtime verification remain required. The isolated app retains verified R25 artifacts. See [implementation notes](evidence/implementation-notes.md).

The next table tranche adds ordered native rows/columns, merge/split, table properties, and allocated new-cell content. Twenty-four local cases and 56 shared cases pass focused checks. The final complete private suite passes 2251 cases across 177 files. Native/export checks verify 3747 distinct cases across 226 files after one unchanged corpus-timeout rerun. Read-only resolved grids and explicit auto table views retain native spans. Public defaults stay GFM. Core/Markdown API checks, types, public builds, and the immutable private build pass. Source backup `refs/backup/codex-r19-table-work` retains this progress. A later correction preserves working native geometry for unchanged GFM targets, body/head edits, and explicit covered-slot splits. Fifteen local and sixteen shared cases pass. Composite full-file verification covers 2282 distinct cases across 179 files after the corrected stale-callback fixture rerun. Native composition now passes 27 cases and sixteen real-session cases, including born-cell geometry, original history, foreign-site refusal, and peer owner deletion. The complete-target compiler now uses deferred native regions after topology allocation. Both original combined-growth probes pass. Twenty-eight direct/tracked cases and 24 real-session cases cover born and retained heads, row/column deletion, splits after allocation, both histories, native decisions, peer owner deletion, and DOCX/room reopening. Unchanged default GFM tables retain native paragraph boundaries during body-only later waves. The final complete private runner passes 2362 cases across 183 files, with no failures or skips. Frozen source hashes, private types, the immutable build, root lint, compatibility, and all 198 reference hashes pass. R19 remains open for its recorded list/checklist, link, rich resource, host fallback, retained atomic target, and installed/application requirements. The app remains on verified R25. See the ordered table section in [implementation notes](evidence/implementation-notes.md).

**List insertion checkpoint:** Branch `codex/r19-birth-lists-work` adds native list bindings to born sibling paragraphs. It uses fresh destination identities for controlled numbering and checkbox seeds. New checkbox controls join the existing content insertion before ownership proofs are captured. This prevents Original-view leakage and receipt changes during collaboration publication. Twelve local cases cover bullet, ordered, nested, and checkbox lists, two replacement waves, independent direct output, exact imported paragraph fingerprints, history, decisions, and saved numbering references. Four real-session cases cover bullet/checkbox waves, both reconnect receivers, independent peer suffixes, both actors' histories, remote caret geometry, and saved convergence. The final two-file run passes 16 cases and 294 assertions; `/tmp/r19-birth-list-final-shared.log`. Private types pass 2066 files; `/tmp/r19-birth-shared-types.log`. Compatibility: `.collaboration/changes/sciobot-tracked-born-lists.json`. R19 remains open. Peer edits inside born items, richer structural targets, host fallbacks, complete native package oracles, and installed-runtime verification remain required.

**Peer text inside born lists:** Eight real-session cases pass 296 assertions in `/tmp/r19-born-peer-shared-final.log`. Bullet and checkbox cases cover both client orders and reconnect receivers. An offline participant types inside the first wave while the owner prepares the next wave. The peer text survives both decisions, owner history, decision history, saved convergence, and Word reopening. Rejection compares the complete saved body fingerprint with independently authored peer-run XML and the original imported paragraph identities. Remote caret checks select the live proposed paragraph after replacement. The earlier fixture selected its retired predecessor, which correctly clamps to offset zero. The first standalone oracle omitted imported Word paragraph identities; its replacement retains the captured original identities before any proposal writes. No output assertion normalizes them away. Private types pass 2068 files; scoped formatting and lint pass. R19 remains open for peer deletion during affected edits, full package acceptance oracles, rich fragments, host fallbacks, and installed-runtime verification.

**Payload replacement checkpoint:** Branch `codex/r19-birth-lists-work` adds explicit host fallback options to the private replacement coordinator. Property/atomic proposals reuse one shared preparer and the existing record schemas. Native body edits, old owned inverses, exact localized Word labels, and payload records commit in one review history unit. The later-wave request captures fallback data before asynchronous reads. Missing fallback data and stale label baselines restore the exact package reference, revision, group membership, and history availability. The full target coordinator supports two property/body waves with intervening manual text. The specialized owned replacement command also supports two complete visual/opaque atomic payload waves, retaining original assets and opaque fields for rejection.

An initial four-session regression detects discarded validation projections entering the observed shared journal. Preparation now uses pure domain/atomic planners and performs no derived package writes. All four real-session property cases pass unchanged output expectations. Both client orders and reconnect receivers cover baseline canonical payload retention, offline peer text, author and reviewer history, remote caret geometry, decisions, and saved convergence. The seven-file affected command passes **90 cases and 754 assertions** in `/tmp/r19-payload-all-focused.log`. Private types pass 2073 files in `/tmp/r19-payload-all-types-final.log`. The new atomic fixture needed an explicit `BlockPayload[]` annotation; runtime expectations remain exact. Build verifies 220 modules with native core external. Root formatting, root lint, and compatibility pass; `.collaboration/changes/sciobot-recompose-payload-stage.json` adds no format fields or published source changes. R19 remains open: complete-target atomic dispatch, rich inserted application fragments, broader affected-peer deletion, complete package acceptance oracles, and installed-runtime verification remain required.

The primary integration is `b8c26928`. All 26 changed private source/test files match the isolated checkout byte for byte. The complete seven-file command passes 90 cases and 754 assertions again with the native paragraph fix included. Primary types pass 2073 files; build verifies 220 modules with core external. Root lint, compatibility, backlog validation, and all 199 production source hashes pass. See [payload replacement checkpoint](evidence/recompose-payload-stage.json). R19 remains unchecked.

**Complete atomic target checkpoint:** The compiler dispatches retained registered atoms and opaque MDX roots through D14 whole-payload plans. It retains block identities, unprojected attributes, payload extensions, and private AST data. Unique explicit child identities follow reordered syntax. Removed children do not donate private data to new children. Opaque root renaming retains stored root data. The compiler validates the serialized result against the exact requested target. It refuses ambiguous stored-to-projected child maps before any write. It never rebuilds imported Word content from Markdown.

Six pure planning cases use independently authored complete payloads, including hidden/null attributes, unknown component names, nested private data, reordering, removal, and root renaming. Two local coordinator cases execute visual/custom MDX targets across two waves, manual native edits, both decisions, exact canonical history, and Word reopening. Four real-session cases cover both client orders and reconnect receivers, offline peer text, untouched canonical payload baselines, both actor histories, decisions, remote caret geometry, saved convergence, and room reopening. The six-file affected command passes **52 cases and 612 assertions** in `/tmp/r19-target-atomic-affected.log`. Private types pass 2077 files; build verifies 221 modules with core external. Root lint and compatibility pass; `.collaboration/changes/sciobot-complete-atomic-targets.json` changes no shared format fields or published source. Root formatting passes in `/tmp/r19-target-atomic-root-format.log`. R19 remains open for rich inserted application fragments, broader affected-peer deletion, complete package acceptance oracles, installed-runtime verification, and any required domain media reconciliation.

The atomic target primary integration is `ebff1210`. Its five changed source/test files match the isolated checkout. The complete six-file command passes 52 cases and 612 assertions again with the native paragraph fix included. Primary types pass 2077 files; build verifies 221 private modules. Root lint, compatibility, backlog validation, and all 199 production source hashes pass. See [atomic target checkpoint](evidence/complete-atomic-targets.json). R19 remains open.

**Latest normalization checkpoint (`f81c4114`):** review decisions now run the existing canonical text normalizer over affected story parts, so accepting a later-wave rich hyperlink/image proposal coalesces adjacent `w:t` nodes deterministically. The combined focused command (`tree-session-stages.test.ts`, `document-port.test.ts`, and `recompose-inline-images-shared.test.ts`) passes 40 cases and 732 assertions; the rich-image batch is 6/8. The two remaining failures are both `peerOffset=10` accept cases during collaborative undo: one replica restores the peer text outside the pending hyperlink (`Before.  peer[Lead…]`) while the other retains it inside. R19 stays open for this ownership-reconciliation defect and the recorded rich/application/runtime requirements.

The focused trace also records the remaining boundary: while the peer is offline, its shared paragraph is plain text with owned-format markers; the later proposal hyperlink exists only in the author’s review structure. Reconnection therefore exposes peer characters between the proposal copies without a shared `w` wrapper to inherit. The repair must carry proposal wrapper ownership through the shared ownership/revision metadata before native materialization.

The born-table Original recovery checkpoint makes Original retain the independent peer suffix at the same boundary as native rejection. The full native layout check passes 158 cases. Four real-session cases verify both client orders, both reconnect receivers, exact native decisions, unchanged source fingerprints, Original-to-rejection caret geometry, author/decision history, and DOCX/shared-room reopening. Core/private types, both builds, root format/lint, compatibility, and all 199 source hashes pass. The remaining rich application fragments, package ledgers, peer deletion evidence, and installed application verification still prevent R19 closure. See the Original recovery section in [implementation notes](evidence/implementation-notes.md).

**Admission and first-list history checkpoint:** Runtime `937bdf32` refuses overlapping replacements that change an unselected presence group's editability. Independent outside proposals remain supported. `prepareNativeDocument` initializes empty numbering infrastructure before sharing, so first-list undo/redo retains its namespace. A missing initialization refuses before writes. The complete frozen affected command passes 765 cases across 58 files. Nine actual compiled shared cases pass against one compiled core. Exact native Word, full state-update refusal, actor/reviewer history, geometry, and DOCX/room checks remain enabled. Types, build, formatting, lint, compatibility, and production reference hashes pass. R19 stays unchecked for its remaining insertion, reconciliation, resource, and application requirements. Exact commands, source hashes, fixtures, results, and limits are in [admission evidence](evidence/presence-admission-and-numbering.json).

**Shared history repair checkpoint:** Runtime `827f0820` publishes evaluation transport state after synchronous Yjs cleanup. Runtime `8d967ee7` and follow-up `9c8a08c8` preserve structural peer-text reconciliation and exact compound typing origins. These repairs resolve all 64 recorded later-wave table/list regressions. Every composer fixture remains byte-identical to `14a60030`. The complete gate passes 333 cases and 16,071 assertions. The collaboration gate passes 235 cases, and the compiled atom/code gates pass 159 cases. Seventy-two exact earlier-reader snapshots retain edits and both histories. No oracle, timeout, shared format or production import boundary changes. See [repair evidence](evidence/composer-shared-history-repairs.md) and its exact commands. R19 remains open for its remaining insertion, reconciliation, resource and installed application requirements.

[Interior atom-table peer continuity](evidence/atom-table-interior-peer-continuity.md) records unchanged native gap identities, whole-control replacement boundaries, and causal terminal ordering. Both leaves remain open.

[D30 acceptance](evidence/d30-closure.md) closes inline atom planning and materialization. R19 remains open for the remaining ownership reconciliation routes.

**Working-property reset checkpoint (`82fb3064`):** pending application payloads now prevent the canonical native no-op shortcut. Exact working-target equality still avoids redundant writes. Returning an owned property to its committed value runs the existing inverse transaction. The focused `recompose-payloads.test.ts`, `recompose-payloads-collaboration.test.ts`, and `recompose.test.ts` command passes 23 cases and 323 assertions. Four new real-session cases cover both client orders and reconnect receivers, offline peer text, reset undo/redo, and saved convergence. Application commit `6e80987c2` installs the exact rebuilt runtime and passes 31 browser cases, 29 unit cases, both typechecks, and a production build. Private mention types also coexist with the host Plate plugin: `5dcc784a` removes the conflicting global mdast augmentation, and the dedicated host fixture and 94 parser regressions pass. Compatibility records are `.collaboration/changes/sciobot-working-property-noop.json` and `.collaboration/changes/sciobot-host-markdown-types.json`. R19 remains open for its recorded rich-image ownership defect and remaining integration requirements.

## R20

- [x] Keep imported and foreign revisions separate

**Depends:** [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/foreign-revisions.ts; `sibling foreign-revisions.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Scope recomposition to application-owned groups. Preserve imported and other-author revisions; report unsupported overlaps before writes.

**Close when:** An edit outside a foreign revision succeeds and preserves it. An unsupported overlap fails without modifying tracking mode or any revision.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Existing `recomposeOwnedProposals` and `replaceOwnedProposals` keep application-owned groups explicit through the ownership ledger. The native inverse planner preserves unselected foreign and imported revisions, while overlapping unselected proposals refuse before changing tracking mode, canonical content, group membership, or history. Edits outside foreign markup still commit and preserve the foreign revision through undo/redo and DOCX reopening.

The focused local run passes 3 cases and 29 assertions (`/tmp/r20-foreign-focused.log`), covering preserved foreign markup, saved reload without adopting foreign groups, and atomic overlap refusal. The shared run passes 8 cases and 1,232 assertions (`/tmp/r20-foreign-shared.log`) across both client orders and reconnect receivers, including peer edits, histories, decisions, geometry, and reopening. No new source file was needed because the existing owned-proposal planner is the production boundary. R19, host integration, and final cutover gates remain separate.

## R21

- [x] Restore pending review state after reload

**Depends:** [B02](02-react-blocks.md#b02), [R14](05-suggestions.md#r14), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/restore.ts`; `fork:packages/sciobot/test/browser/restore.tsx`; `fork:e2e/sciobot-review-restore.spec.ts`

**Read:** [S08](sources.md#s08), [S04](sources.md#s04)

**Change:** Recover groups, property proposals, and active review availability from saved DOCX data. Do not require a running chat session.

**Close when:** Reload a nested property and text review. Counts and decisions match the saved state; no orphan review controls remain.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** B02/R14/V01 are complete. `restoreReviewState` reads current canonical DOCX groups, rebound property/atomic plans, pending-review presence, and intrinsically reviewable group IDs. It uses the current load identity/revision and requires no chat state. It changes no content, metadata, display mode, shared state, or history. Native permissions/history/capacity gates remain authoritative on decisions. Read-only imported groups remain visible. Invalid persisted chains refuse without writes. Fully decided records remain stored without stale review controls.

Eight local cases and sixteen real-session cases prove nested property/text/atomic reload, current owner bindings, mixed prior decisions, immutable repeated reads, imported body/header groups, exact Word/payload values, both client orders/reconnect receivers, DOCX/room reopening, accept/reject undo/redo, unseen peer input, remote selection geometry, deletion during peer editing, and saved convergence. Reopened rooms use fresh deterministic client IDs. The complete stable command `node /tmp/codex-r21-private-runner.mjs --jobs 2` passes **1944 cases, zero failures, zero skipped, 162 files**; `/tmp/codex-r21-full-private.log`. It retains repository process isolation, memory/pressure gates, deadlines, and tally behavior. Private types pass **1817 files**. The external-runtime build verifies **160 modules**, one external core, and a React-free main. Source formatting, root lint, compatibility, source coverage, backlog, and diff checks pass. Compatibility records 110 decisions with the earlier native field change unchanged; `.collaboration/changes/sciobot-restore-review-state.json`. No published API or shared format changes occur.

`bunx playwright test -c e2e/sciobot-review-restore.config.ts` passes **six Chromium cases**; `/tmp/codex-r21-browser-first.log`. Actual hard reloads restore nested review from only DOCX bytes and optional native room data. Counts, IDs, pending properties, per-group/bulk decisions, fresh native history, real keyboard input and model/DOM selection, both peers, remote geometry, mixed decisions, owner deletion, and resolved-state reload pass. No stale group controls remain. The React fixture uses one native document-change subscription and immutable-package caching through `useSyncExternalStore`. Playwright owns strict port 5194 without server reuse and releases it after success. R21 browser types pass **1564 files** with `bun check -p tsconfig.review-restore-browser.json`. The broader existing browser configuration still fails at unchanged `test/browser/selection.tsx:58`, TS2769 for the old registration object. The exact same error is reproduced with the R18 baseline file list; `/tmp/codex-r21-baseline-browser-types-verified.log`. That fixture and the shared browser config remain unchanged. This broader check is recorded as failing. All new recovery browser code passes its dedicated check.

Runtime `refs/backup/codex-r21-runtime` is commit `0b7a45d22d786c9b0fdf00031da255720958cca5`, tree `962cbafe51f0fb432666912acf1cceb5bd481aba`. Immutable private build `fork-r21` produces `sciobot-docx-02aab5f6480c.tgz`, SHA256 `02aab5f6480cf624af257619593936f7a34da2e2493e9402a6bc3c1259061f11`. Public input guards and all four reused public hashes match R51/R18. Installed main/collaboration/source bytes, five hashes, direct SDK dependency, four overrides, frozen lock, line endings, and one physical core/Yjs pass. Actual installed recovery passes **24 cases, 1659 assertions**; `/tmp/codex-r21-installed-recovery.log`. App checks pass **53 units**, **16 browser cases**, frontend/harness types across **6642/2208 files**, and production build. The unit runner exits before build. The owned port-8123 server stops after browser success. App backup `refs/backup/codex-docx-app-r21-final` is commit `162ce100a2731f449d18550fd15c0f40f0d48549`, tree `7b2eb284e1cb0d19306c48ec1ed6d12e1a9ae62b`. Before/after-build app trees match. Only package/lock/source records and the added adapter archive differ from R18. Both branch HEADs, real indexes, and all 198 reference source hashes remain unchanged. Original native transport fixtures use the EigenPal Pro License and stay outside installed archives. R22 retains external Word reconciliation. Application review UI integration, host admission, and cutover remain separate. See [implementation notes](evidence/implementation-notes.md) and [host recovery](evidence/review-history-scope.md#recover-review-state-after-docx-reload).

## R22

- [x] Reconcile property proposals after external Word review

**Depends:** [R12](05-suggestions.md#r12), [R21](05-suggestions.md#r21), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/reconcile-external.ts`; `fork:packages/sciobot/src/review/external-review-state.ts`; `fork:packages/sciobot/src/review/reconcile-external.test.ts`; `fork:packages/sciobot/test/support/external-review.ts`

**Read:** [S09](sources.md#s09), [S25](sources.md#s25)

**Change:** Reconcile retained metadata with Word fallback revisions after external accept or reject. Preserve ambiguous edits and report their conflict.

**Close when:** Fixtures from Word acceptance and rejection select the correct property values. An externally edited fallback never silently rolls back.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** R12/R21/V01 are complete. Pure inspection compares exact retained native fallback ownership and recorded text endpoints. Reconciliation selects recorded property/atomic values and appends existing immutable decision receipts in one guarded native history transaction. It preserves native Word content, unrelated payload values, opaque metadata, original proposal bytes, and ambiguous edits. Partial review, changed fallback text, overlapping proposals, missing owners, and changed committed fields remain explicit conflicts. Metadata-only equal labels remain pending. Group, bulk, legacy property/atomic decisions, and reconciliation share one reentry guard. Complete existing removal/field budgets apply before mutation. The stable private suite passes **2007 cases, zero failures, zero skipped, 165 files**; `/tmp/codex-r22-full-private.log`. Focused local checks pass 89 cases and 787 assertions. Actual installed bundles pass **63 external-review cases and 3828 assertions**; `/tmp/codex-r22-installed-reconciliation.log`. These include 27 local, 24 real-session, and 12 capacity cases. Both client orders, reconnect receivers, both actors' repeated history, peer fallback edits, owner deletion, geometry, and DOCX/room reopening pass. All eight captured property-presence fixtures retain absence, null, false, zero, and unrelated fields. Private types pass 1824 files; build verifies 163 modules, external core, and a React-free main. Root lint, formatting, compatibility, source coverage, and backlog checks pass. Compatibility records 111 decisions with the earlier native field change unchanged. Runtime `refs/backup/codex-r22-runtime` captures commit `8dc49ba06d6b05b1e1391d26fef861d2207abaf9`, tree `fc2c8faa8c96d89a64b994d1acd19a4d05afe277`. The app installs `sciobot-docx-7e0d4419e150.tgz`; all four public archives retain verified R51/R21 bytes. Archive hashes, installed bytes, immutable provenance, dependencies, lock, line endings, and one physical core/Yjs pass. App units pass 53 cases; frontend/harness types pass 6645/2211 files; production build passes. The initial browser run records 15 passes and one ordinary typing-history failure. The unchanged eight-case file passes after the full source suite exits, with the same deadline. All eight direct-edit cases already pass in the initial run. Logs preserve both results without claiming a proven timing cause. App backup `refs/backup/codex-docx-app-r22-final` is `2279c1312f93583438c5130d1aa8a4e09c022cce`, tree `1264a3986a73b25bb9cb651c622121393bd673da`. Before/after-build app trees match. Both branch HEADs, real indexes, and all 198 reference source hashes remain unchanged. Independent fixtures model documented WordprocessingML acceptance/rejection; they are not desktop Word captures. Actual desktop Word version, visual evidence, and saved-file interoperability remain V06. No public, native shared-format, or evaluated provider implementation changes occur. Original test transport uses the EigenPal Pro License and stays outside archives. Host admission, application review UI, and cutover remain separate. See [implementation notes](evidence/implementation-notes.md) and [reconciliation behavior](evidence/review-history-scope.md#reconcile-externally-reviewed-fallback-text).

## R23

- [x] Route manual suggestion mode through native tracking

**Depends:** [R04](05-suggestions.md#r04), [R05](05-suggestions.md#r05), [R06](05-suggestions.md#r06), [R14](05-suggestions.md#r14), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/manual-tracking.ts`; `sibling manual-tracking browser test`

**Read:** [S01](sources.md#s01), [S07](sources.md#s07), [S25](sources.md#s25)

**Change:** Connect the user suggestion toggle to native tracked typing, deletion, and formatting. Keep manual authors separate from AI sessions.

**Close when:** Manual insert, delete, and bold changes create review items. AI recomposition leaves those items intact.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Implemented from `127d6be6` on `codex/react-block-proof`. `review/manual-tracking.ts` connects a native suggestion toggle without replacing the configured user author. It checks mode gates before execution, prevents deferred author requests after refusal, flushes buffered typing under its original mode, and preserves native selection. AI executors retain explicit attribution and ownership. No shared schema or public document API changes.

`bun test --preload ./scripts/test/preload-dom.ts` across `manual-tracking.test.ts`, `manual-tracking-collaboration.test.ts`, `recompose.test.ts`, `recompose-mixed.test.ts`, and `recompose-collaboration.test.ts` passes 52 cases and 1134 assertions (`/tmp/sciobot-r23-native.log`). The strengthened ten manual cases pass 240 assertions (`/tmp/sciobot-r23-native-verified.log`). Complete native paragraphs, proposed bold properties, exact revision records, and configured authors survive two AI waves and both AI decisions. Same-label user/AI authors never grant ownership. Mode transitions, missing author/module refusal, exact package history, and Word reopen pass.

Four real-session cases cover both client orders, concurrent AI/user edits, both AI decisions, independent peer history, reconnect, deletion during peer editing, remote selection geometry, complete native recovery, and equal saved package digests. One strengthened run hit the original five-second deadline during parallel checks. The unchanged focused files pass in 7.60 seconds total after other checks finish; deadlines and oracles stay unchanged.

`bunx playwright test --config e2e/sciobot-manual-tracking.config.ts` passes two Chromium cases in 30.1 seconds (`/tmp/sciobot-r23-final-browser.log`). Real toggle clicks, typing, Backspace, bold clicks, two AI waves, accept/reject, exact history, mode-off typing, complete paragraph preservation, and equal saved Word packages pass. Private/browser type checks pass 2352/1965 files. The private build verifies 307 modules with native core external. Scoped formatting and lint pass. Compatibility decision: `.collaboration/changes/sciobot-manual-tracking.json`. Only the private adapter changes; no published release bump is required. Transport remains test-only under the EigenPal Pro License.

## R24

- [x] Plan removal of owned pending revisions

**Depends:** [R01](05-suggestions.md#r01), [R02](05-suggestions.md#r02), [R03](05-suggestions.md#r03), [R14](05-suggestions.md#r14), [R52](05-suggestions.md#r52)

**Primary files:** `fork:packages/sciobot/src/review/owned-proposal-plan.ts`; `sibling owned-proposal-plan.test.ts`

**Read:** [S08](sources.md#s08), [S25](sources.md#s25)

**Change:** Identify owned pending revision sites and their inverse operations without touching foreign sites. Preserve original canonical node references.

**Close when:** A pure plan removes only selected owned groups. Original text and property presence match the committed view without a whole-tree replacement.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, guarded runtime `refs/backup/codex-r52-guard-runtime`, commit `92ec6f1e573fb7ef256b09f8b4889816c5c2f249`; all dependencies are closed. `planOwnedProposalRemoval` requires explicit host-retained group ownership and exact complete native/property receipts. It re-reads current groups and validates property/atomic fallbacks. Its detached, frozen, revision-bound result contains scoped native inverse operations, property rejection intents, and imported-address retention intents. It contains no canonical snapshots or mutation callback. Pure native previews preserve unselected revision identities and authored attributes. The planner changes no package, canonical reference, display mode, editor, or history. Twenty-five focused cases pass 181 assertions; `/tmp/codex-r52-foreign-property-planner.log`. They include all eight captured property-presence fixtures, same-attribution foreign groups, native text/formatting, atomic and metadata-only groups, nested revisions, multiple stories, native table properties/merge/split, exact independent Word output, original references and prior redo, bounded hostile inputs, and retained independent pending paragraph properties. Complete private source passes 2052 cases across 167 files. The actual installed planner export passes its complete file within the 178-case runtime matrix. Private/core types, immutable builds, formatting, root lint, 198-source-hash audit, and compatibility pass. `.collaboration/changes/sciobot-owned-proposal-removal-plan.json` records no shared-format impact. The pure function needs no live-session test or published changeset. R52 supplies its verified native prerequisite. App backup `refs/backup/codex-docx-app-r52-final` preserves the matching installed runtime. R25 retains re-reading and atomic native/metadata application with replacement proposals; R19 retains later-wave behavior. See [implementation notes](evidence/implementation-notes.md) and [review history scope](evidence/review-history-scope.md).

## R25

- [x] Apply an owned-proposal replacement transaction

**Depends:** [R24](05-suggestions.md#r24), [D15](04-edit-planner.md#d15), [V01](11-verification.md#v01), [R53](05-suggestions.md#r53)

**Primary files:** `fork:packages/sciobot/src/review/replace-owned-proposals.ts`; `sibling replace-owned-proposals.test.ts`

**Read:** [S08](sources.md#s08), [S25](sources.md#s25)

**Change:** Apply owned revision removal and replacement proposals inside one validated transaction. Restore the prior state on failure.

**Close when:** Replacing a pending inserted word succeeds without nested owned revisions. Injected failure restores canonical content, metadata, and history.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Runtime `refs/backup/codex-r25-runtime` records commit `377ae00ff7420168aede906e286fd819ee302630`. `replaceOwnedProposals` validates exact host ownership and the current revision inside the native gate. It applies selected native inverses, immutable property decisions, and replacement body text in one publication and history unit. Replacement coordinates use the intermediate document. Other replacement kinds remain unsupported. Detached native previews measure transient removals, fields, and package cascades before shared admission. Fingerprints preserve unselected revisions, including matching-author revisions. Pure inverse planning runs before the observed stage; this prevents preview removals from entering the live journal. The live stage requires the exact validated starting package. Sixteen local cases pass 101 assertions; sixteen shared cases pass 568 assertions. They cover pending-word replacement without nested owned revisions, save/reopen, exact rollback, ownership receipts, prior redo, concurrent edits/deletion, both client orders, remote caret geometry, and capacity refusals. The broad run records 2087 passes and nine timeouts across 170 files. Complete sequential reruns of its three affected files pass all 74 cases without assertion or deadline changes. Combined file coverage verifies 2096 distinct cases. Actual installed bundles pass 102 cases across six complete replacement/text/cell files. Private types pass 1833 files. App checks pass 53 units, frontend/harness types across 6647/2213 files, production build, frozen installation, exact archive bytes, and one physical core/Yjs. Formatting, root lint, compatibility, and all 198 reference source hashes pass. Four public archives remain byte-identical to R53; only the private SDK rebuilds. App backup `refs/backup/codex-docx-app-r25-final` retains the verified installation. `.collaboration/changes/sciobot-replace-owned-proposals.json` records no shared-format change. R19 retains later-wave orchestration. See [implementation notes](evidence/implementation-notes.md) and [review history scope](evidence/review-history-scope.md).

## R26

- [x] Add canonical tracked hyperlink target changes

**Depends:** [D03](04-edit-planner.md#d03), [R04](05-suggestions.md#r04), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/store/tree-op-tracked-hyperlink.ts`; `sibling native tracked hyperlink regression`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Add the smallest native tracked mutation for one existing hyperlink target or tooltip. Preserve its run content and unrelated shared relationships.

**Close when:** Original and Proposed views retain their respective URLs and tooltips. Accept keeps the target; reject restores the original. A URL shared by another link stays unchanged.

**Check:** Cover native revision identity, atomic refusal, undo/redo, save/reopen, and two-session concurrency/reconnect. Update published contracts, API snapshots, changesets, and collaboration decisions.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `setHyperlinkTarget` accepts optional `revision` (`tree-op-tracked-hyperlink.ts`). Word has no property record for a link target, so the tracked form replaces the link: the old link keeps its target and its runs are struck through the native tracked `deleteText`; a copy with the new target/tooltip follows as one `w:ins` with exactly copied runs (properties and unknown children). The untracked retarget runs outside journal capture to compute attributes only, so peers never see the link re-aimed outright. Relationships are never rewritten. Resolution now hoists markers out of a link it empties of runs. Subranges, links with fields, notes, drawings, bookmarks/range markers, or nested revisions refuse; proofing marks are allowed. Core: 6 cases (URL, bookmark, tooltip-only: accept equals the direct command ignoring proofing marks, exact reject, save/reopen both decisions; copied run properties and retained link attributes; no-op; atomic refusals). Two-session coverage runs through R27's executor (below). Limits: concurrent text typed at the link end reads before the copy in Proposed for clients 73/41 (never lost; R43). Changeset `.changeset/tracked-hyperlink-targets.md`; decision `.collaboration/changes/tracked-hyperlink-targets.json`; compatibility passes. API snapshot refresh shares R11's pending build. Docs: tracked-changes guide and hyperlink feature row.

## R27

- [x] Execute tracked hyperlink plans

**Depends:** [D03](04-edit-planner.md#d03), [R26](05-suggestions.md#r26)

**Primary files:** `fork:packages/sciobot/src/review/track-hyperlinks.ts`; `sibling track-hyperlinks.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Lower D03 hyperlink target and tooltip deltas through R26. Revalidate old values and package revision before one guarded transaction.

**Close when:** A URL-only AI suggestion changes no text or marks. Accept matches direct native editing; reject restores the original link. The review receipt retains stable native addresses.

**Check:** Cover accept/reject, stale plans, shared relationship isolation, undo/redo, save/reopen, and the focused two-session scenario.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `trackHyperlinkPlan` rechecks the package revision and each D03 old target/tooltip, previews every retarget and both decisions on an immutable package with scratch relationship mints, then mints relationships inside one gated, separately undoable transaction and returns stable Word addresses per link. Tooltip removal refuses (D27). 12 native cases: URL, bookmark, tooltip, URL+tooltip; Original keeps old URL/tooltip, Proposed reads the new one with identical text and marks; accept equals A01 direct editing, reject restores the original fingerprint; the twin link sharing the relationship keeps it; one-step undo/redo; save/reopen addresses and both decisions; stale/invalid/wrong-old-value/tooltip-removal refusals; no rPrChange written; editing/suggesting preserved. 12 two-session cases: both client orders and receivers, accept/reject, peer text, remote caret geometry, proposal and decision undo/redo, saved digests; a peer appending at the link end (text kept once, order limit R43); a peer deleting the link text offline (the proposal copy remains the pending link). Decision `.collaboration/changes/sciobot-tracked-hyperlinks.json`. `link`/`unlink` edits remain R42.

## R28

- [x] Resolve local collaboration IDs to native parts

**Depends:** [V01](11-verification.md#v01), [R04](05-suggestions.md#r04)

**Primary files:** `fork:packages/core/src/editor/surface-scope.ts`; `fork:packages/core/src/editor/__tests__/local-node-part-routing.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Resolve live collaboration IDs without a part prefix through canonical package indexes. Keep story stores unopened.

**Close when:** Cell proposals remain reviewable after collaboration undo/redo. Body and related-part local IDs resolve to their actual parts.

**Check:** Add native lookup regressions and run the focused two-session cell history regression. Include a core changeset and compatibility decision.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `packages/core/src/editor/surface-scope.ts` now resolves unprefixed collaboration site IDs through canonical body/related-part indexes without opening story stores. Two native routing cases pass. Review command/resolution regressions and related story checks pass 74 tests with 421 assertions across five files. The focused real two-session cell history subset passes nine tests with 158 assertions: review decisions work again after redo, through reconnect, with remote cursor geometry and saved convergence. Core typecheck/build and its 15-entry API check pass with no new snapshot drift. Private typecheck/build, scoped Prettier, root lint, diff check, and compatibility policy pass. Core patch changeset: `.changeset/local-collaboration-node-routing.md`; decision: `.collaboration/changes/local-collaboration-node-routing-compatibility.json`. No shared formats change. R29 retains two separately reproduced peer-suffix failures; R07 remains unchecked. The complete private suite currently reports 318 pass and two fail. See [implementation notes](evidence/implementation-notes.md).

## R29

- [x] Preserve concurrent peer text beside tracked cell replacements

**Depends:** [D07](04-edit-planner.md#d07), [R04](05-suggestions.md#r04), [R28](05-suggestions.md#r28), [V01](11-verification.md#v01), [R39](05-suggestions.md#r39), [R40](05-suggestions.md#r40)

**Primary files:** `fork:packages/sciobot/src/review/track-table-cells-collaboration.test.ts`; native tracked-text/provider boundary responsible for revision scope

**Read:** [S25](sources.md#s25)

**Change:** Fix the native/provider boundary that can apply a tracked deletion to a peer suffix the proposer never observed. Preserve the peer's text and marks.

**Close when:** Concurrent `Two` to bold `Changed` and a peer suffix produce Original `Two Bob` and Proposed bold `Changed` followed by plain ` Bob`. Accept/reject and proposal undo/redo preserve the suffix in both sessions.

**Check:** Keep the focused two-session regression. Check both reconnect delivery orders, native review decisions, remote selections, and saved convergence. Add the relevant changeset and compatibility decision for the actual fix.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Deterministic regression: `packages/sciobot/src/review/track-table-cells-collaboration.test.ts`, fixed clients 73/41, receivers 0 and 1. Concurrent replacement of `Two` with bold `Changed` and peer append ` Bob` preserves Original `Two Bob` but hides the suffix in Proposed. Clients 41/73 and other-cell races pass. The live regression still reports nine passes and two failures. The complete private suite after B05 reports 342 passes and two failures. A [plain Yjs reproduction](evidence/r29-yjs-format-boundaries.json) reproduces the same leak with no native editor or evaluated transport imports. Both reconnect orders fail when the proposer client is 73 and the peer is 41. Both shared documents converge on the wrong revision scope. A native-only display or operation-order patch cannot repair this boundary behavior. The MIT-only diagnostic lives at `packages/sciobot/test/diagnostics/yjs-format-boundaries.ts`; run its recorded command to reproduce it. The shared-text inspection [artifact](evidence/r29-concurrent-attributes.json) confirms that the merged `Two Bob` segment inherits the deletion, old text shell, bold, and formatting revision from observed `Two`. This happens inside the evaluated provider, before native projection. A fix must bound revision and formatting ownership to observed characters; changing native display reads cannot repair this shared state. Required support cannot close through refusal or skipping these cases. The evaluated transport remains test-only under the EigenPal Pro License.

R39 closes the observed-attribute protocol foundation. Its exported boundary preserves scoped native formatting and portable redo lineage in twenty real-session cases. R29 remains open for production integration and deterministic native replacement placement. Moving the replacement alone does not bound inherited attributes. Attribute repair alone leaves client-dependent suffix placement. Both requirements remain necessary.

**Closed by:** claude, branch `codex/react-block-proof`. Native baseline `9770e78`, local working tree. Both requirements are now integrated. `packages/sciobot/src/review/leading-replacements.ts` runs inside the guarded `tracked-inline` transaction and moves only this transaction's fresh insertion wrappers ahead of their touching deletions; R40 pairs those insertion-first halves. An unseen append therefore follows the replacement in either client order. The fixture enables the R39 owned-text boundary before bootstrap. The expanded matrix exposed one more boundary defect: with clients 73/41, receiver 1, and an italic peer suffix, the owned-text repairs recursed until the stack overflowed. The receiver's repair relayed immediately, ahead of the other peer's offline structs. The partial receiver then applied its deletions, repaired using incomplete intents, and reverted the first repair. `owned-text.ts` now defers repair while `doc.store.pendingStructs` or `pendingDs` exist. It repairs every deferred text once the dependencies integrate. Four new MIT-only partial-delivery cases in `owned-text.test.ts` pass; with the deferral disabled, three of them fail with non-convergent relays. Focused run: `track-table-cells-collaboration.test.ts` (27 cases: accept/reject × plain/italic peer suffix × both fixed client orders × both reconnect receivers, other-cell races, proposal and decision undo/redo, raw-room reopen with decisions, remote caret geometry, observer publications without foreign revision/bold wrappers, saved digests, owner deletion), `replacement-boundary-collaboration.test.ts`, and `packages/sciobot/src/collaboration/`: 77 pass, 1971 assertions. Complete private suite: 1039 pass, 0 fail, 22170 assertions, 86 files; the two historical R29 failures are gone. Private build passes with 109 modules and the native core external. Private typecheck reports one unrelated error in a concurrent worker's new `src/diff/target-c01-matrix.test.ts`; touched files typecheck. Scoped Prettier/oxlint pass. Compatibility: `.collaboration/changes/sciobot-replacement-boundary.json`; policy passes with 55 decisions. No published package source changed, so no changeset is needed. The legacy Yjs diagnostic still reproduces the unprotected leak. The evaluated transport remains test-only and unchanged.

## R30

- [x] Execute tracked checklist control changes

**Depends:** [D06](04-edit-planner.md#d06), [P21](03-agent-projection.md#p21), [R03](05-suggestions.md#r03), [R04](05-suggestions.md#r04), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/track-checklists.ts`; `sibling track-checklists.test.ts`

**Read:** [S04](sources.md#s04), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Apply checklist creation, checked-state, nesting, and removal proposals. Keep the application-owned control identity and unrelated form controls. Persist old values when native revisions do not encode checkbox state.

**Close when:** Accept matches the direct native result for all four changes. Reject restores the original control state, glyph, indentation, and text. Later proposals preserve unresolved older decisions through save/reopen.

**Check:** Add native regressions with explicit Original/Proposed results, review decisions, undo/redo, and save/reopen. Use two real sessions for state/text concurrency, reconnect, deletion, peer-preserving history, and remote selection geometry.

**Claim:** codex checklist agent, branch `codex/react-block-proof`. **Evidence:** `trackChecklistPlan` uses the guarded native review coordinator. Native `w:del`/`w:ins` wrappers retain complete old/new checkbox controls; `w:pPrChange` retains paragraph properties. Deleted glyphs serialize as `w:delText`. Original control IDs survive rejection; application control IDs survive accepted state changes. Nesting preserves authored glyph formatting. Checked-only changes preserve existing numbering properties. Detached native transactions isolate direct previews from the live primitive journal. Each proposal emits one live transaction. Unchanged properties retain older revision addresses. Same-control state and nesting proposals remain independently reviewable in both orders after saving. Unrelated form controls remain unchanged. Conflicting changes to an already replaced control require explicit owned-group replacement; no older proposal is silently absorbed.

Focused verification passes 70 tests and 1,145 assertions across four files. Nineteen native checklist cases cover exact decisions, emitted journals, saved deletion markup, identities, stale plans, locks, and later proposals. Thirty-two real two-session cases cover both client orders and receivers, peer insertion/deletion, history, saved convergence, reconnect, and caret geometry. Fifteen direct-list regressions and four shared-allocation regressions pass. Scoped formatting, lint, and diff checks pass. Root integration passes private types across 1975 files, the private build with 196 modules, repository lint, and compatibility policy with 129 decisions and no format changes. The real transport fixture uses the EigenPal Pro License. Compatibility decision: `.collaboration/changes/sciobot-tracked-checklists.json`. The private executor uses existing native revision markup and review storage. No shared schema, protocol version, or published package changeset changes.

## R31

- [x] Preserve inserted row content through native review

**Depends:** [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/store/table-authoring-plan.ts`; `fork:packages/core/src/store/store/tree-op-tables.ts`; `fork:packages/core/src/automation/__tests__/table-authoring.test.ts`; `fork:packages/sciobot/src/diff/table-rows-collaboration.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Populate fresh tracked rows under their owning row revision. Preserve inherited paragraph-mark formatting without redundant empty formatted runs after acceptance. Keep existing cell value changes independently tracked.

**Close when:** Accepting a populated tracked row matches the direct native fingerprint, including formatted empty-paragraph seed runs. Rejecting restores the original. Two real sessions retain values and formatting through concurrency, undo/redo, reconnect, deletion during peer editing, and save/reopen.

**Check:** Add native core regression coverage and the focused two-session scenario. Run core typecheck/build/API checks, a changeset, and the compatibility policy.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Fresh row initial values now belong to their row insertion revision, preserving inherited paragraph-mark formatting without an extra empty run after acceptance. Existing cell changes retain inline attribution. Both modes create consistent fresh row/cell property holders; allocation budgets include them. Imported source rows remain unchanged during creation. Two new core regressions retain exact direct/accepted fingerprints and rejected originals, including absent source property holders and inherited font/size values. The seven-file native command in [implementation notes](evidence/implementation-notes.md) passes 196 cases and 757 assertions. `packages/sciobot/src/diff/table-rows-collaboration.test.ts` passes both fixed-client reconnect orders: two cases and 144 assertions. Two real sessions cover populated row creation, concurrent outside edits, peer-preserving proposal undo/redo, remote caret geometry inside new cells, accept and decision undo/redo, deletion during peer editing, reconnect, and saved pending/accepted/final digest convergence. The independent direct oracle retains landed new paragraph identities after redo and compares exact accepted fingerprints. The final row/cell/plan subset passes 35 cases and 360 assertions. The full private suite reports 515 pass and the same two R29 failures across 517 tests, 49 files, and 6,891 assertions. Core/private typecheck/build, scoped format/lint, root lint, diff checks, and compatibility policy pass. Scoped core API extraction checks 15 entries without drift; the unscoped command lacks PDF declaration artifacts. Changeset: `.changeset/tracked-row-content-shape.md`. Compatibility record: `.collaboration/changes/tracked-row-content-shape.json`; no shared format changed. The private build verifies 67 source modules and one external core. R08 retains complete row execution, rich D16 content, and imported deletion-history property absence. R29 remains open. The evaluated transport remains test-only under the EigenPal Pro License.

## R32

- [x] Persist independent application property proposal records

**Depends:** [B02](02-react-blocks.md#b02), [D13](04-edit-planner.md#d13), [R03](05-suggestions.md#r03), [V01](11-verification.md#v01), [R33](05-suggestions.md#r33)

**Primary files:** `fork:packages/sciobot/src/review/property-record-schema.ts`; `fork:packages/sciobot/src/review/property-record-store.ts`; `fork:packages/sciobot/src/review/stored-property-proposals.ts`; `sibling property-record-store.test.ts`; `sibling property-records-collaboration.test.ts`

**Read:** [S04](sources.md#s04), [S09](sources.md#s09), [S25](sources.md#s25)

**Change:** Persist one bounded custom XML node per pending property proposal. Retain explicit old/proposed presence, attribution, stable block IDs, and native revision references. Rebind source addresses from stable block IDs after reload. Feed stored proposals into pure committed/working reads. Preserve future-version records and unrelated OPC content.

**Close when:** Known records save/reopen with exact semantic values and native address references. Malformed stores, duplicate identities, unsupported versions, and missing block owners have explicit outcomes. Independent records survive offline two-session creation, reconnect, peer-preserving undo/redo, and removal. Pending records never replace live story trees or lose unrelated payloads. R12 still owns actual fallback revisions, property transactions, and decisions.

**Check:** Run focused schema/store/review tests and the two-session carrier scenario. Check save/reopen convergence, private typecheck/build, scoped format/lint, root lint, compatibility, source coverage, and backlog validation. Carrier tests do not establish live R12 execution.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Twelve schema/store cases prove versioned per-proposal records, exact presence, attribution, native address retention, future opaque data, limits, strict ownership, causal order, removal markers, and native save/reopen. Reserved `scio-review_` nodes share the existing application payload store. Block readers keep this family separate. This avoids a second custom part's concurrent first-create race. Stored stable block IDs rebind to changed canonical owners after source insertion and reload. Committed/working reads automatically include stored proposals. Six real-session carrier cases cover both fixed client orders and reconnect receivers, independent offline record creation, peer text, selection geometry, proposal/removal undo/redo, owner deletion during peer updates, and saved convergence. Empty anonymous undo wrappers stay inert; nonempty malformed nodes refuse. Removal markers retain field identities and reserve resolved IDs. Fifty-four focused adapter tests pass with 813 assertions across eight files. Full private suite: 823 pass and the same two R29 failures, 825 tests, 72 files, 17167 assertions. Private typecheck passes 1634 files; build verifies 99 private modules, one external core peer, and a React-free main. Scoped formatting/lint, root lint, compatibility, source coverage, backlog, and diff checks pass. Compatibility: `.collaboration/changes/sciobot-property-record-storage.json`; 45 decisions, with the earlier native format-field change unchanged. Generic native representation remains compatible. All 196 application source hashes and revision `9693712d` remain unchanged. This private schema needs no published changeset or public API extraction. R12 remains unchecked for tracked metadata/fallback transactions and actual accept/reject. These carrier checks do not prove those required operations. See [record format](evidence/property-review-record-format.md) and [implementation notes](evidence/implementation-notes.md).

## R33

- [x] Give native automation transactions explicit undo boundaries

**Depends:** [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/editor/paginated-surface-contract.ts`; `fork:packages/core/src/editor/paginated-surface.ts`; `fork:packages/core/src/editor/__tests__/automation-history-boundary.test.ts`; `fork:packages/sciobot/src/review/property-records-collaboration.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Allow an application-owned automation call to request a separate native undo unit. Preserve existing streaming capture defaults and write gates. Property record transactions use this boundary rather than joining a later user or decision edit.

**Close when:** Adjacent explicit calls undo separately, default streaming calls keep existing behavior, and failed/viewing calls preserve state/history. Real sessions preserve peer content through separate proposal/removal units and undo/redo. No test introduces a timing delay to force separation.

**Check:** Run focused native history tests and the two-session carrier scenario, core/private typecheck/build, API extraction, scoped format/lint, root lint, compatibility, and backlog checks. Add a changeset for the published native contract.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Optional `separateHistory` supplies the existing native fresh-transaction policy through `applyAutomationOps`. Existing callers retain their prior capture policy and all write gates. Three native cases verify adjacent body/package transactions, separate undo/redo, default local history, and no-op/refused/viewing calls without timed separation. The broader five-file native history subset passes 45 cases and 256 assertions. R32's six real-session carrier cases verify independent proposal/removal history, peer text preservation, reconnect, owner deletion, geometry, and saved convergence. Core typecheck passes 1456 files. Core build passes; scoped API extraction/check scans 15 entries with no errors or drift. Private typecheck/build, scoped format/lint, root lint, compatibility, source coverage, backlog, and diff checks pass. Full private suite: 823 pass and the same two R29 failures. Changeset: `.changeset/automation-explicit-history-boundary.md`. Compatibility: `.collaboration/changes/automation-explicit-history-boundary.json`; existing native tokens, operations, and journals remain compatible. No Pro implementation changes. R12 retains actual property/fallback execution. See [implementation notes](evidence/implementation-notes.md).

## R34

- [x] Commit package records after native identity allocation

**Depends:** [R33](05-suggestions.md#r33), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/binding/tree-session-contract.ts`; `fork:packages/core/src/binding/tree-session-apply.ts`; `fork:packages/core/src/editor/paginated-surface-contract.ts`; `fork:packages/core/src/editor/paginated-surface.ts`; `fork:packages/core/src/editor/__tests__/automation-post-ops.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Add post-operation package edits to the existing guarded native transaction. Receipts can read actual allocated revision identities before publication. Preserve default pre-operation resource hooks, actor binding, native validation, and history.

**Close when:** A post-operation record contains the actual native revision address in one publish and one undo unit. A failing post hook rolls back text, package metadata, revision, and history. Viewing and native operation failures never run post hooks. Both real sessions retain record addresses through synchronization and history.

**Check:** Run native transaction/history tests, focused two-session property cases, core/private typecheck/build, API extraction/check, scoped format/lint, root lint, compatibility, source coverage, and backlog checks. Add the published native changeset.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `TreeApplyOptions.packageEditsAfterOps` runs package hooks after native operations and related-part operations. The single-story route also completes native collaboration normalization before the hooks. The surface exposes optional `packageEditPhase: 'after'`; default pre-operation hooks retain existing behavior. All hooks remain inside the original gated transaction and native actor/history context. Four native cases prove actual tracked identity receipts, one publication, one undo unit, exact redo, complete rollback after thrown hooks, skipped hooks on operation/viewing refusal, unchanged before-hook order, and atomic-session rollback. The broader native history subset passes 49 cases and 287 assertions across six files. R12's twelve real-session cases prove receipt addresses through concurrency, reconnect, decisions, history, deletion, geometry, and saved convergence. Core typecheck passes 1456 files; core build and scoped API extraction/check pass. The API check scans 15 entries with no errors or drift. Private typecheck/build, scoped formatting/lint, root lint, compatibility, source coverage, backlog, and diff checks pass. Changeset: `.changeset/automation-post-operation-package-hooks.md`. Compatibility: `.collaboration/changes/automation-post-operation-package-hooks.json`; no shared operations, journals, or transport version fields change. Full private suite: 864 pass and the same two R29 failures. No Pro implementation changes. See [implementation notes](evidence/implementation-notes.md).

## R37

- [x] Preserve offline peer text through middle paragraph review joins

**Depends:** [R35](05-suggestions.md#r35), [R36](05-suggestions.md#r36), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/store/paragraph-mark-merge.ts`; `sibling revision-paragraph-merge.ts`; `fork:packages/core/src/store/package/retired-paragraph.ts`; `fork:packages/core/src/layout/story-roots.ts`; `fork:packages/core/src/editor/body-paragraph-section-index.ts`; `fork:packages/core/src/store/__tests__/joined-paragraph.test.ts`; `fork:packages/sciobot/src/review/track-blocks-collaboration.test.ts`; `sibling joined-paragraph-conflicts.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Retain recoverable identity when a review decision removes a middle paragraph mark. Route unseen peer text to the native forward join without losing text, duplicating it, or changing untouched sibling identities. Preserve ordinary Word join behavior and explicit owner deletion.

**Close when:** Rejecting an inserted middle paragraph while a disconnected peer adds `Peer` produces `Opening` then `PeerClosing`. Accepting a deleted middle paragraph preserves its peer text at the same native join. Both client orders and reconnect receivers pass creation/decision history, geometry, owner deletion, and save/reopen.

**Check:** Keep all eight exact failing R05 expectations for insertion and deletion. Add owner-deletion counterparts and native joined-versus-terminal regressions. Check pending formatting/section markers and pure Word export. Run focused native and real-session suites.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78` plus local changes. Actor-bound decisions retain source owners and runs with a resolved generic joined marker. Every view joins peer text into the native survivor; pure Word export performs physical joining without live mutation. All eight original failing insertion/deletion expectations pass unchanged. The focused command in [implementation notes](evidence/implementation-notes.md) passes 55 cases and 957 assertions: 19 native cases and 36 real-session cases. Both client orders and reconnect receivers cover creation/decision history, remote geometry, explicit owner deletion, whole-section rejection, normal split/join editing, table-cell Backspace/range replacement, and save/reopen. Native cases cover orphan fallback, owner boundaries, section/format metadata, source run formatting, the survivor's paragraph properties, copy isolation, malformed markers, and pure snapshots. Body-section lookup now includes merged members, invalidates newly visible identities, and accepts actual opaque collaboration IDs. Two focused section cases pass 11 assertions. The store/binding regression passes 941 cases and 4554 assertions across 34 files. Standard-floor display checks plus an unchanged isolated corpus rerun cover all 593 selected cases across 59 files; no source timeouts or assertions are relaxed. Converter suite: 213 pass, 1241 assertions, 31 files. Full private suite: 951 pass and the same two R29 failures across 953 cases, 80 files, and 19792 assertions. Core/private typecheck pass 1459/1652 files; both builds pass. The private build verifies 104 modules and one external core peer. Root/scoped lint, formatting, compatibility, source coverage, API check, and backlog checks pass. API baseline: 15 entries, zero errors/drift, unchanged 2443 warnings and 380 forgotten exports. Changeset: `.changeset/joined-paragraph-peer-recovery.md`; decision: `.collaboration/changes/joined-paragraph-peer-recovery.json`, with no new shared fields. All 196 audited application hashes remain unchanged. No Pro, application, or Plate implementation changes. N02/N11 retain production control contexts. R05 still needs N18/R38 inline-code completeness; R19 retains later-wave ownership.

## R38

- [x] Preserve inline code in newly tracked paragraphs

**Depends:** [N18](07-native-content.md#n18), [R35](05-suggestions.md#r35), [P22](03-agent-projection.md#p22)

**Primary files:** `fork:packages/sciobot/src/review/new-paragraphs.ts`; `fork:packages/sciobot/src/review/track-blocks.test.ts`; `focused inline-code projection regression`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Lower compiled inline-code tokens through N18's canonical native representation. Retain code semantics in Original/Proposed Markdown without flattening text or rebuilding imported paragraphs. Preserve neighboring marks and link ownership.

**Close when:** A new paragraph containing plain text, inline code, emphasis, and a hyperlink accepts/rejects exactly. Code text containing backticks, XML characters, and surrogate pairs survives history and save/reopen. Native source style definitions remain intact.

**Check:** Add native AST parity assertions and two real sessions with an independent peer edit. Unsupported-operation refusal does not complete this required capability.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, current private worktree. `trackSiblingPlan` requests compiled code ranges and applies the existing `codeNewParagraphs` helper after hyperlink wrapping and native paragraph allocation. Only new paragraphs receive code identity. Their original insertion revision owns the content; unchanged siblings gain no formatting revision. Eleven new cases assert exact ASTs for plain text, backtick/XML/emoji code, emphasis, and code inside titled hyperlinks. Accept/reject, creation and decision history, pending/resolved save/reopen, unchanged source-style fingerprints, and original sibling object identities pass. A conflicting foreign style refuses the entire proposal with no publication, revision, history, or content changes. Eight new real-session cases cover both client orders, reconnect receivers, independent peer text, remote code-range rectangles/carets, proposal and decision undo/redo, and saved convergence. Four acceptance cases also race paragraph deletion against a peer code change and restore both through deletion undo. Tests read current inserted IDs after history recreation; native Word revision receipts remain stable for decisions. Command: `bun test --preload ./scripts/test/preload-dom.ts --timeout 60000 packages/sciobot/src/review/track-blocks-inline-code.test.ts packages/sciobot/src/review/track-blocks.test.ts packages/sciobot/src/review/track-blocks-collaboration.test.ts`; 66 pass, 1327 assertions, `/tmp/codex-r38-siblings-final.log`. Full private command: `bun test --preload ./scripts/test/preload-dom.ts --timeout 60000 packages/sciobot/src`; 1355 pass, zero failures, 27635 assertions, 117 files, `/tmp/codex-r38-private-full.log`. Private typecheck (1734 files), build (132 modules, one external core), scoped format/lint, root lint, and compatibility pass. Root lint retains existing warnings. Decision: `.collaboration/changes/sciobot-tracked-sibling-execution.json`; 92 decisions. No published API, shared schema, or provider change. No production EigenPal Pro License dependency. N20's application Chromium run separately passes 13 cases after document preparation wiring.

## R39

- [x] Bound shared-text attributes to observed characters

**Depends:** [C03](01-contracts.md#c03), [R28](05-suggestions.md#r28), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/collaboration/owned-text.ts`; `sibling ownership codec and position helpers`; `focused owned-text.test.ts`; `fork:packages/sciobot/test/support/collaboration.ts`

**Read:** [S25](sources.md#s25)

**Change:** Add an application-owned MIT Yjs boundary that records exact observed character ownership for native attributes. Repair inherited foreign revision/property attributes before native observers read shared text. Keep evaluated transport implementation code unchanged. Preserve operation activity and restored-character lineage through undo/redo.

**Close when:** Native bold proposals and disconnected plain/marked peer text retain exact Original/Proposed content in both client orders and reconnect receivers. Native observers never publish a foreign suffix inside the proposal. The protocol handles UTF-16 text, embeds, history, owner deletion, raw-room reopen, and DOCX export without leaking metadata into Word.

**Check:** Keep the independent legacy failure. Add primary MIT-only protocol tests and real native-session cases. Verify deterministic convergence, bounded parsing, lifecycle disposal, package/runtime boundaries, and a documented protocol compatibility decision. Do not claim complete replacement placement from attribute scoping alone; R29 retains that integration.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. The optional private `./collaboration` entry supplies observed insertion baselines, causal formatting records, active scope attributes, before-observer repair, and portable restored-character lineage over external MIT Yjs 13.6.32. Production source imports no editor, Pro, editor-api, Plate, or React runtime. Eighteen MIT-only cases and twenty real-session native cases pass: 38 cases and 873 assertions. The native matrix covers plain/italic peer suffixes, accept/reject, both fixed client orders and reconnect receivers, proposal and decision undo/redo, observer publications, remote caret geometry, raw-room continuation, DOCX reopen, metadata-free exported XML, and owner deletion history. The unchanged legacy diagnostic retains its failure. The independent leading-replacement proof still passes all four order combinations. The fixture enables the production boundary explicitly before bootstrap; its evaluated transport implementation remains unchanged. The complete private suite passes 989 cases with the same two R29 failures, 991 cases, 83 files, and 20665 assertions. Private typecheck passes 1659 files. The build verifies 108 private modules, one external core peer, and a React/Yjs-free document entry. Scoped formatting/lint, root lint, compatibility, source coverage, and backlog checks pass. Compatibility: `.collaboration/changes/sciobot-owned-text-attributes.json`; 53 decisions, with the earlier native format-field change unchanged. The application protocol requires every participant to attach version 1. Legacy rooms require stopped-writer export and fresh bootstrap. Host negotiation, production integration, and large-document performance retain their application/cutover gates. All 196 audited application hashes remain unchanged; only the external application revision advanced to `c842beea4`, and the coverage snapshot records that revision. No published package implementation changes require a changeset or API extraction. R29 remains open for native replacement placement and its unchanged original expectations. See [protocol contract and checks](evidence/owned-text-protocol.md) and [implementation notes](evidence/implementation-notes.md).

## R40

- [x] Pair insertion-first native replacement revisions

**Depends:** [R04](05-suggestions.md#r04), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/store/review-items.ts`; `fork:packages/core/src/editor/review-replacement-pairs.ts`; `focused native pairing regression`; `fork:packages/sciobot/src/review/leading-replacement-collaboration.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Recognize touching native insertion/deletion halves in either physical order. Preserve deletion-first address/range semantics in the paired item, stable pair identities, query opt-out, and independent authors or nesting. Keep ordinary native replacement landing unchanged.

**Close when:** Imported insertion-first replacements form one decision. Accept/reject, per-half query reads, proposal placement, and decision history remain exact. Both real sessions retain an independent peer edit through decisions, history, geometry, and save/reopen.

**Check:** Add native tests for both orders and unrelated halves. Run the focused native and two-session suites, scoped typecheck/build, API checks, lint/format, compatibility, coverage, and backlog validation. Add a published changeset. R29 retains actual replacement generation and observed attribute integration.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Native adjacency reads accept either physical order. The editor pairs deletion-first candidates first, preserving existing precedence for ambiguous chains. Pair objects retain deletion-first addresses/ranges and stable keys. Five native cases pass 71 assertions for both orders, both decisions, query opt-out, exact struck-first caret placement, activation, history, foreign authors, and ambiguity. Twelve real-session cases pass 348 assertions: both decisions, fixed client orders, reconnect receivers, independent peer text, remote caret geometry, history, save/reopen, and owner deletion during peer editing. A separate native review/tracked regression passes 404 cases and 2065 assertions across 24 files. Core/private typecheck pass 1459/1660 files; both builds pass. Scoped formatting/lint, root lint, API extraction/check, compatibility, coverage, backlog, and diff checks pass. API check: 15 entries, zero errors or drift, unchanged 2443 warnings and 380 forgotten exports. Changeset: `.changeset/leading-native-replacement-pairs.md`; compatibility: `.collaboration/changes/leading-native-replacement-pairs.json`, 54 decisions. Shared format fields remain unchanged. The feature matrix and tracked-changes guide describe both orders and preserved ambiguity precedence. R29 retains generated replacement placement and observed attribute integration. No evaluated implementation changes.

## R41

- [x] Execute tracked rich fragments in new rows and columns

**Depends:** [R08](05-suggestions.md#r08), [R09](05-suggestions.md#r09), [D28](04-edit-planner.md#d28), [D29](04-edit-planner.md#d29), [D30](04-edit-planner.md#d30), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/inserted-row-content.ts`; `fork:packages/sciobot/src/review/track-table-rows.ts`; `sibling track-table-rows.test.ts`

**Read:** [S03](sources.md#s03), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Land new-row and new-column cell streams that contain tables, code, application blocks, payload records, and inline atoms through the shared D28-D30 materializer. Keep the content under its row or cell insertion revision.

**Close when:** A tracked new row and a tracked new column with a nested table, a code block, an application block with its payload, and an inline image accept to the A01 direct output for the same plan. Rejection restores the exact original package, including payload and media parts.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Tracked row and column executors now reuse the shared D28-D30 fragment materializer for code blocks, nested tables, and registered application payloads. New fragments remain under the enclosing row/column insertion revision; plain paragraphs, links, lists, and cell properties retain their existing path. The final tracked row/column batch passes 38 cases and 662 assertions (`/tmp/r41-final.log`), including row collaboration, accept/reject, undo/redo, save/reopen, and a rich-row regression with code, nested table, and application payload. The shared prepared-image runtime materializes one-pixel images inside tracked rows and columns; local rejection removes the structural insertion without leaked media or payload state. The combined row/column and core lock batch passes 114 cases and 840 assertions (`/tmp/r41-col-final-reg.log`); Sciobot typecheck passes 2,425 files (`/tmp/r41-col-final-types.log`). The two-editor column oracle passes 2 cases and 28 assertions (`/tmp/r41-column-collab-image.log`) across both client orders, with peer text, review acceptance, history, save/reopen, and ownership health. Both tracked row and tracked column image fragments use prepared resources, remain under structural ownership, and accept/reject without leaked media or payload state. Compatibility: `.collaboration/changes/sciobot-tracked-image-column-ownership.json`; changesets: `.changeset/tracked-rich-cell-fragments.md`, `.changeset/tracked-image-row-ownership.md`, `.changeset/tracked-image-column-ownership.md`.

## R42

- [x] Execute tracked hyperlink insertion and removal plans

**Depends:** [D27](04-edit-planner.md#d27), [R26](05-suggestions.md#r26), [R27](05-suggestions.md#r27), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/store/tree-op-tracked-hyperlink.ts`; `fork:packages/sciobot/src/review/track-hyperlinks.ts`; `sibling track-hyperlinks.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Lower D27 `link` and `unlink` edits as native suggestions. Wrapping text in a new link and removing a link keep the text and marks, using the R26 replacement form: the old runs are struck and their copy, inside or outside the link, follows as a tracked insertion.

**Close when:** A tracked link and a tracked unlink accept to the A01 direct output for the same plan and reject to the original paragraph. Relationships shared with other links stay unchanged.

**Check:** Add native regressions with save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Runtime commit `93f2ad71` completes exact A01 acceptance and original rejection, including pure disjoint links after save/reopen. Source tests retain imported metadata, shared relationships, independent decisions, history, reconnect, peer deletion, and geometry checks through real sessions. The scoped native suite passes 9081 checks across 552 files with zero failures; all 3455 frozen input hashes remain exact. The private broad run passes 2615 cases and finds one stale 2.31 peer-version assertion after the upstream 2.32 upgrade. That assertion is corrected; the package guard and complete concurrent-provenance file pass 19 checks and 307 assertions, including two added pure-link peer cases. The installed browser matrix passes all 22 link/unlink decisions with exact saved and local fingerprints. Application commit `8bc4e7d9b` installs the verified runtime: all 456 archive files match, one core resolves, frozen installation passes, frontend/harness types pass, and 30 integration tests pass. The build verifies 4824 source files and 188 private runtime modules. All normal commit hooks pass. R19 and R43 remain separate. See [verification record](evidence/tracked-hyperlinks.json).

## R43

- [x] Place concurrent link-boundary text deterministically beside tracked retargets

**Depends:** [R26](05-suggestions.md#r26), [R39](05-suggestions.md#r39), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/store/tree-op-tracked-hyperlink.ts`; `fork:packages/sciobot/src/review/track-hyperlinks-collaboration.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Make a peer's text typed at the end of a link, while an offline peer proposes its retarget, read after the proposed copy in both client orders. The struck link keeps the original characters' collaboration identity, so a concurrent deletion of the link text keeps deleting the struck half.

**Close when:** The R26 end-of-link case reads `the site!` in Original and Proposed for clients 41/73 and 73/41 and both reconnect receivers, and the deletion case still keeps the proposal as the only pending link.

**Check:** Keep the focused two-session regressions. Check both decisions, history, and saved convergence.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Tracked hyperlink copies now opt out of source-text lineage capture while retaining fresh native IDs and review ownership. This keeps a peer deletion of the original source from deleting the pending retarget copy. The full collaboration suite passes 12 cases and 314 assertions (`/tmp/r43-fixed-full.log`) across both client orders, both reconnect receivers, both decisions, peer boundary typing, peer deletion, history, and saved convergence. Core tracked-link and copy tests pass 65 cases and 1,045 assertions (`/tmp/r43-final-all2.log`). Core types pass 1,616 files; private types pass 2,423 files. Compatibility decision: `.collaboration/changes/sciobot-hyperlink-copy-boundary.json`; Changeset: `.changeset/hyperlink-copy-ownership.md`. Existing source-lineage capture remains enabled for other native text copies.

**Investigation:** The earlier boundary-order prototype remains historical context. The production fix uses the supported canonical copy seam: retarget copies receive fresh native IDs without source-text lineage, while ordinary native text copies retain lineage capture. This isolates pending review copies from concurrent deletion of their original source and closes the required deletion and boundary cases.

## R44

- [x] Initialize independent review group storage before native collaboration

**Depends:** [B02](02-react-blocks.md#b02), [R32](05-suggestions.md#r32), [N20](07-native-content.md#n20), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/group-record-store.ts`; `sibling group-record-store.test.ts`; `fork:packages/sciobot/src/native/prepare-document.ts`; `fork:packages/sciobot/test/support/collaboration.ts`

**Read:** [S07](sources.md#s07), [S18](sources.md#s18), [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Add bounded application-owned custom XML storage for independent review group records. Initialize its native part before an editor opens or a collaboration room bootstraps, following the N20 preparation contract. Use stable group IDs and native Word revision membership; never persist load-specific canonical node IDs. Keep group records separate from block payload and property proposal schemas. Initialization preserves original native story/furniture/media bytes and unrelated parts. Extend the collaboration fixture to use this same initialization at room creation.

**Close when:** Preparation is idempotent and creates the group store without a user undo step. Existing code styles, conflicting imported styles, native content, and unrelated custom XML parts remain intact. Two offline actors can add independent records, undo/redo one actor, reconnect, and save/reopen without deleting the other actor's record or its shared part. No application group namespace enters native Word content. R14 then replaces its failing native-attribute persistence with this carrier.

**Check:** Add native storage and preparation regressions with exact retained part bytes, malformed/unknown records, resource limits, save/reopen, and history. Add focused two-session carrier tests in both client orders and reconnect directions. Verify no production provider or EigenPal Pro License runtime dependency enters the adapter.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, private working tree. The independent native custom XML carrier uses one immutable record per node, stable compact group IDs, and exact native part/revision membership. It rejects canonical node IDs and paths, malformed records, duplicate IDs/members, unsafe values, and bounded resource overflow. Unknown future records retain exact raw data and refuse replacement/removal. Removed IDs remain reserved with native field identity for undo. Full app preparation initializes code styles and this store before opening; conflicting imported code styles remain exact. Group-only preparation supports test hosts. Room creation initializes the carrier once; joined room snapshots stay unchanged. Changed-part serialization preserves native story, furniture, media, and unrelated custom XML bytes. Twenty-one cases pass with 376 assertions across storage, full preparation, and six real-session cases; `/tmp/codex-r44-focused.log`. Both client orders and reconnect directions cover independent offline records, peer body input and geometry, peer-preserving undo/redo, cross-actor removal, owner deletion, and saved reopening. Private typecheck passes 1749 files; build verifies 139 modules, one external core peer, and a React-free main. Scoped lint/format, compatibility, source coverage, backlog, and diff checks pass. Decision: `.collaboration/changes/sciobot-review-group-storage-initialization.json`; 97 decisions, with the earlier native format-field change unchanged. No production provider, published API, shared-format implementation, or EigenPal Pro License runtime dependency changes. This private carrier needs no published changeset. All 196 audited application source hashes remain unchanged. R14 now uses this carrier and passes the strict original XML comparison. R45 reserves retained native IDs. The app now pins the verified N21 runtime, including this initialized carrier.

## R45

- [x] Reserve retained review revision IDs before allocating new native revisions

**Depends:** [R34](05-suggestions.md#r34), [R44](05-suggestions.md#r44), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/groups.ts`; `fork:packages/sciobot/src/review/group-record-store.ts`; `fork:packages/sciobot/src/review/groups-collaboration.test.ts`; `fork:packages/core/src/store/store/tree-op-revision-ids.ts`

**Read:** [S07](sources.md#s07), [S18](sources.md#s18), [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Prevent new application/native tracked edits from reusing Word revision IDs retained by group records. Native allocation currently scans live revision elements only. Fully resolved records remain needed for native decision undo. A later edit by the same actor can reuse one retired ID with the same author/date, making the old group partly live and read-only. Reserve historical native revision IDs in the allocator before creation and attribution renumbering. Use existing native seams where sufficient. If needed, add a narrow host allocation reservation seam, with no Sciobot namespace dependency in core. Preserve native actor allocation and exact final executor receipts. Never repair this by changing author/date, dropping unresolved history, persisting canonical IDs/paths, or relaxing the read-only conflict gate.

**Close when:** Both actors can decide a native group and then create another group, including in the same paragraph with the same attribution/date. Existing and retired addresses remain distinct. Undo/redo in either actor restores the correct prior group while preserving the later peer group. Both client orders, reconnect receivers, shared-room and DOCX reopening retain exact identity and native decisions. Imported ambiguous address reuse remains separately visible and read-only. The R14 collaboration regression passes without dropping inactive records or weakening native XML comparisons.

**Check:** Add native allocator reservation regressions for live and retired IDs, wrap/limits, actor allocation, formatting, text, structural creation, and attribution renumbering. Use real two-session tests for decisions followed by new edits, offline creation, peer-preserving cross-actor undo/redo, deletion, geometry, reconnect, and saved convergence. Keep public API/compatibility evidence exact if a native infrastructure seam changes.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, current fork. `RevisionIdReservationSource` supplies bounded pure host reads before package/session transactions. The optional review contribution forwards it to the package store. `runWithRevisionIdReservations` gives immutable preflights the same native allocation scope. Core knows no application namespace. Owned decimal reservations retain actor allocation, normalize numeric spellings, respect Word bounds, wrap safely, and restore nested scopes after return or throw. Text, format, structure, attribution renumbering, and clipboard allocation use the retained IDs. The private source reads active, inactive, and removed group records without modifying them. Each text proposal now uses a separate native undo unit; rapid same-actor proposals cannot merge through the collaboration capture window. Fifteen native cases cover live/retired IDs, decimal spellings, malformed/resource input, exhaustion, nesting, ownership, all native allocation lanes, clipboard, source refresh, and history. Four private cases cover retained/tombstoned records, normal suggesting input, low-level native session writes, and unknown/malformed data without writes. Ten real-session cases cover both actor orders and reconnect receivers, accept/reject followed by same-paragraph proposals, original/proposed text, exact distinct native IDs, offline peer groups, cross-actor undo/redo, remote geometry, deletion during peer input, restored peer text, room/DOCX reopening, and saved convergence. The eight previously failing R14 cases also pass. Focused tests: 66 pass, zero failures, 1051 assertions; `/tmp/codex-r45-focused-final.log`. Full private suite: 1460 pass, zero failures, 29728 assertions, 125 files; `/tmp/codex-r45-private-full.log`. Native regression, module/isolation, and lane checks pass 124/24/36 cases. Core/private typechecks pass 1471/1754 files; core build/declarations and the 140-module external-core private build pass. The store options moved into a type file to retain the existing line cap. Scoped format, root lint, source audit, compatibility, backlog, and diff checks pass. The new public infrastructure has a minor changeset: `.changeset/retained-native-revision-ids.md`. Public type exports include the snapshots and container types passed to the source. API extraction/check scans all fifteen core entries with zero errors and no new forgotten exports; `/tmp/codex-r45-api-check-final.log`. Exactly the old `editor:OoxmlPackage` baseline entry is pruned because it is now exported. Decision: `.collaboration/changes/native-revision-allocation-reservations.json`; 98 compatible decisions, with the earlier native format-field change unchanged. No provider, shared operation schema, journal version, or EigenPal Pro License implementation changes. Native author/date values and retained records stay intact. All 196 audited application sources remain unchanged; reference HEAD `89d7a134`. The app pins the verified N21 runtime `f8a2d15e1a686fd504d8ae6255cab3d5dfca68a1`. Its native core is the unchanged published R14 artifact. Installed group/history/reopen checks and app regressions pass; see N21.

## R46

- [x] Store independent application block decision receipts

**Depends:** [R12](05-suggestions.md#r12), [R13](05-suggestions.md#r13), [R44](05-suggestions.md#r44), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/payload-group-decisions.ts`; `sibling payload-group-decisions.test.ts`; `fork:packages/sciobot/src/review/group-record-store.ts`

**Read:** [S07](sources.md#s07), [S08](sources.md#s08)

**Change:** Add bounded immutable per-block decision receipts to the preinitialized independent group carrier. Keep property proposal JSON unchanged. Persist proposal/block identity and accept/reject disposition without canonical node IDs. Preserve native-only records, unknown future data, and independent history. Read concurrent equal dispositions as resolved and opposite dispositions as an explicit conflict. Group decision commands remain R15/R16.

**Close when:** Different block decisions from one proposal merge without replacing shared proposal JSON. Undo removes only the actor's receipt and retains the peer's receipt. Conflicting decisions stay visible as conflicts. Native revision reservation accepts known decision receipts without interpreting them as native membership.

**Check:** Test strict decoding, immutable writes, future/raw preservation, resource bounds, save/reopen, and legacy carrier initialization. Use real two-editor sessions for both actor orders and reconnect receivers, independent offline decisions, history, peer text, deletion, remote geometry, and room/DOCX reopening.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, private working tree. The real-session shared-JSON probe produces `shared-proposal:invalid-json` after concurrent removal of different blocks; `/tmp/codex-r46-shared-json-probe.log`. Immutable schema-2 receipts instead use fresh independent nodes in the existing prepared group carrier. Each receipt stores only its own ID, proposal ID, block ID, and disposition. Original multi-block proposal JSON and native XML stay unchanged. Strict bounded decoding rejects extra canonical addresses, malformed IDs, live getters, cycles, depth/node/byte overflow, changed IDs, and invalid labels. Native-only schema-1 records retain their exact format and reservation behavior. Schema 3 and later retain exact raw JSON and refuse interpretation or replacement. Legacy proposals need no rewrite; live writes still require carrier initialization before room creation. Equal same-group dispositions retain both receipt IDs. Opposite dispositions remain explicit conflicts without a winner. Fourteen real-session cases cover both client orders and reconnect receivers, independent offline decisions, both actors' receipt undo/redo, peer text, owner deletion with peer edits inside the deleted owner, restored geometry, room/DOCX reopening, and saved convergence. Focused command `bun test packages/sciobot/src/review/payload-group-decisions.test.ts packages/sciobot/src/review/payload-group-decisions-collaboration.test.ts packages/sciobot/src/review/group-record-store.test.ts packages/sciobot/src/review/group-allocation.test.ts` passes 31 cases and 531 assertions; `/tmp/codex-r46-focused-final.log`. Full private `bun run test` passes 1484 cases, zero failures, 30279 assertions, 128 files; `/tmp/codex-r46-private-full.log`. Private typecheck passes 1760 files. Build passes 141 private modules, one external core, and a React-free main. Scoped formatting, root lint, source audit, backlog, and diff checks pass. Compatibility passes 100 decisions with the earlier one-field format change unchanged; `.collaboration/changes/sciobot-payload-group-decision-receipts.json`. No public API, published package, native shared format, or evaluated implementation changes. The application remains pinned to the verified N21 runtime; this storage primitive enables no production decision command. R15/R16 retain exact native/payload resolution, disposition filtering, legacy whole-proposal decisions, and fallback revision-ID retention. R22 retains conflict reconciliation. See [storage format and upgrade limits](evidence/payload-group-decision-format.md) and [implementation notes](evidence/implementation-notes.md).

## R47

- [x] Keep stored application labels atomic under concurrent decisions

**Depends:** [B27](02-react-blocks.md#b27), [R46](05-suggestions.md#r46), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/blocks/payload-label.ts`; `fork:packages/sciobot/src/blocks/payload-opc-codec.ts`; `sibling payload-label.test.ts`

**Read:** [S07](sources.md#s07), [S08](sources.md#s08), [S10](sources.md#s10)

**Change:** Replace character-merged application record-label updates with a bounded atomic scalar representation. Preserve created payload JSON, independent payload field registers, native Word fallback text, unknown application metadata, and any existing native binding semantics. Define legacy-read and initialization behavior without inventing a second document pipeline. Keep native/shared transport schemas and evaluated implementation unchanged.

**Close when:** Both peers accepting the same pending group retain exactly `New middle` in the stored label. Independent conflicting labels converge to one complete authored value. Equal labels never duplicate. Either actor's undo preserves peer labels and payload fields. Saved DOCX and reopened rooms retain the same complete value. Legacy records, XML whitespace, Unicode, bounds, and any binding compatibility have explicit proof.

**Check:** Keep the two failing R15 cases enabled. Add strict codec and two-real-session label regressions for equal/different offline writes, both actor orders and reconnect receivers, both actors' history, owner deletion, peer body edits, geometry, and saved convergence. Verify imported/foreign Word parts remain intact. Record the private metadata compatibility and rollout limits.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, private working tree. Every later label update appends a version-1 `labelValue` child with one complete JSON string value. Independent native child identities retain each actor's update through undo/redo. Document order selects one whole authored value. Created label/data text, existing payload registers, opaque record metadata, foreign stores, Word parts, and native bindings stay unchanged. Legacy records need no conversion or eager initialization. Empty undo wrappers remain inert. Future label versions and malformed/duplicate fields refuse without source mutation. Unsupported future payloads retain opaque label semantics. Bounds cover individual strings, encoded values, authored rows, native slots, and aggregate bytes. A single shared attribute prototype fixed text merging but lost a peer-preserved value after undo/redo; immutable rows replace it. Seven local cases and eighteen real-session cases prove equal/different concurrent labels, legacy/prepared rows, both client orders and reconnect receivers, both actors' history, peer fields/text, owner deletion with edits inside the deleted owner, remote geometry, and room/DOCX reopening. Command `bun test packages/sciobot/src/blocks/payload-label.test.ts packages/sciobot/src/blocks/payload-label-collaboration.test.ts packages/sciobot/src/review/accept-group.test.ts packages/sciobot/src/blocks/payload-opc-codec.test.ts packages/sciobot/src/blocks/payload-fields.test.ts` passes 48 cases and 1069 assertions; `/tmp/codex-r47-focused-final.log`. Full private `bun run test` passes 1534 cases and fails two stronger native acceptance-history cases across 1536 tests, 132 files, and 31776 assertions; `/tmp/codex-r47-private-full.log`. The former label assertions pass. The surviving failures restore a native insertion wrapper while the peer's accepted receipt stays active. They remain enabled under R48. Typecheck passes 1769 files. Build passes 145 private modules, one external core, and a React-free main. Scoped formatting, root lint, source audit, compatibility, backlog, and diff checks pass. Decision `.collaboration/changes/sciobot-immutable-label-registers.json`; 102 decisions, earlier one-field native format change unchanged. Older private readers ignore the new rows, so upgrade private participants together. Native opaque transport needs no shared-format change. No published package, public document API, core/provider implementation, or evaluated implementation changes. The application stays pinned to N21 until complete group acceptance passes. See [format and rollout limits](evidence/payload-label-format.md).

## R48

- [x] Preserve a peer's native review decision through local decision history

**Depends:** [R46](05-suggestions.md#r46), [R47](05-suggestions.md#r47), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/accept-group.ts`; `fork:packages/sciobot/src/collaboration/review-history.ts`; `fork:packages/sciobot/src/review/group-history.ts`; `fork:packages/sciobot/src/review/accept-group-collaboration.test.ts`

**Read:** [S07](sources.md#s07), [S08](sources.md#s08)

**Change:** Keep canonical native revision state consistent with surviving independent group decision receipts through either actor's undo/redo. Use the existing native pipeline and history gates. If a native infrastructure seam is missing, implement the smallest generic Apache core capability. Keep application metadata out of core and evaluated implementation changes out of the production path.

**Close when:** Two peers accept the same application group offline. Both resolve one native fallback. Undoing one actor retains the other's accepted native state and exact Original/Proposed Word content. No pending insertion wrapper or extra native group appears while a peer receipt remains active. Undoing the final remaining decision restores the same pending group and its original native membership. Redo restores acceptance without adding unrelated history entries or dropping peer text.

**Check:** Keep both R15 same-group regressions enabled. Cover both actor/client orders and reconnect receivers, each actor's undo/redo, overlapping peer text, owner deletion, remote geometry, room/DOCX save/reopen, Original/Proposed fidelity, and final saved convergence. Inspect current native/collaboration history interfaces before choosing a mechanism. Do not hide restored native revisions or change the expected results to match the defect.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, private working tree. The unchanged same-group oracle passes when the established owned-text protocol attaches before bootstrap; `/tmp/codex-r48-owned-probe.log`. `attachOwnedReviewHistory` now owns that lifecycle and binds the native session through a receiver-preserving facade. Document mismatch, duplicate ownership, late legacy attachment, poisoned protocol state, disposed scopes, and destroyed documents refuse explicitly. Functions keep stable identity and native private-field receivers, including frozen transports. `acceptReviewGroup` checks the live capability before shared mutation; solo behavior stays unchanged. A shared `Symbol.for` key crosses the separate private bundles without loading Yjs from the main entry. Built-entry use/disposal smoke passes; `/tmp/codex-r48-entry-smoke.log`. No hidden group filtering, history replay, extra normalization transaction, native core change, provider edit, or new shared protocol exists. The V01 owned fixture binds the scope before bootstrap, including room snapshot loads. Four local scope cases and twelve real-session cases cover both client orders, both reconnect receivers, both first actors, each actor's undo/redo, final remaining undo, peer text inside the actual fallback, owner deletion, remote geometry, exact Word Original/Proposed text, canonical revision-site disappearance/restoration, room/DOCX reopening, and saved convergence. The next normal undo removes only prior typing, proving no extra history entry. Both original regression cases stay enabled with unchanged result assertions. Focused command `bun test packages/sciobot/src/collaboration/review-history.test.ts packages/sciobot/src/review/group-history-collaboration.test.ts packages/sciobot/src/review/accept-group-collaboration.test.ts` passes 32 cases and 1648 assertions; `/tmp/codex-r48-focused-final.log`. Full private `bun run test`: 1552 pass, zero failures, 32948 assertions, 134 files; `/tmp/codex-r48-private-full.log`. Typecheck passes 1773 files. Build passes 147 private modules, one external core, and a React-free main. Scoped format, root lint, source audit, backlog, and diff checks pass. Compatibility passes 103 decisions with the earlier native format-field change unchanged; `.collaboration/changes/sciobot-owned-review-history-scope.json`. No published package or public document API changes. All 196 audited application source hashes remain unchanged at revision `89d7a134`. R15 retains final acceptance/runtime verification; R16 retains group rejection. See [scope lifecycle and compatibility](evidence/review-history-scope.md).

## R49

- [x] Redo a complete native property group after concurrent peer acceptance history

**Depends:** [R14](05-suggestions.md#r14), [R48](05-suggestions.md#r48), [R50](05-suggestions.md#r50), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/review/accept-group.ts`; `fork:packages/sciobot/src/review/group-history.ts`; `fork:packages/sciobot/src/collaboration/review-history.ts`; `fork:packages/sciobot/src/collaboration/review-flags.ts`; `fork:packages/sciobot/src/collaboration/review-flag-codec.ts`; `fork:packages/sciobot/src/collaboration/review-flag-items.ts`; `fork:packages/sciobot/src/review/accept-native-group-collaboration.test.ts`; `fork:packages/sciobot/src/review/native-property-history-collaboration.test.ts`

**Read:** [S07](sources.md#s07), [S08](sources.md#s08)

**Change:** Preserve semantic complete-group redo for native table property revisions after two actors accept, then both undo. Native history must re-resolve every restored member. Keep the owned text history solution intact. Use existing canonical operations and exact stable Word membership. Inspect the current public history and journal seams before choosing the smallest private or generic Apache capability. Keep evaluated implementation edits and imports outside the production path.

**Close when:** Both peers accept one complete column-width group offline and merge. Undoing one retains the peer's acceptance. Undoing the final remaining decision restores the same complete pending group. Redoing either decision accepts every table/grid/cell property-change member. Zero pending native sites remain, exact widths stay 3500/2500, peer cell text survives, and saved DOCX has the same result. No unrelated history entry appears.

**Check:** Keep both new concurrent table cases enabled. Extend to both actor/client orders, reconnect receivers, each actor's redo, Original/Proposed widths and native membership, peer text, deletion, geometry, and saved convergence. Preserve exact original expectations. Do not equate accepted widths with an accepted group while native property-change snapshots remain.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, private working tree. R49 uses independent removal-flag ownership through the complete review scope. Each original native frame also owns its independent flag intent rows. The pinned Yjs UndoManager adds the intent root before stepping that frame. Untracked normalization restores live removal ownership before native observers publish canonical state. No provider schema or field name, evaluated implementation, native core edit, hidden revision, action replay, or alternate save pipeline is used. Both original failing table cases retain their exact zero-pending-site assertions and pass. Twenty-two real-session cases cover both client orders, reconnect receivers, both actors' undo/redo orders, final remaining undo, all six original Word property snapshots, exact Original/Proposed widths and cell-property absence, peer text, ordinary typing/proposal history, row deletion, remote geometry, room/DOCX reopening, and reversed normalization/intent delivery. R50 closes complete local capacity preflight. Its original boundary regression and exact target/intent tests pass without weakening no-mutation assertions. Sixteen additional real-session capacity cases verify unchanged native state, vectors, peers, prior redo, exact-budget acceptance, history, geometry, and reopening. Eight payload cases preserve independent payload state and neighbors on refusal, and admit metadata-only groups at full capacity. The live capability is version 3. Owned-text and flag ledger protocols stay at version 1. Full final command `bun test --timeout 30000 --preload ../../scripts/test/preload-dom.ts ./src` passes **1626 cases, zero failures, 46434 assertions, 141 files**; `/tmp/codex-r50-retained-private-final.log`. Typecheck passes 1785 files. Build verifies 151 private source modules, one external core, and a React-free main. Built main/collaboration entries pass actual native peer acceptance, final undo, complete redo, saved convergence, and capacity refusal; `/tmp/codex-r50-retained-built-smoke.log`. Scoped formatting, root lint, compatibility, and source audit pass. Compatibility passes 105 decisions with the earlier one-field native format change unchanged. All 196 audited source hashes and reference HEAD `89d7a134` remain unchanged. The isolated app exactly matches verified N21 backup tree `21db89677bae13059183d904892f5e364854f22b`; all five archive hashes match. No published package, manifest, lock, native core, provider, or evaluated implementation changes. The test transport uses the EigenPal Pro License. R15 retains final acceptance/runtime refresh; R16 retains group rejection. See [scope lifecycle and capacity admission](evidence/review-history-scope.md).

## R50

- [x] Preflight review ownership capacity before native group mutation

**Depends:** [R14](05-suggestions.md#r14), [R48](05-suggestions.md#r48), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/collaboration/review-flags.ts`; `fork:packages/sciobot/src/review/group-history.ts`; `fork:packages/sciobot/src/review/accept-group.ts`; `fork:packages/sciobot/src/collaboration/review-flag-capacity.test.ts`; `fork:packages/sciobot/src/review/group-removal-budget.ts`; `fork:packages/sciobot/src/review/group-removal-budget.test.ts`; `fork:packages/sciobot/src/review/group-capacity-collaboration.test.ts`; `fork:packages/sciobot/src/review/group-payload-capacity-collaboration.test.ts`; `fork:packages/sciobot/test/support/review-capacity.ts`

**Read:** [S07](sources.md#s07), [S08](sources.md#s08)

**Change:** Validate and reserve the complete bounded ownership write budget before a native group command mutates canonical or shared state. Keep ownership record limits meaningful. Inspect the existing pure canonical previews and native transaction seams. Use the smallest application-owned preflight or generic Apache capability. Do not recover by hiding revisions, weakening bounds, dropping ownership records, or creating another history entry.

**Close when:** A room with one remaining target slot refuses a two-flag review command before the first native flag changes. Ownership target/intent rows, native XML, canonical revision, state vector, peer state, and user history stay unchanged. Complete commands within their supported budget still preserve R49 native acceptance/history behavior.

**Check:** Keep the enabled capacity fixture and all expected results unchanged. Supply its required complete count when adapting the private API. Add exact-boundary and intent-capacity cases. Verify refusal before writes through the actual native group path with two real sessions. Retain successful multi-member, peer history, reconnect, geometry, deletion, and saved Word cases. Document any admission budget and rollout limits.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, private working tree. The command reuses its pure native operation and payload/package previews. Each stage counts removed canonical node IDs, including snapshot descendants, opaque metadata, and transient nodes. The live scope reserves the complete worst-case target/retained-intent budget before the native mutation callback. Nested reservations hold parent headroom. Duplicate flag writes consume one intent per command. Invalid counts, exhausted capacity, and incomplete Yjs updates refuse before the callback. Unused reservations release on return or exception. Metadata-only decisions with zero removals still work at full capacity. The local capability advances to version 3; both earlier capabilities refuse. Native/shared formats, journals, public document APIs, published packages, manifests, locks, and provider/evaluated implementation stay unchanged. The original countercase keeps its room, two writes, and all no-mutation assertions. Its call supplies the private API's required two-flag count. It now passes. Nine local capacity cases prove exact target/intent boundaries, untouched vectors/history, supported complete commands, duplicate writes, nested admission, release, incomplete-update recovery, and intent identities retained for another actor's redo. Sixteen real-session cases cover both client orders and reconnect receivers, actual native multi-member refusal, unchanged Word XML/revisions/vectors/peer state, retained prior redo and typing undo, exact-budget acceptance, history, geometry, room snapshots, and DOCX reopening. Eight payload cases prove complete pure previews preserve payloads and neighboring groups on refusal, while metadata-only groups still succeed at a full ledger. Two pure budget cases verify body/header wrapper removal, descendants, part removal, transient stages, and source immutability. Focused checks pass 68 cases and 13182 assertions across seven files before the final incomplete-update case; `/tmp/codex-r50-final-focused.log`. The complete final command `bun test --timeout 30000 --preload ../../scripts/test/preload-dom.ts ./src` passes **1626 cases, zero failures, 46434 assertions, 141 files**; `/tmp/codex-r50-retained-private-final.log`. Retained-intent and native boundary checks pass 25 cases and 803 assertions; `/tmp/codex-r50-retained-intents.log`. Typecheck passes 1785 files. Build verifies 151 private modules, one external core, and a React-free main. Built main/collaboration entries pass actual real-session native peer acceptance, final undo, complete redo, saved convergence, and capacity refusal; `/tmp/codex-r50-retained-built-smoke.log`. Scoped formatting, root lint, and compatibility pass. Compatibility records 105 decisions with the earlier one-field native format change unchanged; `.collaboration/changes/sciobot-review-capacity-preflight.json`. Native tests use the EigenPal Pro License through the test-only fixture. This admission check uses settled local state and worst-case headroom; host room admission and performance gate V05 remain separate. R49 uses this verified prerequisite. R15 retains final runtime refresh. The app stays on verified N21. See [scope lifecycle and capacity admission](evidence/review-history-scope.md).

## R51

- [x] Resolve concurrent native table geometry with one property set

**Depends:** [R14](05-suggestions.md#r14), [R15](05-suggestions.md#r15), [R16](05-suggestions.md#r16), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/store/revision-table-rebuild.ts`; `fork:packages/core/src/store/store/tree-op-revisions.ts`; `fork:packages/sciobot/src/collaboration/review-fields.ts`; `fork:packages/sciobot/src/collaboration/review-sequences.ts`; `fork:packages/sciobot/src/review/group-field-budget.ts`; `fork:packages/sciobot/src/review/native-table-property-history-collaboration.test.ts`

**Read:** [S08](sources.md#s08), [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Resolve the same native table decision concurrently without creating duplicate property elements or containers. Cover generated merges and imported cell deletions with missing current geometry. Preserve unknown source properties, both actors' decisions, and complete pending-state undo. Keep canonical OOXML authoritative and evaluation dependencies test-only.

**Close when:** Both peers contain exactly the independent expected native Word property set and grid after acceptance and each history step. One actor's undo retains the other decision. Both actors' undo restores the complete original pending group and geometry. Complete redo, peer text, remote geometry, and room/DOCX reopening pass. Imported owners with and without existing `tcPr` remain supported. No serializer filter, hidden revision repair, or weakened structural oracle replaces the fix.

**Check:** Run the permanent two-session regression for generated merges and imported cell deletions, both client orders, and both reconnect receivers. Retain strict R17 native fingerprints. Compare the actual shared Word parts with independent native operations. Run native core regressions, private group/history/capacity checks, appropriate type/build/lint/format and API/compatibility gates. If public runtime inputs change, rebuild every affected artifact and verify the installed single-core application.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Apache core reuses source marker, property-container, and column identities during concurrent geometry decisions. Transferred properties detach through replayable canonical child-list primitives. Original source paragraphs use the existing R36 retirement format. Unknown properties and original child order remain intact. Private history 4 owns opaque scalar values and individual sequence presence through pinned MIT Yjs item identities. Complete review previews reserve bounded capacity before writes. Initial existence, restored lineage, neighbor anchors, and local presence transitions preserve repeated undo/redo and later ordinary ordering. Private flag protocol 2 refuses legacy and mixed rooms. Native published format fields remain unchanged. No provider implementation, serializer filter, or document snapshot replay is added.

The permanent 72-case table matrix covers generated/imported horizontal and vertical geometry, existing/missing containers, compacted/retained columns, nested controls, both client orders, both reconnect receivers, and both first-undo actors. Two complete history cycles match independent native Word fingerprints. Peer text, native sites, remote geometry, and room/DOCX reopening pass. The unchanged property/proposal fixtures join this matrix in **104 passing native cases**; `/tmp/codex-r51-preserve-later-order-native.log`. Canonical identity/fidelity/replay checks pass **118 cases, 695 assertions**; `/tmp/codex-r51-transferred-slot-core-final.log`. Ownership/budget checks pass **40 cases, 476 assertions**. The stable complete private suite passes **1874 cases, zero failures, zero skipped, 156 files** through the repository's bounded process-per-file runner; `/tmp/codex-r51-stable-full-private.log`. Core/private typechecks pass 1471/1809 files. Private build verifies 158 modules, external core, and a React-free main. Root lint, formatting, API extraction, compatibility, source coverage, and backlog checks pass. Core API extraction checks 15 entries with zero errors and reviewed baseline warnings. Compatibility records 108 decisions; the earlier one-field native change remains unchanged. Decision: `.collaboration/changes/sciobot-native-table-review-identities.json`; changeset: `.changeset/native-table-review-identities.md`.

Runtime backup `refs/backup/codex-r51-runtime` is commit `415bf8cf0cb86c4cd0385cc2210802e913bbebe6`, tree `e07477bd935371c4f58775dccbd6737782013ac1`. All four published inputs match immutable public build `8eff3d6570afca254a21a3235cb3a5ad940236e2`. Each published package was rebuilt there. The standard core build retains the exactly-one HarfBuzz patch and scoped CSS guards. Existing CSS scans use verified original provider class sources through a read-only link; no provider JavaScript is copied or imported. Build provenance stays in `fork-r51/SCIOBOT_R51_STYLE_INPUTS.json` and `fork-r51-presence/SCIOBOT_R51_BUILD_SOURCE.json` under `/home/lofcz/.cache/sciobot-docx-native/`.

The isolated app installs `docx-editor.dev-core-9c1f2cca7b55.tgz` and `sciobot-docx-f7dacf6b1077.tgz`. All five archive hashes, frozen install, direct SDK dependency, four published overrides, lock, source record, and one physical core/Yjs pass. Installed runtime checks pass **126 native/capacity/bulk cases**, plus **12 bulk payload cases**; `/tmp/codex-r51-installed-native-full.log`, `/tmp/codex-r51-installed-bulk-payload.log`. Application checks pass **53 units**, **16 browser cases**, frontend types across 6640 files, harness types across 2206 files, and production build. Browser tests use owned port 8121; its verified process group is stopped afterward. App backup `refs/backup/codex-docx-app-r51-final` is commit `4ffdd5a7e757eefd754196068398603dcb1723b6`, tree `52837b08bde5dcaf0e1439e30836638ec0027cd4`. Only package/lock/source records and the two changed archives differ from R16. Both real indexes and HEADs remain unchanged. All 198 reference application source hashes remain unchanged. See [private history upgrade](evidence/review-history-scope.md#upgrade-private-history-scopes). History-3 rooms require stopped-writer DOCX export and fresh-room bootstrap. Host admission and V05 performance remain separate. R17 stays open for its own acceptance review.

## R52

- [x] Remove a rejected empty paragraph before a block content control

**Depends:** [R05](05-suggestions.md#r05), [R36](05-suggestions.md#r36), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/store/paragraph-mark-merge.ts`; `fork:packages/core/src/store/__tests__/retired-paragraph.test.ts`; `fork:packages/sciobot/src/review/leading-control-paragraph-collaboration.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Resolve a complete inserted paragraph before a block content control without leaving an extra empty Word paragraph. Reuse native retirement when an actor is attached. Preserve another required paragraph, section properties, opaque markers, and unseen peer text. Keep independent pending properties on their current paragraph. Keep canonical authority and provider implementation unchanged.

**Close when:** Rejecting the inserted first paragraph restores exact original Word output before the control. The same native identity recovers unseen peer text. Both actors' history, owner deletion, reconnect, remote geometry, and DOCX/room reopening preserve supported content. Sole required paragraphs and paragraphs after terminal tables remain available. Independent pending paragraph property revisions retain their original owner in pure and shared resolution.

**Check:** Keep R24's strict exported-Word regression enabled. Add focused pure/actor-bound native cases and real-session cases for both client orders and reconnect receivers. Run native regression, canonical journal, type/build/API/compatibility checks. Rebuild affected public artifacts and verify the installed application before closure.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, guarded runtime `refs/backup/codex-r52-guard-runtime`, commit `92ec6f1e573fb7ef256b09f8b4889816c5c2f249`. Native empty-mark cleanup removes the extra paragraph before a block control. It retains required paragraphs, terminal-table paragraphs, sections, opaque markers, independent pending properties, and their current owners. Actor-bound cleanup reuses existing R36 retirement. Forty-eight focused native cases, 25 pure-planner cases, and 20 real-session cases cover exact independent Word output, both client orders and reconnect receivers, concurrent decisions, two history cycles, unseen peer text, owner deletion, remote geometry, and DOCX/room reopening. The first candidate's separate foreign-property failure remains recorded; its archives were never installed. Guarded source passes 2052 private cases across 167 files and 931 native/journal cases across 31 files. Core/private types pass 1471/1827 files. Immutable public/private builds, 57 exact API snapshots, source formatting, root lint, 198 source hashes, and compatibility pass. Compatibility has 113 decisions and the earlier one-field change unchanged; `.collaboration/changes/native-empty-paragraph-before-control.json`. Changeset: `.changeset/empty-paragraph-before-control.md`. The isolated app installs rebuilt core `84eceaf9ba22` and adapter `25fc9fcca74e`. Its actual installed runtime passes 178 cases across six files. App units pass 53 cases; app/harness types pass 6646/2212 files; production build, frozen install, archive/source/lock/line-ending checks, and one-core/one-Yjs checks pass. Both browser files pass all 16 cases with unchanged deadlines. App backup `refs/backup/codex-docx-app-r52-final`, commit `aa130d51d4eb69cc466800effe5869318fa480da`, preserves exactly five intended artifact/manifest changes. Native provider implementation, native shared formats, and public APIs stay unchanged. Desktop Word V06 remains separate. See [implementation notes](evidence/implementation-notes.md).

## R53

- [x] Plan dependent native operations inside one transaction

**Depends:** [R05](05-suggestions.md#r05), [R52](05-suggestions.md#r52), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/store/tree-store.ts`; `fork:packages/core/src/binding/tree-session-contract.ts`; `fork:packages/core/src/binding/tree-session-apply.ts`; `fork:packages/core/src/binding/__tests__/tree-session-stages.test.ts`; `fork:packages/sciobot/src/review/staged-review-operations-collaboration.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Let a native session transaction read its intermediate working package and stage later native operations. Keep ordinary canonical operations, operation gates, actor attribution, package cascades, journal publication, and history authoritative. Do not install a precomputed package snapshot.

**Close when:** A later stage plans against the result of earlier operations. One publication and undo unit cover every stage and related metadata. Failed native operations or thrown stages preserve content, references, revisions, metadata, subscribers, and prior redo. Stages retain content locks, document protection, and package cleanup. No-op stages publish nothing. Two real sessions preserve peer edits, history, remote geometry, reconnect, and DOCX/room reopening.

**Check:** Add focused native transaction tests and real-session cases. Run relevant native/journal regressions, types, build/API, formatting/lint, and compatibility checks. Rebuild affected public artifacts and verify the installed application before closure.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** R05, R52, and V01 are closed. `TreeTransactionContext.currentPackage()` and synchronous `TreeApplyOptions.stage` expose intermediate native results before one publication. Ordinary native operations retain protection, actor allocation, package cascades, journals, and history. Failed operations, thrown stages, async stages, and caught lifecycle attempts preserve the original transaction state. Retained contexts refuse later reads and writes. No-op stages preserve package identity and prior redo. Primary and named-part note cleanup inspect intermediate parts. Twenty-seven focused native cases include metadata, body/header operations, save/reopen, locks, document protection, allocation, rollback, and one undo unit. Twelve real-session cases cover both client orders and reconnect receivers, concurrent peer text, both histories, repeated undo/redo, deletion, remote geometry, and DOCX/room reopening. Focused source checks pass 39 cases and 587 assertions. Broader native/journal checks pass 617 cases and 2659 assertions across 25 files. The complete private suite passes 2064 cases across 168 files with unchanged deadlines and memory gates. Durable logs and terminal proofs remain in `/home/lofcz/.cache/sciobot-docx-native/r53-work`.

Immutable runtime `refs/backup/codex-r53-api-runtime` is commit `d3606be25437238394ab4f38387f00a29f499228`. All five packages rebuild from its exact source. The binding entry explicitly exports the stage context type; all 15 core API entries pass extraction and checking. Actual installed bundles pass 183 cases across six complete files. Application checks pass 53 units, frontend/harness types across 6646/2212 files, production build, frozen installation, and all archive bytes. The browser harness passes 16 cases and 169 assertions; its owned server stops afterward. Core/private types pass 1473/1830 files. Formatting, root lint, compatibility, and the 198-source-hash audit pass. Compatibility records 114 decisions with the earlier field change unchanged. `.changeset/native-dependent-operation-stages.md` declares a core minor release. `.collaboration/changes/native-dependent-operation-stages.json` records no shared format change. App backup `refs/backup/codex-docx-app-r53-final` preserves commit `d788c8b63e62e9165f06783155a985363cd4302c`, tree `eb06caeb514c0e1340c179890dab988b6c409098`. Both branch HEADs and real indexes remain unchanged. No provider implementation changes occur. Test transport uses the EigenPal Pro License and stays outside archives. R25 retains the validated application replacement command. See [implementation notes](evidence/implementation-notes.md).

**Tracked inline ownership checkpoint:** Runtime `44f01237` couples complete inline payloads to exact native review receipts. Proposal schema 3 preserves native ID reservations after decisions and tombstoning; older built readers refuse all three states. Local and shared decisions, owned inverses, history, source run boundaries, and source hyperlinks pass the recorded targeted checks. See [evidence and remaining scope](evidence/tracked-inline-ownership.md). D30 and R19 remain open.

**Tracked removal coordinator checkpoint:** Runtime `3b388abc` applies removal-only AI targets through the actual native transaction writer and captures schema 3 payload ownership. Eleven built-SDK cases pass, including locked controls, both decisions, exact rejection, history, reopening, and repeated AI waves. Four shared coordinator cases pass both client orders and peer histories. Dependent mixed coordinates, births, image resources, OMML, and final application verification remain open. See [checkpoint evidence](evidence/tracked-inline-ownership.md#ai-coordinator-removal-checkpoint). D30 and R19 remain unchecked.

**Mixed tracked removal checkpoint:** Runtime `7e93de92` reconciles text and format coordinates around native atom deletions. Native source scopes restore exact original run boundaries. Independent decisions and shared histories pass the recorded checks. Guard `cb97baea` rejects proposed grapheme splits before publication; all 22 built-SDK cases pass. See [evidence and verification status](evidence/mixed-inline-review.md). D30 and R19 remain unchecked.

**Tracked retained birth checkpoint:** Runtime `a7a2f968` stages inline math, mentions, and source births through native review with complete payload ownership. Nineteen local cases and four shared coordinator cases pass. All 43 built-SDK cases pass, including mixed removal and independent decisions. A broader owned-presence run exposed an inherited offline suffix placement failure, reproduced before the birth change. See [evidence and remaining work](evidence/tracked-retained-inline-births.md). D30 and R19 remain open.
