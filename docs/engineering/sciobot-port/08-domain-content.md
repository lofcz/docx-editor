# Application content and custom renderers

Reuse application services and content UI. Remove Plate-specific wrappers from each renderer separately.

Read the [dispatch and completion rules](../../../SCIOBOT_PORT_TODO.md) before claiming a leaf.

Each named test or artifact is a deliverable. These specifications do not claim those checks already exist.

## X01

- [x] Render and edit lesson title fields

**Depends:** [B05](02-react-blocks.md#b05), [B06](02-react-blocks.md#b06), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01), [B27](02-react-blocks.md#b27)

**Primary files:** `app:src/components/workspace/docx/blocks/lesson-title.tsx`; `sibling browser test`

**Read:** [S09](sources.md#s09)

**Change:** Adapt title content, grade, lessonMinutes, and organizationalMinutes to registered blocks. Keep the title body native.

**Close when:** Edit each field and save/reopen it. Title text remains natively selectable and each field action has one undo entry.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, app branch `claude/docx-native-editor` (continuing Claude). **Evidence:** Application worktree `sciobot-next-wts/docx-native`, branch `claude/docx-native-editor`, baseline `f32fefc83` plus inherited local work. Vendor source `07701aae993b87f83532e87ee87413fc826b3960`, pinned in `vendor/docx-native/SOURCE.json`. `blocks/lesson-title.tsx` keeps title text native and binds grade, lessonMinutes, and organizationalMinutes to canonical payload fields. Each control action has a separate undo entry; clear removes the attribute. The 3 lesson-title cases pass: real popover/preset/numeric input, native word replacement, undo/redo per field, and saved DOCX reopen with every value. Two real sessions in clients 73/41 and 41/73 verify online field sync, offline grade edits beside peer title typing, disjoint-field merge, one register winner for conflicting grades, peer-preserving history, equal saved digests, and reopen. B27 is closed and its refreshed vendor snapshot passes the formerly blocked scenarios. Combined Chromium command `bun test --timeout 600000 scripts/docx-native/harness/placeholders.test.ts scripts/docx-native/harness/lesson-title.test.ts scripts/docx-native/harness/media.test.ts`: 9 pass, 0 fail, 74 assertions; `/tmp/codex-docx-app-harness.log`. Harness/application typecheck (`bun check -p scripts/docx-native/tsconfig.json`) passes 504 files. Scoped application lint passes with one existing Fast Refresh warning in `media.tsx`.

## X02

- [x] Render preparation recap containers

**Depends:** [B05](02-react-blocks.md#b05), [B06](02-react-blocks.md#b06), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/recap.tsx`; `sibling browser test`

**Read:** [S09](sources.md#s09), [S11](sources.md#s11)

**Change:** Adapt preparation, goals, materials, and adaptations wrappers. Preserve label splitting and existing list normalization.

**Close when:** One fixture per wrapper retains nested text through typing and export. Adaptation labels match the existing renderer.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, app branch `claude/docx-native-editor`. **Evidence:** App commit `10bfbcbec` (worktree `sciobot-next-wts/docx-native`; not pushed). Engine tarballs from fork `refs/backup/codex-wip-2026-10-20`. `src/components/workspace/docx/blocks/recap.tsx` registers `preparation_recap`, `recap_goals`, `recap_materials`, and `recap_adaptations` as `nativeContent` containers. Every row stays a native paragraph in the slot. Headers cover the generated `scio:label` line and use the existing `LL.common.recap.*` labels. Without a label line, they sit above the block. The outer recap shows the existing summary: filled goal rows plus lesson minutes, meaning activity durations plus title `organizationalMinutes`, as `sumLessonDuration` computes. List normalization follows the Plate renderer: rows are bare paragraphs, goal numbers come from the position of non-empty rows, and empty rows get no number. Row badges come from painted paragraph geometry (`blocks/nativeRows.ts` `useRowBoxes`, a MutationObserver over the slot). Adaptation rows use the existing `splitAdaptationRow`, so labels match the static and print renderers: `SPU`, `Asistent`, and none when the colon falls after the 24-character limit. Material rows add through a native `splitParagraph` at the last row's end, which also focuses the new row. They remove through `deleteBlock`; the last row is cleared instead of deleted. Each action is its own undo unit. `bun test --timeout 600000 scripts/docx-native/harness/recap.test.ts` passes 3 Chromium cases and 38 assertions. (1) Render check: labels, numbers 1-3, adaptation labels, and summary `3 cíle · 20 min`. Real typing in each wrapper, Enter in goals (numbers 1-4, summary `4 cíle`), and a relabelled adaptation row (`Suplování`). Material add/remove. The working-view Markdown export contains every typed row inside its `<recap_goals>`, `<recap_materials>`, and `<recap_adaptations>` tags nested in `<preparation_recap>`. Block IDs, control IDs, and parents are unchanged. One undo restores a removed row. (2,3) Two sessions in both client orders (73/41, 41/73): offline concurrent typing in goals and adaptations, then reconnect. Peer row removal while the other session types in a sibling row. Peer-preserving undo/redo. Equal saved digests. Reopened DOCX keeps adaptation labels and exports the nested text. The X04 phases suite still passes 7/7. The commit also fixes a latent X04 bug: React Compiler cached canonical session reads, so derived values such as phase minutes and the recap summary went stale. Canonical reads now take the host `revision`. Harness typecheck, oxlint, and oxfmt pass. Limits: the Plate accordion collapse and per-entity open state are not ported (X17/U07). The empty-materials add card and the material placeholder text belong to U08. Material actions refuse in suggesting mode and need a tracked review transaction. Legacy `1. ` goal prefixes are a conversion concern (I12).

## X03

- [x] Render curriculum recap blocks

**Depends:** [X02](08-domain-content.md#x02), [B20](02-react-blocks.md#b20), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/curriculum.tsx`; `sibling browser test`

**Read:** [S09](sources.md#s09), [S11](sources.md#s11)

**Change:** Adapt curriculum framework and school/catalog source fields. Reuse existing competence displays and application navigation.

**Close when:** A school curriculum fixture retains framework and source across save/reopen. Competence actions receive the correct identifiers.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, application branch `claude/docx-native-cutover`. **Evidence:** Application commits `b81d20862` and `a7b7becdf`. `blocks/curriculum.tsx` keeps framework and source in the payload and competence rows as native paragraphs (`code: text`, or `legacy-id-N: text` for catalogue IDs without a code). Catalogue rows use the existing `ResolvedCurriculumCompetenceRow` and open the competence detail through `useOpenCurriculumDetail`; school (ŠVP) rows use `CurriculumCompetenceRow` without a catalogue action, as Plate does. Catalogue recaps keep the existing `CurriculumAddPopover`; an added row and a removed row are each one native undo step (`blocks/curriculum-rows.ts`). Two defects surfaced and are fixed: block roots had no React Query client, so every catalogue row threw `No QueryClient set` and its recap did not render (block roots now share the application query client); and the framework label sat above the block, where it covered the last row of a neighbouring recap (it is now in the page margin). Chromium `curriculum.test.ts` passes 3 cases: a coded row opens the detail with its code, framework, and label, a `legacy-id-42` row with ID 42, and a school row has no action; and two sessions in both client orders add and remove rows, converge, undo, and reopen with `framework` and `source` unchanged and equal saved digests. The catalogue search itself needs the backend and is not exercised; the test calls the popover's add action directly. Difference: the RVP/ŠVP segmented toggle is not ported; both recaps show.

## X04

- [x] Render phases and phase recaps

**Depends:** [B05](02-react-blocks.md#b05), [B06](02-react-blocks.md#b06), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/phase.tsx`; `sibling browser test`

**Read:** [S09](sources.md#s09)

**Change:** Adapt phase and recap wrappers using canonical phaseType values. Keep labels localized and child text native.

**Close when:** All four phase types render correct labels. Editing a nested paragraph preserves its phaseType and wrapper identity.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, app branch `claude/docx-native-editor`. **Evidence:** App commits `7ca20f0b4` and `81b7f8fc8` (worktree `sciobot-next-wts/docx-native`; not pushed). Engine tarballs come from fork `refs/backup/codex-wip-2026-10-20`. `src/components/workspace/docx/blocks/phase.tsx` registers `lesson_phase` and `recap_phase` as `nativeContent` containers in `WORKSPACE_BLOCKS`. Every child paragraph and activity stays in the native slot. `PhaseBlock` draws the existing `PHASE_ACCENT` rail beside each fragment. On the first fragment it draws a header with the `getPhaseLabel(phaseType, locale)` label and the summed child-activity minutes, read from the canonical block index. That header covers the generated `scio:label` Word line. `RecapPhaseBlock` shows the localized label pill beside the native recap description. Labels never come from stored text. `bun test --timeout 600000 scripts/docx-native/harness/phases.test.ts` passes 7 cases in Chromium. All four `phaseType` values render correct labels in phases and recaps for cs, en, sk, and pl, with correct per-phase minutes. Real typing in nested phase, activity, and recap paragraphs, plus Enter splits, keeps the persistent IDs, canonical control IDs, parents, and payloads. The same header DOM element survives (wrapper identity), and undo restores the text. Two-session cases in both client orders (73/41, 41/73) cover: offline concurrent nested typing in different phase children, reconnect, peer deletion of a nested paragraph while the other session types in the same phase, peer-preserving undo/redo, unchanged wrapper IDs in both sessions, equal saved digests, and labelled reopen. The full harness suite (`bun run test:docx-native`) passes 15 of 15 twice in a row under load average 40-60. Harness waits now use `clickNative` (focus-confirmed clicks) and a 45 s `until` default. Limits: the Plate accordion collapse and phase info tooltip are not ported. Collapse belongs to X17/U07; remove-suggestion collapse belongs to R/U review leaves.

## X05

- [x] Render activity name and duration controls

**Depends:** [B05](02-react-blocks.md#b05), [B06](02-react-blocks.md#b06), [R12](05-suggestions.md#r12), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/activity.tsx`; `sibling browser test`

**Read:** [S09](sources.md#s09)

**Change:** Adapt activity name, duration, and resource fields. Route field changes through the canonical payload transaction.

**Close when:** Editing name or duration creates one property change when review is active. Native activity content remains unchanged.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, `codex/react-block-proof`, isolated application worktree at `116bde74e77d66b9f4300a4f836ad055e6fa7029`.

**Evidence:** The activity renderer supplies localized name, duration, and resource controls. Duration retains the Plate zero-through-120 bounds. Direct edits use native payload transactions. Suggested fields use `recomposeOwnedProperties`, introduced in fork commit `72464e44b507fb3fb10a8c94147a492cbe8fa30b`. Explicit field intent retains the native owner and body. AI container identity rules remain unchanged. The stored executor distinguishes manual field proposals. Exact returned receipts and content proofs authorize consecutive replacements. Invalid reads refuse writes. Resource-only edits retain the generated label. Reverting fields removes the owned proposal and retains history.

The source batch passes **53 cases, zero failures, 875 assertions**, covering default AI identity refusal, explicit resource waves, native body fingerprints, both decisions, history, reopen, offline peer edits, both client orders, and saved convergence. Command: `bun test --timeout 60000 --preload ./scripts/test/preload-dom.ts packages/sciobot/src/review/recompose-properties.test.ts packages/sciobot/src/diff/domain-properties.test.ts packages/sciobot/src/review/recompose-payloads.test.ts packages/sciobot/src/review/recompose-payloads-collaboration.test.ts packages/sciobot/src/review/property-projection.test.ts`.

The installed application activity batch passes **29 cases, zero failures, 733 assertions**: `DOCX_NATIVE_SNAPSHOT=/home/lofcz/.cache/sciobot-docx-native/port-current-integration bun test --timeout 120000 scripts/docx-native/harness/activity.test.ts`. Browser actions verify each field, one pending group, unchanged canonical body and owners, consecutive mixed fields, no-op/reset, both decisions and histories, offline typing/deletion, both client orders, pending/resolved reopen, saved digests, and caret geometry. The associated direct-edit and export regression files pass their nine cases. Application and harness typechecks pass 6,832 and 2,402 files. All five archive hashes match the manifest. All consumers resolve one core. Frozen installation passes without changes. The fork commit passes repository hooks. Scoped formatting, lint, and diff checks pass.

## X06

- [x] Derive phase and preparation duration totals

**Depends:** [X01](08-domain-content.md#x01), [X04](08-domain-content.md#x04), [X05](08-domain-content.md#x05)

**Primary files:** `app:src/lib/docx/phase-durations.ts`; `sibling phase-durations.test.ts`

**Read:** [S09](sources.md#s09), [S15](sources.md#s15)

**Change:** Derive totals from activity payloads with the existing phase rules. Do not save duplicate derived duration values.

**Close when:** Changing one activity updates its phase and recap totals. Undo returns all displayed totals without a second write.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, isolated application worktree, commit `a30e288500c6442f68c191e0cbf74322250b3f4e`.

**Evidence:** `src/lib/docx/phase-durations.ts` derives phase totals, totals by phase type, organizational minutes, and lesson totals. It follows the existing root-phase/direct-activity rules. Recap, phase, and recap-phase renderers share this read-only calculation. Pending property values participate. Immutable package identity caches each result and invalidates it after local edits, remote edits, decisions, undo, and redo. No derived duration is persisted. Four focused unit cases verify exact totals, repeated phase types, preparation time, root-only ownership, orphan/nested exclusions, first-title rules, and empty/nonnumeric values.

The completed combined browser command passes **45 cases, zero failures, 1,045 assertions**: `DOCX_NATIVE_SNAPSHOT=/home/lofcz/.cache/sciobot-docx-native/port-current-integration bun test --timeout 120000 scripts/docx-native/harness/duration-totals.test.ts scripts/docx-native/harness/phases.test.ts scripts/docx-native/harness/recap.test.ts scripts/docx-native/harness/activity.test.ts`. Six total-specific cases cover direct edits and accepted/rejected suggestions in both client orders. Native browser input changes one activity. Phase totals, phase-recap totals, and the lesson recap update together. One undo restores every display. Offline peer typing and deletion remain intact. Peer deletion history, matching saved digests, pending/final reopening, and caret geometry pass. Existing activity, phase, localized label, and recap cases also pass. The same batch verifies the activity toolbar placement fix: generated labels carry the controls, and preceding native paragraphs remain clickable. Application and harness typechecks pass 6,833 and 2,403 files. Focused lint, formatting, and diff checks pass.

## X07

- [x] Port preparation retiming actions

**Depends:** [X06](08-domain-content.md#x06), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/retime-document.ts`; `sibling retime-document.test.ts`

**Read:** [S15](sources.md#s15)

**Change:** Reuse retiming rules and write all changed activity durations in one canonical transaction.

**Close when:** Existing retiming examples produce the same minutes. One undo restores every affected duration.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, isolated application worktree, commits `59f372e90ad4b881cc5e0e4918e3408f01417399` and `c2df9c9c3de2752ac5961476a65e866f86f6b2d4`.

**Evidence:** `src/lib/docx/retime-document.ts` connects the existing title duration control to one native transaction. `src/components/lesson-plan/retime-values.ts` shares allocation rules with the Plate adapter. The reference examples retain exact whole-minute allocations and the one-minute activity minimum. Short lessons merge root phases into Preparation. Their activity controls, introductory paragraphs, and recap descriptions move through native tree primitives. The action updates title metadata and generated Word labels in the same history unit. Unrelated attributes, formatted text, bookmarks, hyperlinks, native identities, and unsupported future payload records remain intact. The action uses direct editing mode and refuses conflicting pending field proposals. It never disables tracking. The detached preview does not replace the live package.

The browser regression exposed incorrect sibling order after native undo. Fork commits `4151d7b1` and `929e0ea4a9bf68f3a4c9a7c133e55bc76b76d906` fix ordered field history and pre-bootstrap admission. Private field protocol 3 retains same-transaction sibling anchors and original Yjs slots. Low-level profiles 1 and 2 retain their previous behavior. Old private room profiles refuse the new scope. Upgrade private SDK consumers together and bootstrap a fresh room from saved DOCX, as described in `packages/sciobot/README.md#ordered-field-history`. DOCX payload and review record formats remain unchanged. No published core format changes.

The source batch passes **96 cases, zero failures, 1,547 assertions**: `bun test --timeout 60000 --preload ./scripts/test/preload-dom.ts packages/sciobot/src/collaboration/review-sequences.test.ts packages/sciobot/src/collaboration/review-fields.test.ts packages/sciobot/src/collaboration/review-field-retraction.test.ts packages/sciobot/src/collaboration/review-history.test.ts packages/sciobot/src/review/recompose-properties.test.ts packages/sciobot/src/review/recompose-payloads-collaboration.test.ts packages/sciobot/src/blocks/move.test.ts`. These cover separate removals, peer insertions at each neighboring position, both delivery orders, repeated histories, profile fencing, mutation-free premature admission, and actual native property/recomposition collaboration. A cold-Yjs-state probe also retains the restored order and profile header.

The focused application unit batch passes **17 cases**: `bun x rstest src/lib/docx/retime-document.test.ts src/lib/preparationTiming.test.ts src/lib/docx/phase-durations.test.ts`. The integration boundary batch passes **52 cases, zero failures, 1,887 assertions**, including activity suggestions, direct AI edits, duration totals, title controls, retiming, and text export. Command: `DOCX_NATIVE_SNAPSHOT=/home/lofcz/.cache/sciobot-docx-native/port-current-integration bun test --timeout 120000 scripts/docx-native/harness/retiming.test.ts scripts/docx-native/harness/lesson-title.test.ts scripts/docx-native/harness/activity.test.ts scripts/docx-native/harness/duration-totals.test.ts scripts/docx-native/harness/direct-edits.test.ts scripts/docx-native/harness/text-export.test.ts`.

Final verification against installed fork `929e0ea4a9bf68f3a4c9a7c133e55bc76b76d906` passes **eight browser cases, zero failures, 898 assertions**: `DOCX_NATIVE_SNAPSHOT=/home/lofcz/.cache/sciobot-docx-native/port-current-integration bun test --timeout 120000 scripts/docx-native/harness/retiming.test.ts scripts/docx-native/harness/lesson-title.test.ts`. Browser input selects 20- and 90-minute targets in both client orders. Offline peer typing, resource edits, and deletion survive retiming. One undo restores every duration and the original phase order; redo restores the retimed result. Peer deletion history stays independent. Exact native paragraph fingerprints, control IDs, Word label text, future payload JSON, saved digests, reopen, cursor geometry, and owned-scope health pass. Invalid budgets change no native data. Application and harness typechecks pass 6,835 and 2,406 files. Five archive hashes match the manifest. Every consumer resolves one core. Frozen installation, scoped formatting, lint, diff checks, and required fork hooks pass.

## X08

- [ ] Render and replace lesson visuals

**Depends:** [B05](02-react-blocks.md#b05), [B11](02-react-blocks.md#b11), [D14](04-edit-planner.md#d14), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/lesson-visual.tsx`; `sibling browser test`

**Read:** [S09](sources.md#s09), [S12](sources.md#s12), [S15](sources.md#s15)

**Change:** Reuse visual selection and replacement services. Preserve kind, phase, prompt, pristine file ID, alt text, and displayed variant.

**Close when:** Replace an explainer image and undo it. Export includes the intended image and metadata; entitlement URLs never replace the pristine reference.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X09

- [x] Preserve listening markup and audio references

**Depends:** [B02](02-react-blocks.md#b02), [P05](03-agent-projection.md#p05), [D14](04-edit-planner.md#d14), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/listening.ts`; `app:src/lib/docx/listening.test.ts`; `app:scripts/docx-native/harness/listening-edits.test.ts`

**Read:** [S12](sources.md#s12), [S11](sources.md#s11)

**Change:** Preserve speakers, lines, voices, portraits, seed, audioRef, and exercise data in the payload. Keep the native transcript complete.

**Close when:** Existing listening Markdown fixtures roundtrip. Editing a nearby sentence retains generated audio and speaker references.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Application `c5b7bbf8c` and `b5b8cee742ee9bb85cde4b36b9af685ca5643e77` adapt the existing shared grammar to B02/P05 native payload storage. Both captured listening Markdown fixtures roundtrip with complete audio/hash/seed, speaker voices/portraits, stable keys, transcript text, and unknown data retained. Nearby edits preserve complete payloads and exact native control fingerprints. Both client orders, both decisions, both histories, remote caret geometry, deletion during peer editing, both reconnect receivers, saved convergence, and cold reopening pass. The final units pass 12 cases. Eight installed listening browser cases pass across review/offline commands with 140 assertions. The initial affected atomic/listening batch passes nine cases. Frontend/harness types pass 6,844/2,418 files; scoped checks pass. Complete command and independent oracle details are in [listening evidence](evidence/native-listening-data.md). Runtime stays `4f0029f0b8df49f1aa966c43ce52a7a97b8443b0`. X10 retains playback, hydration, generation, and freshness integration.

## X10

- [x] Render listening exercises with existing services

**Depends:** [X09](08-domain-content.md#x09), [B05](02-react-blocks.md#b05), [B17](02-react-blocks.md#b17), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/listening.tsx`; `app:src/lib/docx/listening-audio.ts`; `app:src/lib/docx/listening-locks.ts`; `app:scripts/docx-native/harness/listening-player.test.ts`

**Read:** [S12](sources.md#s12)

**Change:** Adapt listening playback, generation, and cache hydration. Keep local playback state separate from saved exercise data.

**Close when:** A restored exercise hydrates existing audio without regeneration. Playback and long transcripts remain usable inside paginated content.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Application `a7d38a465` integrates existing playback, cache hydration, and explicit generation with native payload storage. Thirteen distinct player browser cases verify local playback, long transcripts, two-session history/convergence, stale response cancellation, review freshness, locked playback, and cold reopening. The final unit batch passes 31 cases. Frontend/harness types pass 6,849/2,424 files. Commands, fixtures, exact oracles, and limits are in [listening player evidence](evidence/native-listening-player.md). Runtime remains `4f0029f0b8df49f1aa966c43ce52a7a97b8443b0`.

## X11

- [x] Render source citation chips

**Depends:** [B05](02-react-blocks.md#b05), [B20](02-react-blocks.md#b20), [B21](02-react-blocks.md#b21), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/source-chip.tsx`; `app:src/lib/docx/source-citation.ts`; `app:scripts/docx-native/harness/source-citations.test.ts`

**Read:** [S10](sources.md#s10)

**Change:** Adapt source and book-source chips with existing provenance services. Preserve URL, title, author, type, bookId, and trust.

**Close when:** A chip opens the matching provenance view. Save/reopen and atomic AI replacement retain all citation attributes.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Application `e9eeaa81f` installs native runtime `926527dd155f7bd647d538d7ecbcadea2683d4ec` and connects existing source/book provenance. Ten distinct installed browser cases verify complete citation attributes, two AI waves, both decisions, original-view provenance, exact native hit geometry, independent peer history, offline deletion, and save/reopen convergence. The native batch passes 79 cases and 1,094 assertions; fifteen application units and frontend/harness types pass. Native mixed revision text stays visible. Commands, scoped limits, archive hashes, and earlier-reader checks are in [source citation evidence](evidence/native-source-citations.md).

## X12

- [x] Render source excerpt containers

**Depends:** [X11](08-domain-content.md#x11), [B06](02-react-blocks.md#b06), [D12](04-edit-planner.md#d12), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/source-excerpt.tsx`; `sibling browser test`

**Read:** [S10](sources.md#s10)

**Change:** Adapt excerpt framing around native body text. Apply existing citation identity rules independently from descendant edits.

**Close when:** Changing an excerpt sentence preserves its wrapper. Changing an identity field produces the expected structural review group.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, application branch `claude/docx-native-cutover`, fork `00c33e27`. **Evidence:** Application commits `0bbcdef9a` (frame and body suite) and `780e29a73` (identity cases in the harness). `SourceExcerptBlock` frames the native body and keeps provenance controls in the page margin. The four identity cases that failed at the checkpoint (`Only complete native paragraphs without section marks can be removed here`) now pass on the current fork: a changed citation title creates one structural presence proposal that removes the old `source_excerpt` and adds the new one, the stored payload stays unchanged until a decision, and accept or reject leaves one excerpt with the expected title, no pending groups, and the earlier body edits; undo and redo follow. Chromium: `source-excerpts.test.ts` (8 cases: body edits, wrapper identity, offline neighbor deletion, reopen) and `source-excerpt-identity.test.ts` (4 cases: both client orders with accept and reject) pass 12 cases together.

## X13

- [x] Render callout blocks

**Depends:** [B05](02-react-blocks.md#b05), [B06](02-react-blocks.md#b06), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01), [X37](08-domain-content.md#x37)

**Primary files:** `app:src/components/workspace/docx/blocks/callout.tsx`; `sibling browser test`

**Read:** [S16](sources.md#s16)

**Change:** Adapt existing callout attributes and content UI. Keep nested lists and text native.

**Close when:** The callout Markdown fixture survives import, typing, and export with its original attributes.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Runtime `f5bb40c9c703e61abef8cec05021d02eda715f78` and application `fad7d96a2232a048418c4d3069212bc27c1fda28` implement native callouts. Appearance helpers and the emoji picker are reused without Plate runtime imports. Native paragraphs and nested lists retain their slots. The Markdown import fixture preserves variant aliases, icon, explicit background, and opaque metadata through browser typing and export. Three container browser cases cover both client orders, offline peer edits, deletion, both histories, exact owner identities, geometry, native save convergence, and reopening. The combined container/retiming batch passes eight cases and 960 assertions. Native authoring and captured projections pass 47 cases and 214 assertions. Frontend and harness types, scoped formatting/lint, five archive hashes, frozen install, and single-core resolution pass. See [container evidence](evidence/native-container-renderers.md).

## X14

- [x] Render story blocks

**Depends:** [B05](02-react-blocks.md#b05), [B06](02-react-blocks.md#b06), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01), [X37](08-domain-content.md#x37)

**Primary files:** `app:src/components/workspace/docx/blocks/story.tsx`; `sibling browser test`

**Read:** [S16](sources.md#s16)

**Change:** Adapt story metadata and content UI without PlateElement. Keep the story body in a native slot.

**Close when:** A nested story fixture preserves metadata and paragraph editing through save/reopen.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Runtime `f5bb40c9c703e61abef8cec05021d02eda715f78` and application `fad7d96a2232a048418c4d3069212bc27c1fda28` implement localized story chrome around native descendants. The imported fixture nests a callout and a two-level native list inside the story. Narrator and opaque metadata remain exact through editing, concurrent peer deletion, both actors' histories, saved convergence, and reopening. Native owner identities and remote caret geometry remain equal. The combined browser batch passes eight cases and 960 assertions; source authoring/projection checks pass 47 cases and 214 assertions. Frontend/harness types and scoped checks pass. See [container evidence](evidence/native-container-renderers.md).

## X15

- [x] Render solution blocks

**Depends:** [B05](02-react-blocks.md#b05), [B06](02-react-blocks.md#b06), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01), [X37](08-domain-content.md#x37)

**Primary files:** `app:src/components/workspace/docx/blocks/solution.tsx`; `sibling browser test`

**Read:** [S16](sources.md#s16)

**Change:** Adapt solution reveal controls and attributes. Keep reveal state separate from saved answer content.

**Close when:** Reveal and hide retain the answer. Export includes the same solution content as the existing static renderer.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Application `fe157dbd7beb83d80769eff8ccca74c520367909` registers a native solution renderer over runtime `f5bb40c9c703e61abef8cec05021d02eda715f78`. Reveal state stays local to one editor and synchronizes its page fragments. Continuation controls, nested synchronous AI reveal, native selection, and input capture expose answers before editing. Screen-only hiding retains native Word pagination. Print and exports retain full answer content. Exact native state, package revision, history, and saved digest remain unchanged by reveal/hide. The final solution/container browser batch passes eight cases and 137 assertions, including both client orders, independent peer visibility, concurrent edits/deletion, both histories, caret geometry, saved convergence, and collapsed reopening. Frontend/harness types and scoped format/lint pass. No runtime pin, shared format, canonical content model, or public API changes. See [solution evidence](evidence/native-solution-renderer.md).

## X16

- [x] Render editable code blocks

**Depends:** [B05](02-react-blocks.md#b05), [B06](02-react-blocks.md#b06), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/code.tsx`; `sibling browser test`

**Read:** [S13](sources.md#s13)

**Change:** Preserve code text, language, and whitespace. Reuse syntax display while routing edits to canonical native text.

**Close when:** Tab, Enter, paste, and language changes retain exact code text. Word fallback and Markdown export preserve whitespace.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, application branch `claude/docx-native-cutover`, fork `62c7e8ef`. **Evidence:** Application commit `da8533f99`. The `code_block` view (`src/components/workspace/docx/blocks/code.tsx`) adds a language selector, an edit dialog, and a copy button over the native Word fallback, which paints one code-styled line per source line. Saves go through the new fork API `updateAtomicBlock`, which replaces the payload and regenerates the fallback in one undo unit, keeps the block ID, and refuses suggest and view mode. The dialog inserts a tab for Tab and keeps Enter and paste as plain text. Fork `blocks/update-atomic.test.ts` passes 3 cases, including two real collaboration sessions in both receiver orders. Chromium `code-blocks.test.ts` passes 3 cases: Enter, Tab, clipboard paste, language change, exact payload, Word lines, saved `w:tab` and repeated spaces, Markdown export, undo and redo, suggest-mode refusal, and two sessions with peer typing, history, equal saved digests, and equal Markdown. Pages show the code style without syntax colors. See [native code blocks](evidence/native-code-blocks.md).

## X17

- [x] Render collapsible toggle blocks

**Depends:** [B05](02-react-blocks.md#b05), [B06](02-react-blocks.md#b06), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/toggle.tsx`; `sibling browser test`

**Read:** [S13](sources.md#s13)

**Change:** Adapt collapse controls around native descendants. Reveal collapsed ancestors for AI navigation and selection.

**Close when:** An AI target inside a closed toggle becomes visible. Collapse does not delete or duplicate its native content.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, application branch `claude/docx-native-cutover`, fork `c16989ea`. **Evidence:** Plate keeps a toggle as an indented paragraph and its Markdown serializer drops the toggle title, so no Markdown shape existed. The fork registers `toggle` as a plain container (`<toggle>` with native children); its first paragraph is the summary. Fork `projection/toggle.test.ts` plans the tag, projects it back exactly, and plans a body change as one text operation. Application commit `ca138f7b5` adds `blocks/toggle.tsx` and the slash entry. Collapsing hides every native paint element except the summary on screen only; Word export, print, and the store never change. `blocks/collapsible-reveal.ts` (generalized from the solution reveal) opens a toggle or solution when a selection, typing, or a review reveal reaches hidden content, including blocks on pages that are not laid out yet. Chromium `toggles.test.ts` passes 5 cases: collapse with unchanged Markdown, history, and saved digest; summary typing that keeps it closed and Enter that opens it; review navigation into a closed toggle; slash insertion with undo; and two sessions in both client orders with edits in summary and body, local open state, author undo, and equal saved digests. `direct-edits.test.ts` adds an AI wave inside a closed toggle on a later page.

## X18

- [ ] Render table of contents navigation

**Depends:** [N02](07-native-content.md#n02), [B05](02-react-blocks.md#b05), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/toc.tsx`; `sibling browser test`

**Read:** [S13](sources.md#s13)

**Change:** Derive entries from native heading identities. Reuse existing depth options and scroll behavior.

**Close when:** Rename and move a heading. Entries update and navigate to the correct page after repagination.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X19

- [x] Map application columns to native table geometry

**Depends:** [B02](02-react-blocks.md#b02), [B06](02-react-blocks.md#b06), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/blocks/columns-codec.ts`; `sibling columns-codec.test.ts`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Represent column groups as borderless native tables with application metadata. Preserve nested blocks and per-column widths.

**Close when:** Two and three-column fixtures retain nested text and tables through save/reopen. The Word fallback preserves column ordering.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Fork commit `c8fe2c79`. Plate serializes layouts as `<column_group>` with `<column width=…>` children. The Sciobot registration now maps `column_group` and `column` to native container blocks (`projection/block-tags.ts`, `registry/sciobot-parse.ts`, `registry/sciobot-diff.ts`), so each column is a nested content control that holds its native blocks and keeps its `width` and other attributes in the payload. Design decision: the columns are nested containers in reading order, not a borderless table. Word therefore shows the columns one after another, which preserves their order; side-by-side page layout in Word is a recorded follow-up and is not claimed. The same change fixes a defect for every new application block that contains a table: `fragment-seeds.ts` emitted a malformed `w:tblGrid` (an operator-precedence error produced a truncated first `w:gridCol`), so such blocks refused with `parse-error`. `bun test packages/sciobot/src/blocks/columns.test.ts` passes 5 cases: two columns with a nested GFM table and three columns with a nested list keep payload widths, the projected `<column_group>`/`<column width>` tags, identical projection after save and reopen, Word text order, and empty replanning; a live column edit is one undo step with redo and reopen; two real sessions in both receiver orders converge on a column edit beside peer typing in another column, show Bob's remote caret in Alice's layout, keep Bob's text through Alice's undo and redo, converge when Bob deletes text while Alice types, and save equal digests. The registry, projection, diff, and blocks suites (134 files) show only the pre-merge baseline failures (`target-source-map`, `inserted-atoms`). Pre-commit gates pass.

## X20

- [ ] Connect column widths and column controls

**Depends:** [X19](08-domain-content.md#x19), [B05](02-react-blocks.md#b05), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/columns.tsx`; `sibling browser test`

**Read:** [S13](sources.md#s13)

**Change:** Adapt width handles and existing column actions to native table geometry. Keep private column metadata synchronized.

**Close when:** Resize one column and move content between columns. Undo restores widths and content. Native caret geometry follows the resized grid.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X21

- [ ] Render and edit date chips

**Depends:** [B05](02-react-blocks.md#b05), [B20](02-react-blocks.md#b20), [B21](02-react-blocks.md#b21), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/date.tsx`; `sibling browser test`

**Read:** [S13](sources.md#s13)

**Change:** Adapt the date picker and locale display using an inline payload and readable Word fallback.

**Close when:** Choose a date, copy it, and reopen the document. Stored date identity remains independent of the locale display.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X22

- [ ] Render and insert mention chips

**Depends:** [X38](08-domain-content.md#x38), [B05](02-react-blocks.md#b05), [B20](02-react-blocks.md#b20), [B21](02-react-blocks.md#b21), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/mention.tsx`; `sibling browser test`

**Read:** [S13](sources.md#s13)

**Change:** Adapt mention suggestions, insertion, and stored identifiers. Keep the native fallback label readable.

**Close when:** Selecting a suggestion inserts one atomic mention. Save/reopen preserves its identifier, label, and existing trigger behavior.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X23

- [x] Persist inline and block math sources

**Depends:** [B02](02-react-blocks.md#b02), [P04](03-agent-projection.md#p04), [B19](02-react-blocks.md#b19), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/blocks/math-codec.ts`; `sibling math-codec.test.ts`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Store original LaTeX for inline and block equations. Retain imported OMML; generate a readable Word fallback for application equations.

**Close when:** Inline and multiline LaTeX roundtrip exactly. Editing surrounding text preserves imported OMML fingerprints.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/x23-math-sources`. **Evidence:** Runtime `e0846e38`, final test source `ecf76b9f`. The codec preserves exact inline/multiline LaTeX and unknown payload fields. Native insertion commits source metadata and readable Word fallback together. Imported inline and display OMML keep exact fingerprints through adjacent edits, history, and DOCX reopening. Preparation initializes shared payload infrastructure before first writers. `bun test --preload ./scripts/test/preload-dom.ts` over the 23 files in [math storage evidence](evidence/math-source-storage.json) passes 213 cases. Final focused checks pass 17 cases. The actual compiled private/core harness passes all five shared cases, including concurrency, deletion during peer editing, both histories, remote geometry, and DOCX/room reopening. Private types/build, root format/lint, and the collaboration policy pass. X24 controls and D30 AI atom execution remain separate.

## X24

- [x] Connect equation editing controls

**Depends:** [X23](08-domain-content.md#x23), [B05](02-react-blocks.md#b05), [B20](02-react-blocks.md#b20), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/math.tsx`; `sibling browser test`

**Read:** [S13](sources.md#s13)

**Change:** Adapt existing equation input, preview, and insertion controls. Apply source and fallback changes in one transaction.

**Close when:** Edit an inline equation and a block equation. One undo restores both source and displayed equation.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, application branch `claude/docx-native-cutover`, fork `db3144f4`. **Evidence:** `blocks/equation.tsx` registers an `equation` view (rendered formula over the linear-source fallback, reserved height) and an `inline_equation` chip. Double-click or Enter opens a Plate-free MathLive dialog; save calls `updateAtomicBlock`, which now also updates inline controls with plain runs that keep the first run's formatting. Payload and Word fallback change in one undo unit; suggest and view mode refuse visibly. `ui/math-display.tsx` now holds the shared MathLive renderer. Fork `blocks/update-atomic.test.ts` passes 4 cases, including an inline equation that keeps its sentence and bold run through save and undo. Chromium `equations.test.ts` passes 3 cases: block and inline edits by typing in the math field, exact payloads, Word fallback text, Markdown, undo of each edit, redo, native typing after the dialog, suggest-mode refusal, and two sessions in both client orders with peer typing, history, and equal saved digests. Switching between inline and block display is not offered.

## X25

- [ ] Render and edit SMILES chemistry blocks

**Depends:** [B02](02-react-blocks.md#b02), [B05](02-react-blocks.md#b05), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/smiles.tsx`; `sibling browser test`

**Read:** [S13](sources.md#s13)

**Change:** Reuse the existing chemistry renderer and editing controls. Persist SMILES source and an embedded Word image fallback.

**Close when:** Editing a structure changes source and fallback together. Save/reopen preserves the exact SMILES string.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X26

- [ ] Connect image insertion and resizing

**Depends:** [B05](02-react-blocks.md#b05), [B11](02-react-blocks.md#b11), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/image.tsx`; `sibling browser test`

**Read:** [S13](sources.md#s13), [S12](sources.md#s12)

**Change:** Connect embedded images, existing VFS references, dimensions, alt text, and resize controls to native drawings.

**Close when:** Insert and resize an image. Reopen retains its embedded bytes, dimensions, alt text, and pristine application reference.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X27

- [ ] Connect gallery ordering and lightbox controls

**Depends:** [X26](08-domain-content.md#x26), [B05](02-react-blocks.md#b05), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/gallery.tsx`; `sibling browser test`

**Read:** [S13](sources.md#s13), [S11](sources.md#s11)

**Change:** Adapt gallery grouping, ordering, removal, and lightbox UI. Preserve individual image metadata and native fallbacks.

**Close when:** Reorder two images, open the lightbox, and undo the reorder. Markdown and DOCX exports retain the expected order.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X28

- [ ] Connect uploads and media placeholders

**Depends:** [X26](08-domain-content.md#x26), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/uploads.tsx`; `sibling browser test`

**Read:** [S13](sources.md#s13)

**Change:** Reuse the existing upload service and progress UI. Replace a placeholder only when document and placeholder identities still match.

**Close when:** Successful upload replaces its own placeholder once. Cancel or document switch cannot insert media into another document.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X29

- [x] Render audio blocks

**Depends:** [B02](02-react-blocks.md#b02), [B05](02-react-blocks.md#b05), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/media.tsx`; `app:scripts/docx-native/harness/media.test.ts`

**Read:** [S13](sources.md#s13)

**Change:** Adapt existing audio playback and source attributes. Persist a native Word link or attachment fallback.

**Close when:** Playback works after reopen. Export retains the audio source and fallback label without saving transient player state.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, app branch `claude/docx-native-editor` (continuing Claude). **Evidence:** Application worktree `sciobot-next-wts/docx-native`, branch `claude/docx-native-editor`, baseline `f32fefc83` plus inherited local work. Vendor source `07701aae993b87f83532e87ee87413fc826b3960`, pinned in `vendor/docx-native/SOURCE.json`. `AudioBlock` uses native audio controls over a readable Word hyperlink fallback. A generated WAV plays through a real play-button click before and after DOCX import. Playback changes neither payloads nor saved digests; transient player state remains local. The 3 media cases pass: sources and fallback labels survive export/import; XML relationships retain each original target. Two real sessions in both client orders retain payloads/widgets through offline typing, deletion, reconnect, peer-preserving undo/redo, saved convergence and reopen. Combined Chromium command `bun test --timeout 600000 scripts/docx-native/harness/placeholders.test.ts scripts/docx-native/harness/lesson-title.test.ts scripts/docx-native/harness/media.test.ts`: 9 pass, 0 fail, 74 assertions; `/tmp/codex-docx-app-harness.log`. Harness/application typecheck (`bun check -p scripts/docx-native/tsconfig.json`) passes 504 files. Scoped application lint passes with one existing Fast Refresh warning in `media.tsx`.

## X30

- [x] Render video blocks

**Depends:** [B02](02-react-blocks.md#b02), [B05](02-react-blocks.md#b05), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/media.tsx`; `app:scripts/docx-native/harness/media.test.ts`

**Read:** [S13](sources.md#s13)

**Change:** Adapt existing video controls and source attributes. Preserve dimensions, poster data, and native fallback.

**Close when:** A video fixture reopens with the same source and poster. Export retains a usable fallback.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, app branch `claude/docx-native-editor` (continuing Claude). **Evidence:** Application worktree `sciobot-next-wts/docx-native`, branch `claude/docx-native-editor`, baseline `f32fefc83` plus inherited local work. Vendor source `07701aae993b87f83532e87ee87413fc826b3960`, pinned in `vendor/docx-native/SOURCE.json`. `VideoBlock` retains source, poster, width, and height over a native Word hyperlink fallback. The video fixture reopens with the exact source, poster and 640x360 dimensions. The intentionally empty video fixture tests source/control persistence; actual video decoding is outside this evidence. The 3 media cases pass: sources and fallback labels survive export/import; XML relationships retain each original target. Two real sessions in both client orders retain payloads/widgets through offline typing, deletion, reconnect, peer-preserving undo/redo, saved convergence and reopen. Combined Chromium command `bun test --timeout 600000 scripts/docx-native/harness/placeholders.test.ts scripts/docx-native/harness/lesson-title.test.ts scripts/docx-native/harness/media.test.ts`: 9 pass, 0 fail, 74 assertions; `/tmp/codex-docx-app-harness.log`. Harness/application typecheck (`bun check -p scripts/docx-native/tsconfig.json`) passes 504 files. Scoped application lint passes with one existing Fast Refresh warning in `media.tsx`.

## X31

- [x] Render file attachment blocks

**Depends:** [B02](02-react-blocks.md#b02), [B05](02-react-blocks.md#b05), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/media.tsx`; `app:scripts/docx-native/harness/media.test.ts`

**Read:** [S13](sources.md#s13)

**Change:** Adapt file metadata, download actions, and application references. Preserve a native Word label and link.

**Close when:** A restored file block downloads the intended asset. Filename, size, source reference, and fallback survive export.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, app branch `claude/docx-native-editor` (continuing Claude). **Evidence:** Application worktree `sciobot-next-wts/docx-native`, branch `claude/docx-native-editor`, baseline `f32fefc83` plus inherited local work. Vendor source `07701aae993b87f83532e87ee87413fc826b3960`, pinned in `vendor/docx-native/SOURCE.json`. `FileBlock` preserves filename, size, source reference and native Word link. A real download returns the intended filename and exact PDF fixture bytes; a reopened block keeps the same metadata and download target. The 3 media cases pass: sources and fallback labels survive export/import; XML relationships retain each original target. Two real sessions in both client orders retain payloads/widgets through offline typing, deletion, reconnect, peer-preserving undo/redo, saved convergence and reopen. Combined Chromium command `bun test --timeout 600000 scripts/docx-native/harness/placeholders.test.ts scripts/docx-native/harness/lesson-title.test.ts scripts/docx-native/harness/media.test.ts`: 9 pass, 0 fail, 74 assertions; `/tmp/codex-docx-app-harness.log`. Harness/application typecheck (`bun check -p scripts/docx-native/tsconfig.json`) passes 504 files. Scoped application lint passes with one existing Fast Refresh warning in `media.tsx`.

## X32

- [x] Render embedded media blocks

**Depends:** [B02](02-react-blocks.md#b02), [B05](02-react-blocks.md#b05), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/media.tsx`; `app:scripts/docx-native/harness/media.test.ts`

**Read:** [S13](sources.md#s13)

**Change:** Reuse the existing allowed embed types and activation behavior. Store a readable native link fallback.

**Close when:** Supported embeds retain source and dimensions after reopen. Unrecognized types keep their fallback and stored payload.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, app branch `claude/docx-native-editor` (continuing Claude). **Evidence:** Application worktree `sciobot-next-wts/docx-native`, branch `claude/docx-native-editor`, baseline `f32fefc83` plus inherited local work. Vendor source `07701aae993b87f83532e87ee87413fc826b3960`, pinned in `vendor/docx-native/SOURCE.json`. `EmbedBlock` reuses allowed provider parsing and keeps a readable native Word hyperlink fallback. The recognized YouTube fixture reopens at the same source and stored width. An unrecognized provider retains its payload and visible native link without a widget. External providers are blocked in the test; live provider availability is outside this evidence. The 3 media cases pass: sources and fallback labels survive export/import; XML relationships retain each original target. Two real sessions in both client orders retain payloads/widgets through offline typing, deletion, reconnect, peer-preserving undo/redo, saved convergence and reopen. Combined Chromium command `bun test --timeout 600000 scripts/docx-native/harness/placeholders.test.ts scripts/docx-native/harness/lesson-title.test.ts scripts/docx-native/harness/media.test.ts`: 9 pass, 0 fail, 74 assertions; `/tmp/codex-docx-app-harness.log`. Harness/application typecheck (`bun check -p scripts/docx-native/tsconfig.json`) passes 504 files. Scoped application lint passes with one existing Fast Refresh warning in `media.tsx`.

## X33

- [ ] Connect media caption editing

**Depends:** [X26](08-domain-content.md#x26), [X29](08-domain-content.md#x29), [X30](08-domain-content.md#x30), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/caption.tsx`; `sibling browser test`

**Read:** [S13](sources.md#s13)

**Change:** Connect existing caption controls to native editable paragraphs associated with their media block.

**Close when:** Edit and undo a caption without replacing its media. Save/reopen and static export retain caption ordering.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X34

- [ ] Render shared MDX content and reference tags

**Depends:** [P05](03-agent-projection.md#p05), [B05](02-react-blocks.md#b05), [B20](02-react-blocks.md#b20), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/shared-mdx.tsx`; `sibling browser test`

**Read:** [S11](sources.md#s11)

**Change:** Reuse MdxComponentPreview for preserved ASTs. Cover QR, sciocil, competence, law, entity, qref, vfs-cite, and action placeholders.

**Close when:** One fixture per listed tag retains attributes and children. Existing preview actions receive the same identifiers.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X35

- [ ] Render shared MDX material action tags

**Depends:** [X34](08-domain-content.md#x34), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/mdx-material-actions.tsx`; `sibling browser test`

**Read:** [S11](sources.md#s11)

**Change:** Connect material-share, material-trash, material-trash-batch, material-reference, material-rating, and upsell through existing application components.

**Close when:** Each action receives preserved node or entity IDs. Opening and cancelling its existing confirmation flow leaves document data unchanged.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X36

- [ ] Render data-only suggestion payloads

**Depends:** [P05](03-agent-projection.md#p05), [X34](08-domain-content.md#x34), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/mdx-suggestions.tsx`; `sibling browser test`

**Read:** [S11](sources.md#s11)

**Change:** Reuse the suggestions payload renderer with its complete nested AST. Keep these content suggestions distinct from document review groups.

**Close when:** Nested suggestion messages and data-only children survive reopen. Their renderer does not change document review counts.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X37

- [x] Project application container tags for callout, story, and solution

**Depends:** [P04](03-agent-projection.md#p04)

**Primary files:** `fork:packages/sciobot/src/projection/block-tags.ts`; `focused projection test`

**Read:** [S16](sources.md#s16), [S11](sources.md#s11)

**Change:** Add `callout`, `story`, and `solution` container definitions to `DOMAIN_BLOCK_TAGS` with the application Markdown rules' attribute mapping. Callout stores the MDX `type` as payload `variant`; legacy `variant` and optional `icon` round-trip. Story and solution serialize their remaining attributes in order (`propsToAttributes(rest)`). Native children keep their paragraph IDs and offsets.

**Close when:** Working-view Markdown of a fixture with one block of each type emits `<callout type="tip" icon="…">`, `<story …>`, and `<solution …>` around native children with the original attributes. Attribute spans resolve to payload property paths. Unknown attributes stay inert data.

**Check:** Add focused projection cases beside the P04 tests. Run private typecheck/build and a compatibility decision. Pure projection needs no session test.

**Claim:** codex X37 worker, branch `codex/react-block-proof`. **Evidence:** Native baseline `3278fa04`, local working tree. `bun scripts/sciobot-port/capture-container-reference.mjs` captures six application attribute outputs and three modern/legacy callout alias parses. `packages/sciobot/src/__tests__/fixtures/container-tags.json` records application revision `aec740a3` and four source hashes. `bun test packages/sciobot/src/projection/container-tags.test.ts packages/sciobot/src/projection/blocks.test.ts` passes 17 cases and 166 assertions. Working-view fixtures retain exact ordered attributes, optional icons, native paragraph IDs, UTF-16 offsets, property paths, nested/empty containers, and save/reopen payloads. Callout serializes `variant` as `type` after remaining attributes, matching the application rule. Unknown properties remain inert, including JSON, boolean, constructor, and event-shaped names. Reads preserve package bytes, revision, and history. `bun run --filter '@sciobot/docx' build` passes with 189 private modules and external native core. Scoped Prettier, oxlint, and diff checks pass. Combined private types pass 1958 files. The completed-leaf checkpoint `988042a5` builds 190 private modules with external native core; The isolated completed-leaf suite passes 2652 cases with one package-path assertion failure caused by shared dependency links. The unchanged package-boundary test passes separately in the real checkout: one case and eight assertions. Logs: `/tmp/codex-closed-private-suite2.log` and `/tmp/codex-main-boundary-final.log`. These runs cover 2653 cases across 197 files; later work remains separately verified. No-impact decision: `.collaboration/changes/sciobot-container-tag-projection.json`. This pure projection change needs no collaboration mutation scenario. X13–X15 application rendering remains separate; no importer or canonical edit path changes.

## X38

- [x] Persist native mention identity and Markdown syntax

**Depends:** [B02](02-react-blocks.md#b02), [P04](03-agent-projection.md#p04), [B19](02-react-blocks.md#b19), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/blocks/mention-codec.ts`; `focused native and collaboration tests`; `fork:packages/sciobot/src/projection/block-tags.ts`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Define application-owned inline mention controls. Store the application `key` separately from display `value`, preserving document identifiers beginning with `/` and optional private attributes. Keep a readable native fallback and project the existing mention Markdown syntax without evaluating content. Reuse the same codec for D30 insertion plans and later X22 controls.

**Close when:** Person and document mentions keep exact identifiers, labels, and private attributes through native insertion, Markdown projection, undo/redo, and save/reopen. Adjacent native text remains editable. The existing mention parser reads the projected identifier and display label exactly.

**Check:** Add focused explicit codec/projection/native cases. Add the V01 two-session scenario for concurrent insertions, deletion during peer editing, both histories, remote geometry, reconnect, and DOCX/room reopening. Run private types/build, scoped formatting/lint, and the collaboration policy.

**Claim:** codex solo, branch `codex/x38-native-mentions`. **Evidence:** Runtime `e0cbe6e8`, final test source `058f66e0`. Native person/document controls preserve exact keys, labels, private attributes, readable fallback, and independent history. Escaped existing mention syntax roundtrips empty, multiline, Unicode, Markdown-looking, and at-sign labels. The captured raw mention parser references remain unchanged. Field spans distinguish label/identifier ownership without false value slices. Native GFM cells stay GFM. Five real shared cases cover concurrent insertion, peer editing during deletion, both histories, remote geometry, unchanged shared state during projection reads, and DOCX/room reopening. `bun test --preload ./scripts/test/preload-dom.ts` over the 45 files in [native mention evidence](evidence/native-mention-storage.json) passes 504 cases at `19bb6de6`. Final field-range checks pass 87 cases across eight files; the actual compiled private/core harness passes 14 cases. Private types (2129 files), build, root formatting/lint with final scoped checks, compatibility policy, and backlog validation pass. All 199 application source hashes remain exact; inventory changes only add X38 to the two mention mappings. X22 browser controls and D30 AI execution remain open.
