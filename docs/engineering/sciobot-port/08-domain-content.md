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

- [ ] Render curriculum recap blocks

**Depends:** [X02](08-domain-content.md#x02), [B20](02-react-blocks.md#b20), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/curriculum.tsx`; `sibling browser test`

**Read:** [S09](sources.md#s09), [S11](sources.md#s11)

**Change:** Adapt curriculum framework and school/catalog source fields. Reuse existing competence displays and application navigation.

**Close when:** A school curriculum fixture retains framework and source across save/reopen. Competence actions receive the correct identifiers.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

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

- [ ] Render activity name and duration controls

**Depends:** [B05](02-react-blocks.md#b05), [B06](02-react-blocks.md#b06), [R12](05-suggestions.md#r12), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/activity.tsx`; `sibling browser test`

**Read:** [S09](sources.md#s09)

**Change:** Adapt activity name, duration, and resource fields. Route field changes through the canonical payload transaction.

**Close when:** Editing name or duration creates one property change when review is active. Native activity content remains unchanged.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, isolated `sciobot-next` worktree at detached commits `7372c120c`, `9767144a1`, and `2ac716769`. **Evidence:** `src/components/workspace/docx/blocks/activity.tsx` renders native activity name and duration controls while leaving the activity body in the native slot. Edit-mode changes use the canonical payload transaction. Suggestion-mode changes use `planDomainProperties` and `trackPropertyPlan`, including the generated fallback label and revision-bound owned review context. The Chromium command `bun test --timeout 300000 scripts/docx-native/harness/activity.test.ts` passes 3 cases: edit-mode name/duration with native body preservation, plus two-session property suggestion convergence for both client orders (`73/41` and `41/73`). The pair cases assert proposed Markdown convergence and the single-session case asserts the native body remains intact. Follow-up `c39572ae1` removes body-text label guessing, requires the explicitly generated label, preserves attribute removals, and allocates unique property proposal IDs. The same three browser cases pass. The follow-up application commit `b05fa25b7` adds activity resource editing in edit mode, localized controls, and a pure stored-property projection for pending header values. The controls preserve the Plate duration bounds of zero through 120 minutes. Harness module arrays now remain stable across renders. Review setup uses the actual installed adapter API rather than ignored options from a newer source checkout.

Command `bun test --timeout 120000 scripts/docx-native/harness/activity.test.ts scripts/docx-native/harness/text-export.test.ts` passes **18 cases, zero failures, 372 assertions** against the installed runtime. Seventeen activity cases cover exact canonical versus proposed attributes; one group and one property membership per name/duration edit; unchanged complete native body fingerprints and control owners; accept/reject; proposal and decision undo/redo; resource edit history; both fixed client orders; offline peer body typing and tail deletion; peer deletion undo/redo; pending DOCX reopening; final DOCX reopening; matching saved digests; and caret geometry. The export regression remains included. Source checks pass 6,677 application files and 2,256 harness files. Scoped formatting, lint, and diff checks pass. X05 stays open for the resource-only suggestion route and consecutive field edits while a prior property proposal remains pending. Those paths must preserve complete generated Word labels and proposal ownership before closure.

## X06

- [ ] Derive phase and preparation duration totals

**Depends:** [X01](08-domain-content.md#x01), [X04](08-domain-content.md#x04), [X05](08-domain-content.md#x05)

**Primary files:** `app:src/lib/docx/phase-durations.ts`; `sibling phase-durations.test.ts`

**Read:** [S09](sources.md#s09), [S15](sources.md#s15)

**Change:** Derive totals from activity payloads with the existing phase rules. Do not save duplicate derived duration values.

**Close when:** Changing one activity updates its phase and recap totals. Undo returns all displayed totals without a second write.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** unclaimed. **Evidence:** pending.

## X07

- [ ] Port preparation retiming actions

**Depends:** [X06](08-domain-content.md#x06), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/retime-document.ts`; `sibling retime-document.test.ts`

**Read:** [S15](sources.md#s15)

**Change:** Reuse retiming rules and write all changed activity durations in one canonical transaction.

**Close when:** Existing retiming examples produce the same minutes. One undo restores every affected duration.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

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

- [ ] Preserve listening markup and audio references

**Depends:** [B02](02-react-blocks.md#b02), [P05](03-agent-projection.md#p05), [D14](04-edit-planner.md#d14), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/blocks/listening-codec.ts`; `sibling listening-codec.test.ts`

**Read:** [S12](sources.md#s12), [S11](sources.md#s11)

**Change:** Preserve speakers, lines, voices, portraits, seed, audioRef, and exercise data in the payload. Keep the native transcript complete.

**Close when:** Existing listening Markdown fixtures roundtrip. Editing a nearby sentence retains generated audio and speaker references.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X10

- [ ] Render listening exercises with existing services

**Depends:** [X09](08-domain-content.md#x09), [B05](02-react-blocks.md#b05), [B17](02-react-blocks.md#b17), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/listening.tsx`; `sibling browser test`

**Read:** [S12](sources.md#s12)

**Change:** Adapt listening playback, generation, and cache hydration. Keep local playback state separate from saved exercise data.

**Close when:** A restored exercise hydrates existing audio without regeneration. Playback and long transcripts remain usable inside paginated content.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X11

- [ ] Render source citation chips

**Depends:** [B05](02-react-blocks.md#b05), [B20](02-react-blocks.md#b20), [B21](02-react-blocks.md#b21), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/source-chip.tsx`; `sibling browser test`

**Read:** [S10](sources.md#s10)

**Change:** Adapt source and book-source chips with existing provenance services. Preserve URL, title, author, type, bookId, and trust.

**Close when:** A chip opens the matching provenance view. Save/reopen and atomic AI replacement retain all citation attributes.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X12

- [ ] Render source excerpt containers

**Depends:** [X11](08-domain-content.md#x11), [B06](02-react-blocks.md#b06), [D12](04-edit-planner.md#d12), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/source-excerpt.tsx`; `sibling browser test`

**Read:** [S10](sources.md#s10)

**Change:** Adapt excerpt framing around native body text. Apply existing citation identity rules independently from descendant edits.

**Close when:** Changing an excerpt sentence preserves its wrapper. Changing an identity field produces the expected structural review group.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X13

- [ ] Render callout blocks

**Depends:** [B05](02-react-blocks.md#b05), [B06](02-react-blocks.md#b06), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01), [X37](08-domain-content.md#x37)

**Primary files:** `app:src/components/workspace/docx/blocks/callout.tsx`; `sibling browser test`

**Read:** [S16](sources.md#s16)

**Change:** Adapt existing callout attributes and content UI. Keep nested lists and text native.

**Close when:** The callout Markdown fixture survives import, typing, and export with its original attributes.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** Blocked on [X37](08-domain-content.md#x37): the native Markdown projection (`DOMAIN_BLOCK_TAGS`, used by the working view) has no tag definition for this type. The block projects only as its native children, so its export cannot carry the original attributes. The application cannot pass definitions to `readWorkingView`.

## X14

- [ ] Render story blocks

**Depends:** [B05](02-react-blocks.md#b05), [B06](02-react-blocks.md#b06), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01), [X37](08-domain-content.md#x37)

**Primary files:** `app:src/components/workspace/docx/blocks/story.tsx`; `sibling browser test`

**Read:** [S16](sources.md#s16)

**Change:** Adapt story metadata and content UI without PlateElement. Keep the story body in a native slot.

**Close when:** A nested story fixture preserves metadata and paragraph editing through save/reopen.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** Blocked on [X37](08-domain-content.md#x37): the native Markdown projection (`DOMAIN_BLOCK_TAGS`, used by the working view) has no tag definition for this type. The block projects only as its native children, so its export cannot carry the original attributes. The application cannot pass definitions to `readWorkingView`.

## X15

- [ ] Render solution blocks

**Depends:** [B05](02-react-blocks.md#b05), [B06](02-react-blocks.md#b06), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01), [X37](08-domain-content.md#x37)

**Primary files:** `app:src/components/workspace/docx/blocks/solution.tsx`; `sibling browser test`

**Read:** [S16](sources.md#s16)

**Change:** Adapt solution reveal controls and attributes. Keep reveal state separate from saved answer content.

**Close when:** Reveal and hide retain the answer. Export includes the same solution content as the existing static renderer.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** Blocked on [X37](08-domain-content.md#x37): the native Markdown projection (`DOMAIN_BLOCK_TAGS`, used by the working view) has no tag definition for this type. The block projects only as its native children, so its export cannot carry the original attributes. The application cannot pass definitions to `readWorkingView`.

## X16

- [ ] Render editable code blocks

**Depends:** [B05](02-react-blocks.md#b05), [B06](02-react-blocks.md#b06), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/code.tsx`; `sibling browser test`

**Read:** [S13](sources.md#s13)

**Change:** Preserve code text, language, and whitespace. Reuse syntax display while routing edits to canonical native text.

**Close when:** Tab, Enter, paste, and language changes retain exact code text. Word fallback and Markdown export preserve whitespace.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## X17

- [ ] Render collapsible toggle blocks

**Depends:** [B05](02-react-blocks.md#b05), [B06](02-react-blocks.md#b06), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/toggle.tsx`; `sibling browser test`

**Read:** [S13](sources.md#s13)

**Change:** Adapt collapse controls around native descendants. Reveal collapsed ancestors for AI navigation and selection.

**Close when:** An AI target inside a closed toggle becomes visible. Collapse does not delete or duplicate its native content.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

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

- [ ] Map application columns to native table geometry

**Depends:** [B02](02-react-blocks.md#b02), [B06](02-react-blocks.md#b06), [P04](03-agent-projection.md#p04), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/blocks/columns-codec.ts`; `sibling columns-codec.test.ts`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Represent column groups as borderless native tables with application metadata. Preserve nested blocks and per-column widths.

**Close when:** Two and three-column fixtures retain nested text and tables through save/reopen. The Word fallback preserves column ordering.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

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

- [ ] Connect equation editing controls

**Depends:** [X23](08-domain-content.md#x23), [B05](02-react-blocks.md#b05), [B20](02-react-blocks.md#b20), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/blocks/math.tsx`; `sibling browser test`

**Read:** [S13](sources.md#s13)

**Change:** Adapt existing equation input, preview, and insertion controls. Apply source and fallback changes in one transaction.

**Close when:** Edit an inline equation and a block equation. One undo restores both source and displayed equation.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

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
