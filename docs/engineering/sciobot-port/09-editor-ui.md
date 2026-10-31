# Editor interactions and review presentation

Connect existing application controls to native selection, commands, and private review state.

Read the [dispatch and completion rules](../../../SCIOBOT_PORT_TODO.md) before claiming a leaf.

Each named test or artifact is a deliverable. These specifications do not claim those checks already exist.

## U01

- [x] Connect the floating selection toolbar

**Depends:** [C03](01-contracts.md#c03), [B07](02-react-blocks.md#b07), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/chrome/floating-toolbar.tsx`; `sibling browser test`

**Read:** [S01](sources.md#s01), [S17](sources.md#s17)

**Change:** Adapt selection toolbar visibility and positioning. Preserve native selection on button presses. U15 connects the production Add to chat action.

**Close when:** Native text selection opens the toolbar at the correct position after scrolling and zoom. Widget form selections and read-only documents hide it.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, application branch `claude/docx-native-cutover`. **Evidence:** Application commits `d9476aac9` and `f022a7bab`. `chrome/floating-toolbar.tsx` shows the production toolbar (the Add to chat action) above a native range selection. It tracks `selectionChange`, scroll, resize, and the document `selectionchange` event. Collapsed selections, view mode, and selections inside block widgets show nothing. The document event closes a defect: a selection in widget chrome does not move the editor selection, so the toolbar stayed open over it. Button presses keep the native selection. Chromium `floating-toolbar.test.ts` passes 7 cases: keyboard range selection, position above the selection, position after scrolling, hiding after the action and in view mode; position above and inside the selected line at editor zoom 1.5 and 0.75; hiding for a selection in activity widget chrome (fails without the fix); quote focus; and two sessions in both client orders with peer typing during a held selection and equal saved digests.

## U02

- [ ] Connect the slash insertion menu

**Depends:** [B11](02-react-blocks.md#b11), [N02](07-native-content.md#n02), [N08](07-native-content.md#n08), [N09](07-native-content.md#n09), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/chrome/slash-menu.tsx`; `sibling browser test`

**Read:** [S17](sources.md#s17)

**Change:** Port production slash entries, localized filtering, keyboard navigation, and insertion scope. Retain plugin and inferred-feature gates for math, chemistry, and listening. Reuse registered block insertion.

**Close when:** Every captured slash entry inserts its expected block. Escape cancels without stray slash text or history entries.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, application branch `claude/docx-native-cutover`. **Evidence:** Progress, not closure. Application commits `4ae32fc5f` and `287114e8d`. `ui/slash-catalog-data.ts` now holds the production entries without Plate. `chrome/slash-menu.tsx` opens on "/" at a paragraph start or after whitespace, as Plate does; the slash and the query never enter the document, so Escape leaves no text and no history entry, and a query with no match lands as typed text. It filters with the production matcher and the same inferred-feature gates, navigates with arrows, and chooses with Enter, Tab, or a click. Native commands cover text, three headings, quote, bullet, numbered, and to-do lists, table, code block, callout, story, solution, block equation, and inline equation. Chromium `slash-menu.test.ts` passes 15 cases: each entry inserts its block, keyboard choice, Escape with an unchanged revision, Backspace past the slash, "km/h" and an unmatched query as text, and two sessions in both client orders with peer typing, undo, and equal saved digests. The full harness passes 248 cases. Remaining: image, video, audio, file, listening, table of contents, molecule, three columns, toggle, and date entries (they do not appear until the native editor can insert them).

## U03

- [x] Connect block menus and block selection

**Depends:** [B07](02-react-blocks.md#b07), [B12](02-react-blocks.md#b12), [B13](02-react-blocks.md#b13), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/chrome/block-menu.tsx`; `sibling browser test`

**Read:** [S17](sources.md#s17)

**Change:** Adapt block selection, reachable secondary actions, duplication, delete, and multi-block scope. Resolve actions by stable block IDs. Preserve disabled block context menus.

**Close when:** Select two adjacent activities and delete them once. Undo restores both. Duplicating a block gives it fresh payload IDs.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, application branch `claude/docx-native-cutover`, fork `00c33e27`. **Evidence:** Fork `blocks/remove.ts` adds `removeBlocks`, which removes several registered wrappers and their payload records (nested ones included) as one native undo unit and refuses a wrapper inside another selected one. Fork `blocks/duplicate.ts` adds `duplicateBlock`: the copy lands after the original, every copied block gets a fresh payload ID and tag, control `w:id` values and paragraph `w14` identities are re-minted, and edit mode is required. Fork `blocks/block-selection.test.ts` passes 4 cases: two adjacent activities (one with a nested callout) deleted in one step with undo, redo, and reopen; nested-selection refusal; duplication with fresh IDs, independent edits, reopen, undo, and refusals; and two sessions in both receiver orders with peer typing, author undo, and equal saved digests. Application commit `1f92a7d1f` adds a block handle in the left margin of activities and container blocks (`blocks/block-handle.tsx`): click selects, Shift+click extends over sibling blocks, and the first selected block shows Duplicate and Delete. `blocks/block-selection.ts` maps Delete, Backspace, Mod+D, and Escape, and a caret move clears the selection. Chromium `block-selection.test.ts` passes 4 cases: two adjacent activities deleted with Delete and restored by one undo with their payloads; menu deletion and Escape; duplication with a fresh payload ID, Mod+D, and undo; and two sessions in both client orders with peer typing, author undo, and equal saved digests.

## U04

- [x] Connect drag and drop block movement

**Depends:** [B13](02-react-blocks.md#b13), [U03](09-editor-ui.md#u03), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/chrome/drag-drop.tsx`; `sibling browser test`

**Read:** [S17](sources.md#s17)

**Change:** Adapt drag handles, drop targets, and nested movement to native block moves. Preserve keyboard access to move actions.

**Close when:** Move an activity across a page and into another phase. The drop creates one undo entry and retains metadata.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, application branch `claude/docx-native-cutover`, fork `00c33e27`. **Evidence:** Application commit `db3ba4c1c`. A press on a block handle that moves more than five pixels drags the block; a still press selects it. `blocks/block-drag.ts` finds the paragraph fragment nearest the drop point and lands the block before or after that paragraph's outermost ancestor inside the nearest registered container or the body. A container of the dragged block's own type is not a destination, so an activity dropped on another activity lands beside it. The move goes through the fork `moveBlock`, which keeps payload records and descendant IDs and commits one native undo unit. Chromium `block-drag.test.ts` passes 3 cases: an activity dragged from the evocation phase on page 1 into the understanding phase on page 2, with unchanged payloads and native body, no block selection after the drop, and one undo and one redo; and two sessions in both client orders where the peer types in the moved activity's body during the drag, with convergence, author undo, and equal saved digests.

## U05

- [x] Connect Markdown autoformat shortcuts

**Depends:** [N02](07-native-content.md#n02), [N03](07-native-content.md#n03), [N08](07-native-content.md#n08), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/chrome/autoformat.ts`; `sibling browser test`

**Read:** [S17](sources.md#s17)

**Change:** Map the production autoformat rules to native commands. Keep code and widget form inputs outside document autoformat.

**Close when:** Captured heading, list, and mark shortcuts produce expected structures. Immediate undo restores the literal trigger text.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, application branch `claude/docx-native-cutover`, fork `96c45407`. **Evidence:** Application commit `017b84942` installs the fork's input rules in every native workspace editor: `installBlockInputRules` (`#`, `##`, `###`, `>`, and the heading and quote chords), `installListEditing` (bullet, ordered, and checklist triggers and list keys), `installMarkShortcuts`, and `installTableNavigation`. A defect surfaced and is fixed in the fork: the rules excluded widget furniture by `[data-docx-block-ui]`, which no element carries, so keys typed in widget forms could reach document shortcuts; they now use `[data-sciobot-block-ui]`, and `chrome/marks.test.ts` adds a regression that fails without the fix. Chromium `autoformat.test.ts` passes 7 cases on the packaged blank template: heading, quote, bullet, and ordered triggers, each with undo of the typed word and then of the shortcut back to the literal trigger; bold and italic chords; and two sessions in both client orders with peer typing and equal saved digests. The `---` rule trigger did not fire in the browser case and is not claimed. The full harness passes 233 cases; widget suites pass 60 cases on the vendored fix.

## U06

- [x] Connect emoji insertion

**Depends:** [N03](07-native-content.md#n03), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/chrome/emoji.tsx`; `sibling browser test`

**Read:** [S17](sources.md#s17)

**Change:** Reuse emoji search, trigger, and insertion controls with native text insertion.

**Close when:** Insert a multi-codepoint emoji at the caret. Arrow navigation and deletion retain valid text boundaries.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, application branch `claude/docx-native-cutover`, fork `4bd80ef8`. **Evidence:** Application commit `ff0f00ef3`. `chrome/emoji-menu.tsx` opens on `:` plus a query, searches `@emoji-mart/data` names and keywords, and inserts the native emoji through ordinary text input. `chrome/trigger-menu.ts` is the shared trigger lifecycle for the emoji and slash menus. A colon inside a word, an unmatched query, and Enter on a bare `:` stay literal text. A core defect surfaced: Backspace and Delete removed one UTF-16 unit, which splits a surrogate pair, so the store refused the edit and the key did nothing after an emoji. Fork commit `4bd80ef8` removes one grapheme (`editor/grapheme-step.ts`), with `editor/__tests__/grapheme-deletion.test.ts` (9 cases: rocket, rainbow flag, skin tone, and accented letter in both directions, undo). Chromium `emoji-menu.test.ts` passes 3 cases: a rainbow flag inserted at the caret between text, arrow keys and Backspace across it, keyboard choice, literal colons, and two sessions in both client orders with peer typing and equal saved digests.

## U07

- [x] Match exit-break and trailing-paragraph behavior

**Depends:** [B07](02-react-blocks.md#b07), [N02](07-native-content.md#n02), [N08](07-native-content.md#n08), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/chrome/breaks.ts`; `sibling browser test`

**Read:** [S17](sources.md#s17)

**Change:** Map exit-break keys and end-of-document typing to native paragraph operations. Keep empty trailing content out of false AI diffs.

**Close when:** Exit a table, code block, and activity with the captured keys. Typing after the last atomic block creates one valid paragraph.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, application branch `claude/docx-native-cutover`, fork `cdad157c`. **Evidence:** Fork `chrome/exit-break.ts` maps Mod+Enter and Mod+Shift+Enter to one body paragraph after or before the top-level structure at the caret; application commits `85ac47171` and `c94544abf` install it in every native workspace editor. A projection defect surfaced: an empty atomic block at the end of the body exports no fallback text, so its tag joined the end of the previous paragraph and every agent plan of that document failed. Fork commit `cdad157c` starts such a tag on its own Markdown block (`projection/blocks.ts`). Fork `chrome/exit-break.test.ts` passes 9 cases: table, code block, callout, and activity exits after and before with undo, redo, and reopen; refusal outside edit mode; two sessions in both receiver orders with peer typing and undo; and, for an empty last code block with and without a language, two empty exit paragraphs that plan no agent edit and one typed paragraph after the block (both fail without the fix). Chromium `exit-break.test.ts` passes 4 cases: an activity exit with undo, an empty exit after the last code block that changes only blank lines, and two sessions in both client orders with peer typing, undo, and equal saved digests; `slash-menu.test.ts` covers table, code block, and callout exits by keyboard.

## U08

- [x] Connect block placeholder text

**Depends:** [B06](02-react-blocks.md#b06), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/chrome/placeholders.tsx`; `sibling browser test`

**Read:** [S17](sources.md#s17)

**Change:** Adapt contextual placeholder labels using existing localization. Keep labels outside canonical Word text.

**Close when:** Empty supported blocks show their label. Typing hides it. Export and agent reads contain no placeholder text.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, app branch `claude/docx-native-editor` (continuing Claude). **Evidence:** Application worktree `sciobot-next-wts/docx-native`, branch `claude/docx-native-editor`, baseline `f32fefc83` plus inherited local work. Vendor source `07701aae993b87f83532e87ee87413fc826b3960`, pinned in `vendor/docx-native/SOURCE.json`. `mountDocxPlaceholders` paints existing paragraph and localized recap-material labels outside canonical text. Only the focused empty top-level paragraph gets the paragraph label; every empty material row gets its contextual label. Typing hides labels; deletion and undo restore the appropriate labels. Their DOM positions track native geometry. Agent Markdown reads, native body text, and saved XML contain no placeholder strings. The 3 placeholder cases pass: real caret/typing/deletion input, two real sessions in both client orders with distinct local caret labels, concurrent peer edits, reconnect, undo, and matching saved digests. `DocxWorkspaceEditor` disposes label nodes and subscriptions with its editor. Combined Chromium command `bun test --timeout 600000 scripts/docx-native/harness/placeholders.test.ts scripts/docx-native/harness/lesson-title.test.ts scripts/docx-native/harness/media.test.ts`: 9 pass, 0 fail, 74 assertions; `/tmp/codex-docx-app-harness.log`. Harness/application typecheck (`bun check -p scripts/docx-native/tsconfig.json`) passes 504 files. Scoped application lint passes with one existing Fast Refresh warning in `media.tsx`. Scoped oxfmt write/check passes for implementation and harness files.

## U09

- [x] Connect the review bar and group navigation

**Depends:** [R14](05-suggestions.md#r14), [R15](05-suggestions.md#r15), [R16](05-suggestions.md#r16), [R17](05-suggestions.md#r17), [R18](05-suggestions.md#r18), [R21](05-suggestions.md#r21), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/review/review-bar.tsx`; `sibling browser test`

**Read:** [S07](sources.md#s07), [S08](sources.md#s08), [S17](sources.md#s17)

**Change:** Adapt counts, next/previous, decisions, busy state, and shortcuts. Reveal nested targets and navigate by stable group identity.

**Close when:** Navigation reaches a group on a virtual page and inside a toggle. Resolving the active group selects the correct remaining group.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, application branch `claude/docx-native-cutover`. **Evidence:** Application commits `19c3244ea`, `14c765e0c`, and `ca138f7b5`. `review/review-bar.tsx` mounts in the native panel. It shows the active position and count, previous and next (buttons and Alt+Up/Alt+Down), accept and reject of the active group, and Accept all and Reject all through the fork group decisions. The active group is tracked by its stable ID; after a decision the group that followed it becomes active. Navigation opens a collapsed toggle or solution that hides the anchor, selects the anchor, and calls `revealPosition`. The workspace scroller carries the engine class `docx-editor__scroll-container`, and the editor pins 100% zoom so block widgets keep their scale. Chromium `review-bar.test.ts` passes 3 cases: navigation to a group two pages away, wrap-around shortcuts, accept and reject of the active group with the next selection, Reject all, undo restoring the group with its ID, and two sessions in both client orders with peer typing, author undo, and equal saved digests. `toggles.test.ts` navigates into a closed toggle and accepts the group there.

## U10

- [ ] Render old and proposed review content with word hints

**Depends:** [R14](05-suggestions.md#r14), [B06](02-react-blocks.md#b06), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/review/group-preview.tsx`; `sibling browser test`

**Read:** [S07](sources.md#s07), [S08](sources.md#s08), [S18](sources.md#s18)

**Change:** Show contiguous old content before proposed content with symmetric word hints. Keep preview DOM outside authoritative document text.

**Close when:** The C01 nested change shows one tight hunk. Preview text never enters copyable native selection, agent reads, or saved DOCX.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## U11

- [x] Render granular property review controls

**Depends:** [R12](05-suggestions.md#r12), [U09](09-editor-ui.md#u09), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/review/property-preview.tsx`; `sibling browser test`

**Read:** [S04](sources.md#s04), [S09](sources.md#s09), [S17](sources.md#s17)

**Change:** Show old and proposed activity names and durations beside their existing wrapper. Connect per-group accept and reject controls.

**Close when:** A duration change displays only the property change. Accept, reject, and undo preserve native descendants and nearby groups.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, application branch `claude/docx-native-cutover`. **Evidence:** Application commit `ed2c37735`. `review/property-preview.tsx` compares an activity's committed payload with its working projection and, for the activity's pending property groups, shows only the changed properties (name, duration, resource) with the old value struck and the proposed value inserted, plus accept and reject. It sits beside the activity, outside the native body. Decisions go through the fork group decisions. Chromium `property-review.test.ts` passes 4 cases: a duration proposal beside a text proposal in the same activity body shows one change (10 to 25); accept and reject each leave the text group and the native body fingerprint unchanged and update the payload; undo restores both groups and the original attributes; and two sessions in both client orders accept the property while the peer types, with equal Markdown, no panel for the peer after the decision, and equal saved digests.

## U12

- [x] Scroll to the first newly introduced suggestion

**Depends:** [U09](09-editor-ui.md#u09), [A06](06-direct-ai.md#a06), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/docx/review-scroll.ts`; `sibling review-scroll.test.ts and browser case`

**Read:** [S07](sources.md#s07), [S08](sources.md#s08)

**Change:** Port SuggestionScrollOrchestrator behavior across repeated streams. Reveal collapsed ancestors before measuring the destination.

**Close when:** A second wave scrolls to its first new group once. Existing groups and repeated renders do not cause extra scrolling.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, application branch `claude/docx-native-cutover`. **Evidence:** Application commits `303a85c32` and `ca138f7b5`. The review bar remembers the group IDs a session has shown. When a change brings a group it has not shown, it opens any collapsed toggle or solution that hides the group and scrolls once to the first such group with `revealPosition`; the selection does not move. Undo, redo, peer edits, and repeated renders bring back known groups or no group and do not scroll. Chromium `direct-edits.test.ts` passes 2 cases: a first wave, a second wave that replaces it, and undo; and, in two sessions, an AI wave inside a closed toggle on a later page that opens it and scrolls once, while the peer's toggle stays closed and the peer's typing does not scroll.

## U13

- [ ] Connect selection AI actions and menu lifecycle

**Depends:** [A04](06-direct-ai.md#a04), [A05](06-direct-ai.md#a05), [A08](06-direct-ai.md#a08), [U03](09-editor-ui.md#u03), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/chrome/ai-menu.tsx`; `sibling browser test`

**Read:** [S19](sources.md#s19), [S17](sources.md#s17)

**Change:** Connect the existing AI menu shortcut, selected-block scope, insert/generate mode, and edit mode. Reuse existing chat transport.

**Close when:** Generate inserts at the selected scope. Edit targets the selected content. Closing or switching documents cleans its transient anchor.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## U14

- [ ] Connect AI and remote cursor presentation

**Depends:** [B07](02-react-blocks.md#b07), [A04](06-direct-ai.md#a04), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/chrome/cursors.tsx`; `sibling browser test`

**Read:** [S17](sources.md#s17), [S18](sources.md#s18), [S25](sources.md#s25)

**Change:** Adapt existing cursor overlays to native geometry. Keep AI presentation cursors separate from optional collaboration presence.

**Close when:** AI and remote selections align inside a nested activity after zoom and repagination. Removing a peer removes its overlay.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## U15

- [x] Connect the Add to chat quote action

**Depends:** [U01](09-editor-ui.md#u01), [P20](03-agent-projection.md#p20), [I01](10-app-integration.md#i01), [V01](11-verification.md#v01)

**Primary files:** `app:src/components/workspace/docx/chrome/quote-action.tsx`; `sibling browser test`

**Read:** [S27](sources.md#s27), [S21](sources.md#s21)

**Change:** Reuse attemptAddToChat and the quote store with native mapped selection data. Include documentTitle, materialId, line numbers, text, and containsMdx.

**Close when:** An allowed click adds one exact quote and deselects the editor. Denied, empty, or unresolved selections add nothing. Pending text quotes match read_document.

**Check:** Exercise the real button, eligibility result, and quote store in Chromium. Check document switches during eligibility waits and two-session pending text selection.

**Claim:** claude, application branch `claude/docx-native-cutover`. **Evidence:** Application commits `d9476aac9` and `b836ed92a`. `quoteNativeSelection` (in `chrome/floating-toolbar.tsx`) reads the working view, maps the selection with the fork `resolveQuoteSelection`, refuses a caret, unresolved, or empty quote, asks the Plus gate (`attemptAddToChat`), and writes `documentTitle`, `materialId`, line numbers, text, and `containsMdx` to the quote store. It then collapses the selection. A denied add keeps the selection and adds nothing. A defect surfaced and is fixed: a caret quoted its whole line. Chromium `floating-toolbar.test.ts` checks that an allowed click adds one quote equal to its lines in the working view that `read_document` returns and deselects the editor; a caret and a denied gate add nothing (the case fails without the fix); and two sessions in both client orders quote current lines after peer typing.

## U16

- [x] Focus native document content from chat quotes

**Depends:** [P11](03-agent-projection.md#p11), [I06](10-app-integration.md#i06), [U07](09-editor-ui.md#u07), [V01](11-verification.md#v01)

**Primary files:** `app:src/lib/quote-focus.ts`; native branch and focused browser test

**Read:** [S27](sources.md#s27), [S21](sources.md#s21)

**Change:** Route document quote clicks to native working-view line mappings. Preserve material hydration, pending focus, workspace activation, and existing non-document routes.

**Close when:** Clicking a quote selects and scrolls to its actual pending content. Cross-material clicks wait for hydration and focus once. Stale loads cannot redirect focus.

**Check:** Cover the chat caller, same and different material routes, nested blocks, missing lines, two-session changes, and saved native reopen. Read or selection actions must not change document history.

**Claim:** claude, application branch `claude/docx-native-cutover`. **Evidence:** Application commits `da8b42214`, `b83ab4805`, and `b8592486d`. `quote-focus.ts` routes document quote clicks to `native-quote-focus.ts` when a native document is open. It reads the working view, maps the quoted lines through the projection source map (`quote-target.ts`, exact offsets through the fork `sourceOffsetsAt`), selects the native range, reveals it, and focuses the editor. A stale read after a document switch never moves focus. Two defects surfaced and are fixed for the native and Plate routes: a quote of another material set its pending focus before `hydrateFromMaterial`, which clears it, so the focus was lost; and a slow material load could switch the workspace after a later click. The pending focus now carries its material, is set after hydration, and only that material's native editor applies it; a click counter drops overtaken loads. Chromium cases: in `floating-toolbar.test.ts`, a quote click selects content that quotes back to the identical quote, a pending focus applies once ready, and a quote of pending suggested text selects that text and scrolls it into view; in `native-persistence.test.ts`, a quote of another saved native material opens it and selects its line once, and an overtaken load neither switches the material nor leaves a pending focus.
