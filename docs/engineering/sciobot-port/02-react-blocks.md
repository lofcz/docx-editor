# Production React block host

Build on examples/react-blocks. The proof remains evidence, rather than the production implementation.

Read the [dispatch and completion rules](../../../SCIOBOT_PORT_TODO.md) before claiming a leaf.

Each named test or artifact is a deliverable. These specifications do not claim those checks already exist.

## B01

- [x] Define a versioned block payload schema

**Depends:** [C03](01-contracts.md#c03)

**Primary files:** `fork:packages/sciobot/src/blocks/payload-schema.ts`; `sibling payload-schema.test.ts`

**Read:** [S09](sources.md#s09), [S10](sources.md#s10), [S11](sources.md#s11), [S24](sources.md#s24)

**Change:** Define block ID, type, version, attributes, and opaque payload fields. Preserve unknown fields without executing them.

**Close when:** Fixtures cover nested blocks, absent properties, unknown versions, and malformed payloads. Decoding never invents a valid replacement payload.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `payload-schema.test.ts` passes seven tests and 35 assertions for nested payloads, complete ASTs, missing properties, opaque future versions, malformed data, resource limits, inert unknown keys, and rejected non-JSON values. Tag IDs reserve the native 64-character tag budget. Package typecheck, scoped Prettier, and oxlint pass. Fixture: `block-payloads.json`.

## B02

- [x] Persist payloads in a DOCX custom XML part

**Depends:** [B22](02-react-blocks.md#b22), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/blocks/payload-part.ts`; `sibling payload-part.test.ts`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Replace alias JSON storage with an application custom XML part. Keep short block identifiers in content-control tags.

**Close when:** Save/reopen retains nested payloads, unrelated custom parts, relationships, and content types. Undo restores both payload and Word fallback.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Three focused cases in `payload-part.test.ts` pass: payload/fallback/tag update as one undo unit and save/reopen; atomic refusal for bad fallback, viewing mode, and stale reads; two real sessions with offline peer input, presence, undo/redo, peer deletion, reconnect, and equal saved digests. Payload plus schema/codec suite passes 15 cases. Package typecheck, scoped Prettier, and oxlint pass. The build boundary guard verifies eight private modules and external core. Tracked properties remain assigned to R12.

## B03

- [x] Preserve block IDs across save and reopen

**Depends:** [B02](02-react-blocks.md#b02), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/blocks/block-index.ts`; `sibling block-index.test.ts`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Resolve persistent application IDs to live canonical nodes. Handle missing, duplicate, and orphaned IDs without changing unrelated blocks.

**Close when:** Two identical-looking blocks retain distinct identities after reopen. Invalid references produce a bounded diagnostic.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `block-index.test.ts` passes five cases and 26 assertions. Identical-looking blocks retain distinct IDs through save/reopen; nested controls resolve parents; duplicate, missing, malformed, and orphan references produce bounded diagnostics without document writes. Two real sessions cover offline input, undo/redo, presence, removal during peer editing, reconnect, and equal saved digests. Package typecheck, scoped Prettier, and oxlint pass.

## B04

- [x] Extract the registered React block host

**Depends:** [C03](01-contracts.md#c03), [B03](02-react-blocks.md#b03), [B24](02-react-blocks.md#b24), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/react-blocks/{host,block-ui}.tsx`; `host.test.tsx`; `fork:e2e/sciobot-block-host.spec.ts` and its config

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Extract root creation, registration, paint observation, and disposal from the example. Keep canonical content authoritative.

**Close when:** A registered widget mounts once, updates after native edits, and unmounts after disposal. Unregistered content retains native rendering.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. The optional private React entry supplies registration, persistent per-fragment roots, native paint observation, parking, updates, and idempotent disposal. Four focused host tests pass, including two real sessions with both reconnect delivery orders, independent local state, canonical payload/native text synchronization, peer-preserving history, remote caret geometry, deletion, and saved digests. The built entry passes three Chromium cases with real native typing, arbitrary React form state, native page moves, lifecycle cleanup, saved DOCX reopen, and zero page errors. Unregistered content retains its native fallback. B24 fixes the separately reproduced imported-payload undo prerequisite. Core and private typechecks, browser fixture typecheck, external-runtime builds, API check, scoped format/lint, root lint, diff check, and compatibility policy pass. Build checks verify 45 private modules, one external core, and separate React/document entry points. Negative controls reject React in document tools and Pro in the React entry. Decision: `.collaboration/changes/sciobot-react-block-host.json`; no host shared-state changes. Full private suite: 336 pass and the same two R29 failures, across 338 tests and 36 files. B05/B06 retain providers and native slots; B14-B17 retain measurement and continuation policy. See [implementation notes](evidence/implementation-notes.md).

## B05

- [x] Supply application providers to block roots

**Depends:** [B04](02-react-blocks.md#b04), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/react-blocks/providers.tsx`; `sibling providers browser test`

**Read:** [S24](sources.md#s24), [S11](sources.md#s11)

**Change:** Add an explicit provider wrapper for each independent root. Refresh provider values without recreating widget identity.

**Close when:** A widget reads locale, application services, and live provider changes. Two editor instances do not share provider state.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Explicit per-host provider callbacks wrap every root through a stable React boundary. `setProviders` refreshes values without document writes or widget remounts. The callback receives editor, block, and fragment scope. Two real sessions with two native roots per editor pass both reconnect orders: two expanded cases and 86 assertions. Locale and application services update independently, local state and native children persist, shared payload/text edits survive history, and deletion during peer editing preserves the surviving context and saved convergence. The provider Chromium case checks live values, focused React input, native typing, exact model/DOM caret continuity, a second typed edit, two editor contexts, saved digests, and zero page errors. All six browser cases pass; the affected provider case reruns after waiting for buffered typing to paint. Private/browser typechecks, external-runtime build, scoped format/lint, root lint, diff check, and compatibility policy pass. Build checks verify 48 private source modules, a React-free document entry, and one external core. Full private suite: 342 pass and the same two R29 failures across 344 tests and 39 files. The expanded multiple-root assertions pass separately. No-impact decision: `.collaboration/changes/sciobot-block-providers.json`. See [implementation notes](evidence/implementation-notes.md).

## B06

- [x] Attach native content inside React containers

**Depends:** [B03](02-react-blocks.md#b03), [B04](02-react-blocks.md#b04), [B25](02-react-blocks.md#b25), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/react-blocks/native-slots.ts`; `sibling native-slots browser test`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Restore borrowed paragraph and table elements before repaint. Reattach them before native selection restoration.

**Close when:** Typing inside nested containers edits native content. Repaint retains selection and never duplicates native paragraph elements.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `NativeSlotLease` restores original paint nodes at exact comment placeholders before paint. Canonical ancestry assigns nested paragraph/table/row/cell ownership; existing slots reattach before native selection restoration. Transparent wrappers retain native coordinates. Eight focused host/slot cases pass; expanded two-session controlled-cell scenarios pass both reconnect orders with 294 assertions. They cover concurrency, peer-preserving undo/redo, remote selection geometry, deletion, and saved convergence. The built optional React entry passes five Chromium cases, including real typing in nested containers and cells, caret continuity across refresh and another edit, React state, unregister/register, disposal, no duplicate paragraphs, concurrent sessions, and saved DOCX reopen. No page errors occur. B25 closes the discovered wrapped-row boundary prerequisite. Private/core/browser typechecks, external-runtime builds, core API check, scoped format/lint, root lint, diff check, and compatibility pass. Build checks verify 47 private source modules and one external core. Full private suite: 340 pass and the same two R29 failures across 342 tests and 38 files. Provider inheritance, inline slots, measurement, and continuation sizing retain separate leaves. See [implementation notes](evidence/implementation-notes.md).

## B07

- [x] Preserve native range selection through wrappers

**Depends:** [B06](02-react-blocks.md#b06), [V01](11-verification.md#v01), [B26](02-react-blocks.md#b26)

**Primary files:** `fork:packages/sciobot/src/react-blocks/selection.ts`; `sibling selection browser test`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Map selection through nested and continued block fragments. Keep widget form selections separate from document selections.

**Close when:** Keyboard extension crosses a container boundary and a page boundary. Copy returns the selected document range.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `readBlockSelection` reads the flushed native model selection and maps it through canonical block ancestry. It returns document-ordered blocks with `within`, `crosses`, or `contains` relations and the native fragment ordinals the range reaches. Those ordinals match React root keys. `isWidgetSelectionNode` identifies widget furniture. `BlockUi` now carries the B26 `data-docx-host-ui` marker. Fixture: `blockSelectionDocument()` in `packages/sciobot/test/support/react-block-document.ts`, with a native container over three pages, a nested container, and a textarea widget. `e2e/sciobot-block-selection.spec.ts` passes three Chromium cases, twice in a row, on the built React entry. Real Shift+ArrowUp leaves the nested and outer containers. Real Shift+ArrowDown runs from page 0 to page 2, and then leaves the outer container. Each real Ctrl+C payload equals the canonical paragraph-text oracle for the model range. Real textarea focus, Ctrl+A, Shift+ArrowLeft, typing, and mouse drag keep the document range, package revision, and text unchanged. The two-session case checks remote selection geometry across pages, offline typing on both sides, reconnect, undo/redo, block deletion during peer editing, and saved digest convergence. The full browser config passes nine cases. Private react-block unit tests pass 10 cases. The private typecheck reports one error in another worker's `src/ai/direct-baseline.test.ts` and no errors in B07 files. The private build verifies 115 modules. Scoped Prettier and oxlint pass. `packages/sciobot/test/browser/vite.config.ts` adds `SCIOBOT_BROWSER_STABLE=1`: it disables HMR and file watching, so concurrent source edits cannot reload a page during a run. Compatibility: `.collaboration/changes/sciobot-block-selection.json`. See [implementation notes](evidence/implementation-notes.md).

## B08

- [x] Isolate widget form input events

**Depends:** [B06](02-react-blocks.md#b06), [V01](11-verification.md#v01), [B28](02-react-blocks.md#b28)

**Primary files:** `fork:packages/sciobot/src/react-blocks/input-boundary.ts`; `sibling input-boundary browser test`

**Read:** [S24](sources.md#s24)

**Change:** Route form beforeinput, keyboard, clipboard, and composition events to React controls. Route native content events to the editor.

**Close when:** Typing, paste, cut, and undo inside a textarea never modify Word text. Native text input still works.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `input-boundary.ts` owns `WIDGET_INPUT_EVENTS`, `isWidgetInputEvent`, and `isolateWidgetInput`. The host attaches it after `createRoot`, so React root handlers run first. Then the boundary stops bubbling for events from `BlockUi` only. Native slot content keeps the editor's routing. Each mount releases its listeners on disposal. `selectstart` joins the stopped events. B28 covers the capture-phase readers and the painted caret. `e2e/sciobot-block-input.spec.ts` passes two Chromium cases on the built React entry. Real textarea typing, clipboard paste, cut with clipboard read-back, Ctrl+Z, Ctrl+Shift+Z, Backspace, Enter, Ctrl+B, and a CDP IME composition change only the textarea. The document revision and text stay unchanged. The painted caret hides while the textarea has focus, and the textarea caret color is not transparent. Native typing then lands, and two document undo steps revert only the two document characters. The two-session case keeps widget typing local through a paused room, concurrent peer text, reconnect, focus retention, peer undo/redo, and peer deletion of the widget block during typing. Saved digests converge. Peer edits in the shared test page use automation operations, because a programmatic selection in the other editor moves the page focus. The full browser config passes 11 cases. Private react-block unit tests pass 10 cases. Private typecheck passes 1685 files, and the build verifies 118 modules. Scoped Prettier and oxlint pass. Compatibility: `.collaboration/changes/sciobot-block-input.json`. See [implementation notes](evidence/implementation-notes.md).

## B09

- [x] Regress composition event routing at the block boundary

**Depends:** [B08](02-react-blocks.md#b08), [V01](11-verification.md#v01)

**Primary files:** `fork:e2e/sciobot-block-ime.spec.ts`; `host fixes only when this test fails`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Exercise composed text in a native slot and a React input. Use browser composition event sequences; keep real desktop IME verification in B23.

**Close when:** Each composition commits once. Selection remains valid after repaint. Widget composition does not enter native Word content.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** astra B09, branch `codex/react-block-proof`. **Evidence:** Baseline `3278fa04`, local working tree. `e2e/sciobot-block-ime.spec.ts` passes three Chromium cases. The CDP sequence updates `か` to `かな`, then commits `仮名`. Native content becomes exactly `Inner text仮名`; one undo removes the composition and one redo restores it. A peer repaint retains the native caret at offset 12, and following keyboard input lands there. Remote selection geometry agrees across both real sessions. Both reconnect orders preserve concurrent peer text, history, and saved/reopened semantic digests. React textarea composition commits once, keeps both document revisions unchanged, and retains focus and offset 14 after repaint. The other textarea stays unchanged. Peer deletion during active widget composition removes the widget without leaking text into Word. No page errors occur. No host or production changes were needed; shared operations and formats stay unchanged. Command: `bunx playwright test --config e2e/.b09.config.ts`, using a temporary config on port 5197 with `SCIOBOT_BROWSER_STABLE=1`; 3/3 pass with unchanged 60-second test deadlines. Initial cold-server attempts timed out before assertions; a separate browser warm-up completed dependency optimization. Scoped Prettier and oxlint pass. Dedicated e2e typecheck passes 1531 files after an explicit registry Map type fixes existing harness inference; runtime behavior is unchanged. The temporary config is removed after verification; the shared block-host config includes this spec. This proves browser composition event routing only. Actual operating-system IME remains B23.

## B10

- [x] Copy and paste registered blocks

**Depends:** [B02](02-react-blocks.md#b02), [B03](02-react-blocks.md#b03), [B06](02-react-blocks.md#b06), [V01](11-verification.md#v01), [B29](02-react-blocks.md#b29)

**Primary files:** `fork:packages/sciobot/src/react-blocks/clipboard.ts`; `sibling clipboard browser test`

**Read:** [S11](sources.md#s11), [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Carry payloads with private clipboard data and native Word or HTML fallback. Assign fresh IDs and remap internal references on paste.

**Close when:** Copying a nested activity preserves attributes and text. Pasting twice creates distinct IDs. External plain-text paste remains native.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. The probe found that native fragment paste already carries block SDT tags, but the source IDs stay. The block index then reports `duplicate-reference` and drops both copies. `mountBlockClipboard` adds `application/x-sciobot-blocks+json` beside the native text and HTML flavours on copy and cut. It collects payloads for every block the B07 selection reaches and their registered descendants. A paste with that flavour or a native fragment runs `pasteWithBlocks`: native `pasteRich`, then `repairPastedBlocks`, as two frames of one explicit history group (B29). The repair gives each pasted control a fresh ID, writes its payload record, and remaps string references between pasted blocks. Pasted tags without a known payload are cleared, so their controls keep native content only. Other pastes stay on the native handler. `readBlockClipboardData` bounds size, entry count, labels, and payload schema. `e2e/sciobot-block-clipboard.spec.ts` passes two Chromium cases twice in a row, with real keyboard selection, Ctrl+C, Ctrl+X, Ctrl+V, Ctrl+Z, and Ctrl+Shift+Z. A nested activity copy keeps `Warmup`/`10`, the nested note, and its `activityRef`, which follows the new activity ID. Two pastes create distinct IDs with no index diagnostics. One undo removes one whole paste and its records, and save/reopen keeps the same IDs. External plain text pastes natively with one revision and no block change. Cut carries the payloads before the native delete. The two-session case covers concurrent pastes in both sessions, peer text, reconnect, convergent block indexes, paste undo/redo that keeps the peer paste and text, and source deletion during a peer paste. Saved digests converge. The full browser config passes 13 cases. Private typecheck passes 1703 files. Scoped Prettier and oxlint pass. Compatibility: `.collaboration/changes/sciobot-block-clipboard.json`. Limit: the payload flavour and repair work inside one application. Word and other applications receive the native fallback only. See [implementation notes](evidence/implementation-notes.md).

## B11

- [x] Insert a registered container block

**Depends:** [B02](02-react-blocks.md#b02), [B03](02-react-blocks.md#b03), [B06](02-react-blocks.md#b06), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/blocks/insert.ts`; `sibling insert.test.ts`

**Read:** [S09](sources.md#s09), [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Insert one registered container with payload, native fallback, and a caret target in one store transaction.

**Close when:** Insert an activity into a phase. One undo removes payload and fallback. Redo restores both and a valid caret.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** claude (worker B2), branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `src/blocks/insert.ts` exports `insertBlock`. In one gated `applyAutomationOps` package-edit transaction with its own undo unit, it places a registered container (alias, `scio:1:` tag, optional `scio:label` generated label, one or more editable paragraphs) in the body or inside a registered parent's `w:sdtContent`, before a named child or at the end (before section properties), and writes its payload record. It then places the caret in the first editable paragraph. New paragraphs receive `w14:paraId` identities through the published minting helpers, seeded by the block ID. Collaboration renames hook-created nodes during the commit, so the result and caret are re-read by tag after commit; a caret built from pre-commit IDs pointed at a removed node. Atomic refusals, each leaving the canonical fingerprint unchanged: existing or invalid IDs, a missing, locked, data-bound or inline parent, an anchor outside the parent, empty or multiline paragraphs, stale revision, suggestion mode (`unsupported-operation`; tracked insertion stays with review executors), and view mode. `insert.test.ts` passes 8 cases. An activity inserted into a phase carries payload, alias, generated label and caret; typing at the caret lands in the block; one undo removes block, fallback and payload together (fingerprint and store equal the original); redo restores both with a caret that `caretAt` resolves; save/reopen indexes the block under `phase-2` with its payload. Body insertion lands before a sibling and at the end. Refusals are covered. Two real sessions in both client orders (73/41, 41/73) and both receivers insert offline beside peer typing in the same phase, then converge with payload, peer text and saved digests. Alice's typing in the new block shows as her remote caret at a resolvable geometry on Bob's layout, and undo of typing and insertion preserves Bob's text, then redo restores the block. A peer deleting the parent phase during insertion converges. All 48 blocks-suite cases pass. Private typecheck is clean in owned files; the remaining errors are in in-progress `src/review` files. The private build verifies 127 modules; scoped oxlint/prettier pass; the compatibility policy passes with 79 decisions. Decision: `.collaboration/changes/sciobot-block-insert.json` (no-impact). Exported from the package entry. Delete/unwrap and movement remain B12/B13.

## B12

- [ ] Delete or unwrap a registered block

**Depends:** [B11](02-react-blocks.md#b11), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/blocks/remove.ts`; `sibling remove.test.ts`

**Read:** [S09](sources.md#s09), [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Implement explicit delete and unwrap commands. Preserve children during unwrap and remove only owned metadata during deletion.

**Close when:** Unwrap retains native text. Delete removes its payload. Undo restores nested descendants and attributes exactly.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## B13

- [ ] Move a registered block between containers

**Depends:** [B11](02-react-blocks.md#b11), [B12](02-react-blocks.md#b12), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/blocks/move.ts`; `sibling move.test.ts`

**Read:** [S09](sources.md#s09), [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Move a block using native store operations. Keep block IDs, payload references, and child ordering stable.

**Close when:** Move an activity between phases. One undo returns its original position. Reject moves into the block itself.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## B14

- [x] Measure widget height without undo noise

**Depends:** [B04](02-react-blocks.md#b04), [B06](02-react-blocks.md#b06), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/react-blocks/measurement.ts`; `sibling measurement browser test`

**Read:** [S24](sources.md#s24)

**Change:** Reserve native height from measured widget geometry. Separate derived measurements from user history and prevent resize feedback loops.

**Close when:** Expanding a widget moves following native text once. Undo reverses the user edit without extra measurement steps.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex, branch `codex/react-block-proof` (continuing Claude). **Evidence:** Engine baseline `9770e78f` plus local work. The inherited `BlockMeasurement` implementation measures intrinsic widget height, accounts for zoom, and publishes layout-only reservations. Native height excludes reservations, so resize settles without feedback or extra history. `bun test --preload ./scripts/test/preload-dom.ts --timeout 30000 packages/sciobot/src/react-blocks/measurement.test.tsx packages/core/src/editor/__tests__/block-reservations.test.ts`: 12 pass, 211 assertions; `/tmp/codex-docx-measurement-unit.log`. `SCIOBOT_BROWSER_STABLE=1 bunx playwright test --config e2e/sciobot-block-host.config.ts e2e/sciobot-block-measurement.spec.ts --timeout 180000`: 3 pass; `/tmp/codex-docx-measurement-browser-final.log`. Real button clicks grow the widget, move following text once, and undo one user change per step. Real native typing lands on the moved line. Two real sessions cover both reconnect receivers, offline growth beside peer text, matching remote caret geometry, deletion during growth, undo/redo, and saved convergence. Private/core typecheck, private/core builds, core API check, scoped format/lint, and compatibility policy pass. The only continuation fix in this area removes an impossible textValue test on a typed element root.

## B15

- [x] Retain widget state through repagination

**Depends:** [B06](02-react-blocks.md#b06), [B14](02-react-blocks.md#b14), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/react-blocks/host.tsx`; `fork:packages/sciobot/test/browser/fragments.tsx`; `fork:e2e/sciobot-block-fragments.spec.ts`

**Read:** [S24](sources.md#s24)

**Change:** Key fragment roots by block identity and continuation identity. Move roots between page elements without recreating local state.

**Close when:** A counter and focused input survive a page move. A native container continues across three pages without duplicated content.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** astra B15, branch `codex/react-block-proof`. **Evidence:** Baseline `3278fa04`, local working tree. Existing `host.tsx` keys roots by block ID and continuation ordinal; no production change was needed. New `packages/sciobot/test/browser/fragments.tsx` combines measured React widgets with nested native containers and two real sessions. `e2e/sciobot-block-fragments.spec.ts` passes both reconnect orders. Exact expected paragraph arrays verify three native continuation pages without duplicated content. A native spacing edit moves the widget between pages while preserving its input DOM node, focus, caret offset 10, local counter, and text. Following keyboard input lands at that caret. Reconnect retains focus and combines peer text with the page move. Real document Ctrl+Z/Ctrl+Shift+Z reverse and restore pagination while preserving peer text and widget state. Document history intentionally restores the document caret; focus retention is asserted for page moves and reconnect. Remote selection geometry matches in both sessions. Peer deletion during local widget typing unmounts each root once, preserves remaining native paragraphs, and converges after save/reopen. Command: `bunx playwright test --config e2e/.b15.config.ts`; 2/2 pass in 16.9 seconds with unchanged 60-second deadlines. The temporary config uses stable port 5197 and is removed after verification. The shared block-host config registers this spec. Scoped formatting and oxlint pass. The final combined browser typecheck passes 1533 files against isolated completed-leaf checkpoint `988042a5`; `/tmp/codex-closed-browser-types2.log`. No shared operations or formats change; no new compatibility decision is needed.

## B16

- [ ] Release parked roots after page and editor disposal

**Depends:** [B15](02-react-blocks.md#b15), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/react-blocks/retention.ts`; `sibling retention browser test`

**Read:** [S24](sources.md#s24)

**Change:** Park roots during page virtualization. Release roots, observers, and listeners when blocks or the editor disappear.

**Close when:** Twenty scroll and document-switch cycles return root and observer counts to baseline. Parked widgets resume with correct state.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## B17

- [ ] Handle widgets taller than one page

**Depends:** [B14](02-react-blocks.md#b14), [B15](02-react-blocks.md#b15), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/react-blocks/overflow.ts`; `sibling overflow browser test`

**Read:** [S24](sources.md#s24), [S12](sources.md#s12)

**Change:** Use a bounded scroll region for oversized opaque widgets. Preserve the complete native fallback for paginated Word output.

**Close when:** A long listening exercise remains fully reachable and editable. Native following content does not overlap. Export retains the full transcript.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## B18

- [x] Define inline widget measurement contracts

**Depends:** [B01](02-react-blocks.md#b01), [C07](01-contracts.md#c07)

**Primary files:** `fork:packages/core/src/contracts/inline-widget-metrics.ts`; `contract tests and API snapshot if exported`

**Read:** [S10](sources.md#s10), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Define renderer-neutral width, ascent, descent, and native fallback rules for registered inline widgets. Keep React outside layout.

**Close when:** Contract fixtures represent a source chip and a date chip. The unregistered path keeps native Word metrics.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** claude (worker B2), branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `packages/core/src/contracts/inline-widget-metrics.ts` defines a pure, import-free contract: `InlineWidgetMetrics` (width, ascent, descent in CSS px at 100% zoom, relative to the baseline), `InlineWidgetRequest` (control ID, tag, fallback text, available line width), an `InlineWidgetMetricsProvider` that returns null for controls it does not own or has not measured, and `resolveInlineWidgetBox`. Rules: a valid answer is quantized to the layout's 1/1000 px fixed-point grid, half to even (the shaped-advance rounding), so equal measurements give equal boxes in every session. A box wider than the line is clamped to it and marked `clamped`. No provider, a declining provider, or a not-yet-measured widget keeps native Word metrics (`native`/`unregistered`), so nothing guessed is reserved. Non-finite, non-positive width or ascent, negative descent, non-numeric, or over-bound (2112 px) answers keep native metrics as `invalid-metrics`. `contracts/__tests__/inline-widget-metrics.test.ts` passes 10 cases with explicit expected boxes. Source chip: `{118.4996, 15.0004, 4.25}` resolves to `{118.5, 15, 4.25}`. Date chip: `{84.0625, 13.4, 3.4}` resolves to `{84.062, 13.4, 3.4}` (half-even). Line clamping, a zero-descent box, unregistered and untagged controls, no provider, a pending measurement, nine invalid answers, the quantization grid, and an import-free source check are also covered. Core typecheck passes (1467 files); scoped oxlint and prettier pass; the lane boundary check passes; the compatibility policy passes with 84 decisions. The module is internal and unexported, so the API snapshot is unchanged. Changeset: `.changeset/inline-widget-metrics-contract.md`; decision: `.collaboration/changes/inline-widget-metrics-contract.json` (no-impact). B19 consumes it at the layout boundary; B20 supplies host measurements; a public export, if B20 needs one, carries its API snapshot then.

## B19

- [ ] Reserve inline widget geometry in layout

**Depends:** [B18](02-react-blocks.md#b18), [B02](02-react-blocks.md#b02), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/layout/inline-widget-layout.ts`; `focused layout and paint tests`

**Read:** [S25](sources.md#s25)

**Change:** Consume inline metrics through the permitted layout boundary. Reserve boxes without changing the canonical fallback text.

**Close when:** A chip wraps at its measured width. Adjacent native text and caret geometry match explicit box coordinates.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## B20

- [ ] Mount React inline widgets in native boxes

**Depends:** [B04](02-react-blocks.md#b04), [B19](02-react-blocks.md#b19), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/react-blocks/inline-host.tsx`; `sibling inline-host browser test`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Mount registered inline views into measured native boxes. Reuse provider and disposal handling from the block host.

**Close when:** A source chip renders between native words. Zoom, repaint, and provider updates retain identity and correct hit geometry.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## B21

- [ ] Edit inline widgets as native atoms

**Depends:** [B20](02-react-blocks.md#b20), [B08](02-react-blocks.md#b08), [B10](02-react-blocks.md#b10), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/react-blocks/inline-input.ts`; `sibling inline-input browser test`

**Read:** [S10](sources.md#s10), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Connect arrow navigation, selection, deletion, copy, and paste to the inline content-control boundary.

**Close when:** Backspace deletes one chip. Undo restores its payload. Copy/paste retains attributes with a fresh ID and no split atom.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## B22

- [x] Encode payload parts and content-control references

**Depends:** [B01](02-react-blocks.md#b01)

**Primary files:** `fork:packages/sciobot/src/blocks/payload-opc-codec.ts`; `sibling payload-opc-codec.test.ts`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Encode and decode application custom XML with bounded payload sizes and stable tag references. Preserve unrecognized fields.

**Close when:** Nested payload fixtures roundtrip without alias JSON. Duplicate references and malformed payload parts produce explicit decode results.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `payload-opc-codec.test.ts` passes five tests and 23 assertions. Native custom XML parts retain complete nested data through save/reopen. Tags fit the 64-character native limit. Duplicate references, duplicate/malformed records, mismatched IDs, and invalid XML labels refuse. Foreign stores and opaque future payloads stay intact. Package typecheck, scoped Prettier, and oxlint pass. This is the immutable package codec; live atomic writes remain B02.

## B23

- [ ] Verify native IME on a supported desktop

**Depends:** [B09](02-react-blocks.md#b09)

**Primary files:** `fork:docs/engineering/sciobot-port/evidence/desktop-ime.md`; `captured regression sequence`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Run actual operating-system IME composition in a native slot and a React input. Record browser, operating system, and input method.

**Close when:** Both inputs commit composition once. Repagination retains selection. Record the actual input sequence; synthetic events alone do not close this leaf.

**Check:** Use desktop Microsoft Word. ZIP inspection or a browser screenshot cannot substitute for this check.

**Claim:** unclaimed. **Evidence:** pending.

## B24

- [x] Preserve imported payload node identities during updates

**Depends:** [B22](02-react-blocks.md#b22), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/package/custom-xml-nodes.ts`; native custom XML regression and private two-session host tests

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Update existing custom XML payload values without replacing imported node identities. Preserve unrelated attributes, children, and records.

**Close when:** Payload undo restores the original valid store in both sessions. No empty replacement node survives. Redo restores the new value and peer text survives.

**Check:** Check imported wrapper and text identities, unknown XML preservation, save/reopen, two-session undo/redo, reconnect, and peer deletion. Include the core changeset, API checks, and compatibility decision.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Imported payload updates previously replaced canonical node identities. The two-session host regression reproduced an empty replacement node surviving undo beside the restored original, causing strict missing-node-id failures. `withCustomXmlNode` now retains imported wrapper, label/data, and text identities, plus foreign attributes, siblings, and unrelated records. Native identity/custom XML tests pass 26 cases and 67 assertions; binding, baseline identity, and tombstone regressions pass 33 cases and 85 assertions. Both real-session reconnect orders preserve old/new valid stores through undo/redo, peer text, removal, and saved convergence. Three built-entry Chromium cases also pass. Core typecheck/build and its 15-entry API check pass without new snapshot changes. Private checks, lint, formatting, diff checks, and compatibility policy pass. Core patch changeset: `.changeset/custom-xml-node-identities.md`; compatible writer decision: `.collaboration/changes/custom-xml-node-identities-compatibility.json`. Provider reading rules and shared formats stay unchanged. See [implementation notes](evidence/implementation-notes.md).

## B25

- [x] Publish boundaries for controls around table rows

**Depends:** [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/layout/content-control-boundary-layout.ts`; `sibling content-control-row-boundaries.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Visit row controls and their cell content during native boundary collection. Preserve geometry and nested lock stacks.

**Close when:** A wrapped row publishes row, cell, and nested block boundaries in document order. Save/reopen preserves the records. Two sessions edit the controlled cell without lost or duplicated content.

**Check:** Run native boundary and table regressions. Check the controlled cell in the two-session native-slot fixture. Include the core changeset, typecheck, build, API check, and compatibility decision.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Boundary collection now visits wrapped rows and their cell content. The native row-boundary regression proves document order, inherited locks, unchanged table geometry, nested records, and serialization. Boundary/table coverage includes 41 cases across six files; the isolated seven-case table-fact file passes with a 30-second command timeout after resource-related five-second timeouts. Both real-session delivery orders pass controlled-cell edits, concurrent sibling-cell edits, peer-preserving undo/redo, deletion, and saved digests. Two built-entry Chromium cases also exercise nested native slots and controlled-cell input. Core typecheck/build and the 15-entry API check pass without new snapshot changes. Scoped format/lint, root lint, and compatibility policy pass. Core patch changeset: `.changeset/content-control-row-boundaries.md`; no-impact decision: `.collaboration/changes/sciobot-native-slots.json`. No shared format changes. See [implementation notes](evidence/implementation-notes.md).

## B26

- [x] Keep host widget selections outside native selection sync

**Depends:** [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/editor/dom-selection.ts`; `fork:packages/core/src/editor/surface-selection-sync.ts`; `fork:packages/core/src/editor/surface-pointer.ts`; `sibling host-ui-selection.test.ts`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Recognize host-owned UI painted inside pages. Its presses, focus, and DOM selections must not arm, adopt, collapse, or overwrite the native document selection.

**Close when:** With a document range active, a press, select-all, and typing in a widget textarea keep the textarea focused. The model range and document text stay unchanged. Unmarked content keeps the existing collapse behavior.

**Check:** Add a focused native regression and run the selection, pointer, caret, and focus suites. Include the core changeset, typecheck, and compatibility decision. Verify the browser case through B07.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. The B07 Chromium probe found a defect. A press on a widget textarea armed a selection gesture. `selectionchange` then reported the widget container, which maps to no model position. The mirror collapsed the document range to its head and wrote the DOM selection back. Chromium moved focus to the pages, so Ctrl+A and typing replaced the whole document with `X`. `data-docx-host-ui` now marks host UI. `isInHostUi` and `domSelectionInHostUi` read the marker for text and element nodes. `adoptDomSelection` ignores those selections. `ownsSelection` refuses DOM writes while host UI holds focus or the DOM anchor. The selection gesture and `surface-pointer` press handlers skip host UI targets. `packages/core/src/editor/__tests__/host-ui-selection.test.ts` passes three cases and 12 assertions, including an unmarked control case. The focused selection, pointer, caret, and focus suites pass 415 cases across 37 files. Core typecheck passes 1459 files. The marker is not a new exported symbol, so API snapshots stay unchanged. The content-controls guide documents the attribute. Core minor changeset: `.changeset/host-ui-selection.md`. No-impact decision: `.collaboration/changes/sciobot-block-selection.json`. Compatibility policy passes with 60 decisions. No shared format or transport data changes.

## B27

- [x] Converge concurrent payload writes to one block

**Depends:** [B02](02-react-blocks.md#b02), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/blocks/payload-part.ts`; `fork:packages/sciobot/src/blocks/payload-opc-codec.ts`; native custom XML replication path that publishes the payload record; `sibling payload-concurrency-collaboration.test.ts`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Make concurrent `writeBlockPayload` commits to the same block converge in every session. Merge disjoint attribute changes when possible; otherwise apply one deterministic winner in all sessions. Never leave sessions with different payloads after reconnect.

**Close when:** Offline sessions write different attributes of one block, for example `grade` and `lessonMinutes`. After reconnect in both client orders and both delivery orders, every session holds the same payload, and disjoint attributes are both retained. Undo/redo of either write and save/reopen keep that convergence. Writes to different blocks still merge as they do today.

**Check:** Add the focused two-session regression in both client orders. Re-run the payload, provider, and host collaboration suites. Record the compatibility decision for any change to the shared payload format or replication.

**Claim:** claude (worker B2), branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Reported by X01 (app branch `claude/docx-native-editor`): with `writeBlockPayload` in both offline sessions, `reconnect(0)` left Alice with only `grade` and Bob with only `minutes`, in both client orders. There were two causes. First, core `setContentControlProperties` replaced `w:tag`/`w:alias` with fresh elements, so concurrent writers merged into a control with two of each. Every remote apply then failed `known-node-invariant` and both sessions entered `error`. The shared Yjs state had in fact converged. Second, payload updates rewrote the record's JSON `data` text, which shared text merges character by character: same-key writes garble the value, and remove-versus-change produces unparsable JSON that invalidates the whole store. Fix: core `editedProperties` now updates an existing property's `w:val` in place and writes nothing for an unchanged value. `writeBlockPayload` omits unchanged tag/alias and never adds a missing alias (created blocks carry one). New `src/blocks/payload-fields.ts`: updates to an existing record leave `data` intact and write per-field registers, `<attr name value>` and `<field name value>` XML attributes on child elements of the record node. Readers apply them over the created JSON. Disjoint fields merge; one value wins per field in every session; a valueless register removes a key; duplicate registers resolve to the last in document order and the next write removes the rest. Record creation and register-free records are unchanged, so all `withPayloadRecord` callers (B10 clipboard, R12 decisions, A01 direct, D29 seeds) benefit without changes. `payload-concurrency-collaboration.test.ts` passes 13 real-session cases, covering both client orders (73/41, 41/73) and both receivers. Disjoint `grade`/`lessonMinutes` writes merge, then undo/redo of either write, save/reopen, and a reopened-room write all converge. Same-key conflicts take one register winner, never a character merge, and the next write leaves one register. Remove-versus-change converges. Writes to different blocks still merge, with fallback text. A peer deleting the block during a write converges. An existing alias stays one property with one winner, and a missing alias is never added. Every case checks saved digests and non-error status. Negative controls: disabling registers fails 6 of 13 cases; reverting the core in-place write fails the alias case. `payload-fields.test.ts` passes 7 pure cases: data retention, register identity, tombstones, top-level fields, key-order no-op, duplicate winners, malformed registers refusing with named reasons, and unchanged register-free records. The core `content-control-property-identity.test.ts` passes 3 cases. Regressions: 363 core content-control/automation/journal-coverage tests and 17 Pro custom-node/content-control collaboration tests pass, and the full private suite passes 1212 of 1212 tests across 105 files. Core and private typecheck pass (1464/1711 files). The private build verifies 126 modules. Scoped oxlint/prettier pass. The compatibility policy passes with 77 decisions. Core patch changeset: `.changeset/content-control-property-identity.md`. Decisions: `.collaboration/changes/content-control-property-identity.json` and `.collaboration/changes/sciobot-payload-field-registers.json` (compatible: private register elements inside the existing record; no migration). Limits: the record label and the native fallback text remain shared text, so concurrent label changes can merge character by character until the next write rewrites them. The application harness lesson-title pair scenario should be re-run by X01. See [implementation notes](evidence/implementation-notes.md).

## B28

- [x] Keep host widget input out of capture-phase document gestures

**Depends:** [B26](02-react-blocks.md#b26)

**Primary files:** `fork:packages/core/src/editor/surface-drawing-gestures.ts`; `fork:packages/core/src/editor/surface-caret.ts`; `fork:packages/core/src/editor/surface-timing.ts`; `sibling host-ui-input.test.ts`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Make capture-phase key readers ignore host UI targets. Hide the painted caret and restore the native caret color while host UI holds focus.

**Close when:** Backspace in a host UI textarea keeps a selected picture and is not prevented. The same press without the marker still reaches the picture reader. The painted caret yields to host UI focus and returns with document focus.

**Check:** Add a focused native regression that fails without the guards. Run the selection, pointer, caret, focus, drawing, image, and textbox suites. Include the core changeset, typecheck, and compatibility decision.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. The pages layer reads keydown and beforeinput in its capture phase. A host cannot stop that phase on its own container. With a picture selected, Backspace in a widget textarea deleted the picture and prevented the textarea edit. `onDrawingKeyGesture` and `timedPress` now ignore host UI targets. `surface-caret.ts` treats host UI focus as a blur, which removes the painted caret and the transparent `caret-color`. `packages/core/src/editor/__tests__/host-ui-input.test.ts` passes three cases. With the guards removed, two of them fail, as expected. The unmarked control case proves the capture-phase leak. The focused selection, pointer, caret, focus, drawing, image, textbox, and timing suites pass 577 of 578 cases across 57 files. The one failure was a load timeout in `image-overlay-frame.test.ts`, which passes separately with a 60-second timeout. Core typecheck passes 1460 files. API snapshots do not change. Core patch changeset: `.changeset/host-ui-input.md`. No-impact decision: `.collaboration/changes/sciobot-block-input.json`. The content-controls guide documents the host responsibility to stop bubbling events. Bubble-phase handlers in `paginated-surface.ts` stay unchanged; that file sits two lines below its line cap.

## B29

- [x] Let a rich paste extend an explicit history group

**Depends:** [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/store/tree-package-fragment.ts`; `fork:packages/core/src/editor/editor-history-group-frames.ts`; `sibling history-group-frames-paste.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Carry the commit's history group into the fragment paste package unit and journal. Give adapters internal frames that bind one open group around several native writes.

**Close when:** A rich paste and a following automation write inside the frames undo and redo as one step. Without frames they remain two steps. Closed handles and handles from another editor refuse frames.

**Check:** Add a focused native regression. Run the clipboard, paste, fragment, and history suites. Include core typecheck, build, API check, line caps, the changeset, and the compatibility decision.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `applyFragmentPaste` promoted its story transaction to a package unit with no group, so the B10 repair write became a second undo step. One undo exposed duplicate block tags. `FragmentPasteInput.historyGroup` now reaches `promoteStoryTransactionToPackageUnit` and `runObservedStoreTransaction`. The surface supplies `commitHistoryGroup.value`. `beginHistoryGroupFrames(editor, group)` wraps the existing automation frames and is exported as `@internal`. `packages/core/src/editor/__tests__/history-group-frames-paste.test.ts` passes three cases. The clipboard, paste, fragment, and history suites pass 488 of 491 cases across 49 files under load. The three failures were timeouts. `clipboard-paste-router.test.ts` and `clipboard-rich-surface.test.ts` then pass 24 of 24 with a 180-second timeout. Core typecheck passes 1463 files, and the core build passes. The core API check reports 15 entries and zero errors, with the unchanged 2443 warnings and 380 forgotten exports. The snapshots include the new optional members. Line caps pass; `paginated-surface.ts` holds 6299 of 6300 lines. Core minor changeset: `.changeset/history-group-fragment-paste.md`. Decision: `.collaboration/changes/sciobot-block-clipboard.json`. The existing journal history-group field carries the group, so shared formats do not change. The B10 two-session case verifies collaborative undo of a grouped paste.
