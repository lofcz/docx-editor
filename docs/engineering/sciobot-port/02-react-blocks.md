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

- [x] Delete or unwrap a registered block

**Depends:** [B11](02-react-blocks.md#b11), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/blocks/remove.ts`; `sibling remove.test.ts`

**Read:** [S09](sources.md#s09), [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Implement explicit delete and unwrap commands. Preserve children during unwrap and remove only owned metadata during deletion.

**Close when:** Unwrap retains native text. Delete removes its payload. Undo restores nested descendants and attributes exactly.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `8326ab22`, completed local source. The private `removeBlock` command explicitly selects delete or unwrap. One gated native undo unit removes the wrapper and only its owned payload records. Deletion includes registered descendants; unwrap retains native children and descendant metadata. Native locks and mode gates remain authoritative. Shared control IDs, stale revisions, missing targets, suggesting mode, and viewing mode refuse without package changes. Suggestion removal stays with the review executors.

The focused native regression compares complete Word structure with literal expected XML while retaining original paragraph identities. Unwrap preserves the selected descendant and its exact position. Deletion leaves a resolvable native caret. Removing the only body block now creates an editable native paragraph through the existing insertion operation. The empty-body case survives one-step undo/redo and exact save/reopen. Foreign custom XML, unrelated payloads, and existing orphan metadata remain exact. A duplicate descendant ID outside the target refuses before writes.

Sixteen real two-session cases cover delete and unwrap, both client orders, both reconnect receivers, and ordinary or single-block documents. Removal runs offline beside peer typing, payload edits, and label edits in a nested note. Deletion wins while active. Undo restores the native descendants with the peer's latest text, attributes, label, and control identity. Unwrap retains that content throughout. Native remote selections resolve after restoration. Redo and complete saved semantic digests converge. The focused command passes 22 cases and 491 assertions before the final caret checks. Log: `/tmp/b12-peer-focused.log`. The transport remains test-only under the EigenPal Pro License.

`bun test packages/sciobot/src/blocks packages/sciobot/src/react-blocks` passes 133 cases and 3177 assertions across 22 files. Log: `/tmp/b12-affected.log`. Private types pass 2037 files; browser types pass 1757. The private build verifies 207 modules with external core and no production Pro imports. Root formatting, root lint, scoped formatting, and diff checks pass. Compatibility checks accept the new decision with zero changed format fields. Decision: `.collaboration/changes/sciobot-block-removal.json`. No published package or public API changes. The production audit and backlog validator verify closure. Container movement remains B13; block menu wiring remains U03.

## B13

- [x] Move a registered block between containers

**Depends:** [B11](02-react-blocks.md#b11), [B12](02-react-blocks.md#b12), [B32](02-react-blocks.md#b32), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/blocks/move.ts`; `sibling move.test.ts`

**Read:** [S09](sources.md#s09), [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Move a block using native store operations. Keep block IDs, payload references, and child ordering stable.

**Close when:** Move an activity between phases. One undo returns its original position. Reject moves into the block itself.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `ddb95987`, completed local source. `moveBlock` places a registered block before a native destination child or at the container end. It supports registered parent controls and the document body. The command retains the block tag, descendant IDs, payload records, nested references, and native ordering. Same-position moves return without history. Cycles, missing anchors, stale revisions, inline destinations, native locks, suggesting mode, and viewing mode refuse atomically. Existing single-control revision wrappers travel with their control. The command never disables tracking.

The focused command `bun test packages/core/src/store/__tests__/move-node.test.ts packages/sciobot/src/blocks/move.test.ts` passes 17 cases and 199 assertions. Log: `/tmp/b13-focused-final.log`. Native expected ordering is explicit, and the moved subtree remains the same object. Complete metadata stays exact. One undo restores the original package; redo restores the moved package. Save/reopen retains complete semantic structure. Four real-session cases cover both client orders and reconnect receivers. Movement runs offline beside nested peer text and payload edits. The moved native subtree retains the peer edits exactly. Undo restores the entire peer-only package on both sessions. Redo, remote caret identity/geometry, and saved digests converge. Four additional shared cases converge source and destination deletion against offline movement. A shared namespace case retains opaque metadata meaning through history and Word reopen. Test transport uses the EigenPal Pro License.

The affected block, native edit, namespace, journal, and isolation run covers 177 cases across 18 files. It reports 172 passes and five isolation-scan timeouts during concurrent root formatting and builds. With formatting finished, all 12 unchanged isolation cases pass at their original limits. Logs: `/tmp/b13-affected.log`, `/tmp/b13-isolation-final.log`. This is combined evidence; the initial affected process exits nonzero. Core types pass 1583 files; private types pass 2039; browser types pass 1757. Core and private builds pass. The private boundary verifies 208 modules and external core. Root formatting, root lint, scoped formatting, and diff checks pass. Public native API and compatibility evidence appears under B32. UI drag handles and movement menus remain U03/U04.

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

- [x] Release parked roots after page and editor disposal

**Depends:** [B15](02-react-blocks.md#b15), [B33](02-react-blocks.md#b33), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/react-blocks/retention.ts`; `fork:e2e/sciobot-block-retention.spec.ts`

**Read:** [S24](sources.md#s24)

**Change:** Park roots during page virtualization. Release roots, observers, and listeners when blocks or the editor disappear.

**Close when:** Twenty scroll and document-switch cycles return root and observer counts to baseline. Parked widgets resume with correct state.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `c7629552`, completed local source. Native paint continues to park offscreen roots while retaining their local React state. B33 now ends every host lifetime when its surface is destroyed or replaced. Block and inline hosts release roots, slots, providers, measurements, and subscriptions. Inline input releases document listeners and its clipboard lease. Shared clipboard ownership clears client callbacks and removes container listeners. Explicit disposal remains idempotent. Hosts use the native lifecycle callback without timers or DOM polling.

The focused Chromium command `SCIOBOT_BROWSER_STABLE=1 bunx playwright test --config e2e/sciobot-block-retention.config.ts` passes three cases. Twenty real scroll cycles retain the exact input DOM node, local value, and counter. Twenty document switches recreate two real sessions without manually disposing their hosts. Root counts, active ResizeObservers, observed elements, and widget-listener roots return to zero between documents and to the same baseline after remounting. Final editor destruction returns every count to zero. Instrumentation counts actual React effect lifetimes, ResizeObserver observations, and widget event subscriptions. Log: `/tmp/b16-browser-final.log`.

Both reconnect receivers run the shared scenario. A parked widget resumes after offline peer typing with its local form state. Deletion races another peer edit. The deleted widget releases its roots and listeners. Undo creates fresh widget state and restores the latest peer text; redo releases it again. A second undo and saved complete semantic digests converge. Native remote selection geometry remains resolvable. The first browser attempt retained the widget's page because its native caret pinned that page. The final fixture moves the caret to the other end before scrolling, which exercises actual virtualization. Its timeout and retention assertions remain explicit. The transport remains test-only under the EigenPal Pro License.

The affected native/window/React suite passes 74 cases and 1376 assertions across 15 files. A later focused command verifies both twenty-cycle block and inline replacement scenarios, passing two cases and 103 assertions. Logs: `/tmp/b16-affected.log`, `/tmp/b16-inline-retention.log`. Native types pass 1583 files; private types pass 2041; browser types pass 1762. Core and private builds pass; the private boundary verifies 209 modules and external core. Root formatting, root lint, line caps, scoped formatting, and public API checks pass. B33 records the public lifecycle hook. Decision: `.collaboration/changes/sciobot-react-retention.json`; no shared format fields change. Minor changeset: `surface-paint-observer-lifetime`.

Existing inline clipboard regressions retain their original assertions and limits. Two cases pass in `/tmp/b16-inline-input-regression.log`; that run later records a context timeout and is explicitly stopped. A standalone attempt then fails server startup before any test runs. A warm-server run passes the two remaining client-order cases but times out during first navigation of the other case. Its trace records 61.5 seconds in page loading. That exact case passes unchanged in isolation: `/tmp/b16-inline-warm.log`. The combined evidence covers all five cases without claiming the earlier processes passed. The inline-host command covers its six existing provider, geometry, typing, deletion, shared history, and reopen cases in `/tmp/b16-host-recheck.log`.

## B17

- [x] Handle widgets taller than one page

**Depends:** [B14](02-react-blocks.md#b14), [B15](02-react-blocks.md#b15), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/react-blocks/overflow.ts`; `sibling overflow browser test`

**Read:** [S24](sources.md#s24), [S12](sources.md#s12)

**Change:** Use a bounded scroll region for oversized opaque widgets. Preserve the complete native fallback for paginated Word output.

**Close when:** A long listening exercise remains fully reachable and editable. Native following content does not overlap. Export retains the full transcript.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `bddc9894`; completed local source. Measured opaque registrations keep one persistent root on the first native fallback fragment. Its scroll viewport fits the remaining page content at the current zoom. Only excess bounded height reserves native flow after the complete fallback. Native slots and inline widgets keep their separate adapters. Viewport furniture isolates React input from native editing. Canonical fallback paragraphs remain unchanged by scrolling and measurement.

The focused Chromium fixture uses ninety listening turns across multiple native pages. Real input reaches and edits the final turn. Save/reopen checks every expected transcript paragraph. Both reconnect delivery orders cover concurrent transcript and following-text edits, history, native typing after scrolling, remote caret geometry, deletion during peer editing, and complete saved digest convergence. Three focused browser cases pass in `/tmp/b17-browser2.log`. The affected host/provider/measurement/lifetime suite passes fourteen cases and 427 assertions in `/tmp/b17-affected.log`. Six original measurement cases pass separately in `/tmp/b17-measurement.log`. Private external-core build passes with 210 modules. Browser types pass 1764 files. The first browser attempt exposed an incorrect test traversal from the document root; traversing the saved Word body fixes that test helper. No acceptance oracle or deadline changed. The completed command `SCIOBOT_BROWSER_STABLE=1 bunx playwright test --config e2e/sciobot-block-host.config.ts sciobot-block-overflow.spec.ts sciobot-block-measurement.spec.ts sciobot-block-fragments.spec.ts` passes all eight cases at their original limits: `/tmp/b17-browser-history.log`. The fixture is `packages/sciobot/test/browser/measurement.tsx`. Scrolling alone preserves the complete saved digests. The first combined regression records seven passes and a rapid-action history failure: `/tmp/b17-browser-final.log`. Two successive payload actions previously merged inside the native capture interval. `writeBlockPayload` now uses the existing separate-history flag for each submitted widget action. The no-delay two-session regression and original measurement cases pass ten cases and 216 assertions: `/tmp/b17-history.log`. The affected payload concurrency, field, and label collaboration suites pass 38 cases and 1092 assertions: `/tmp/b17-payload-affected.log`. Native fingerprints, exact payloads, peer edits, and reopened digests remain checked. Final private build, 1764-file browser typecheck, root format, scoped format, root lint, diff check, and compatibility policy pass. Logs: `/tmp/b17-build-final.log`, `/tmp/b17-types-history.log`, `/tmp/b17-root-format.log`, `/tmp/b17-lint-final.log`, `/tmp/b17-policy-final.log`. Decision: `.collaboration/changes/sciobot-block-overflow.json`; local viewport state changes no shared format fields. The production listening renderer remains its separate domain leaf.

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

- [x] Reserve inline widget geometry in layout

**Depends:** [B18](02-react-blocks.md#b18), [B02](02-react-blocks.md#b02), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/layout/inline-widget-layout.ts`; `focused layout and paint tests`

**Read:** [S25](sources.md#s25)

**Change:** Consume inline metrics through the permitted layout boundary. Reserve boxes without changing the canonical fallback text.

**Close when:** A chip wraps at its measured width. Adjacent native text and caret geometry match explicit box coordinates.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `edc0047d`, completed local source. Optional `TextMeasurer.inlineWidgets` consumes a renderer-neutral provider and revision snapshot. CSS pixel measurements convert to points. Complete visible text controls become one unbreakable span, with native fallback text, source UTF-16 ranges, control identity, and revision metadata retained. Invalid or declined metrics preserve native layout. Width clamps to the current line after unit conversion, avoiding fixed-point rounding overflow. Bidi itemization preserves admitted control edges; restoring the atom retains resolved native direction. Internal and trailing chip whitespace cannot split, shrink, or receive justification slack. Mixed native revision attribution and field atoms retain native spans instead of collapsing into one widget. Body, table, and furniture producer keys include provider identity and metric revision. Read-only layout never changes canonical content or writes shared fields. Paint reserves the published advance and exposes the bound control identity. Core exports the related types from layout and editor entry points. Host renderer mounting and measurement publication remain B20.

Focused layout checks pass eight cases and 45 assertions. They prove exact 30-point and 72-point advances, two-line wrapping, 80-point clamping, 15-point ascent and 3-point descent, adjacent caret positions, hit testing, native fallback, table cells, RTL/justified whitespace, source immutability, paint width, and warm-cache invalidation when a provider or revision changes. Four real shared scenarios pass 204 assertions, covering both client orders and both reconnect receivers, concurrent text around the control, peer-preserving history, deletion during peer editing, remote caret geometry, metric revision changes, saved convergence, and measured save/reopen geometry. Their metrics enter the pure layout boundary over each real session's native package; they do not claim B20 host integration. The affected command runs `inline-widget-layout`, equation, RTL, tabs, paragraph cache, semantic interaction, boundary, CJK, paint, lane-import, and private collaboration regressions: 252 pass, zero fail, 3233 assertions across 13 files (`/tmp/solo-b19-final-affected.log`). Core types pass 1548 files; private types pass 1990 files. The complete core build and scoped API extraction/check pass. Only additive layout/editor API snapshots change. The final review hardening rerun passes 12 focused/shared cases and 249 assertions, including exact native spans in All Markup, Original, and Proposed views (`/tmp/solo-b19-review-focused.log`). Core build and API check pass again after that guard. Adapter parity, single-core dependencies, and browser bundle boundaries pass 24 tests and 50 assertions across five files. The private build verifies 202 modules with external core. Root lint and scoped formatting pass; line caps pass after extracting equation span assembly and declaring seven reviewed paint composition lines. Backlog validation passes 270 leaves with 164 closed; all 198 production source hashes match. Compatibility policy passes 131 decisions with zero format fields changed. Minor changeset `.changeset/inline-widget-layout-boxes.md`; compatibility decision `.collaboration/changes/sciobot-inline-widget-layout.json`. The test transport uses the EigenPal Pro License; the production path remains Apache core and private adapters.

## B20

- [x] Mount React inline widgets in native boxes

**Depends:** [B30](02-react-blocks.md#b30), [B04](02-react-blocks.md#b04), [B19](02-react-blocks.md#b19), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/react-blocks/inline-host.tsx`; `sibling inline-host browser test`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Mount registered inline views into measured native boxes. Reuse provider and disposal handling from the block host.

**Close when:** A source chip renders between native words. Zoom, repaint, and provider updates retain identity and correct hit geometry.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`; implementation starts from `01081732`. `mountReactInlineHost` reuses block roots, explicit providers, input isolation, parking, and deferred disposal. One root represents each persisted inline payload. Intrinsic width and baseline measurements publish through B30 without document writes. Native layout supplies root and renderer geometry. Missing or declined measurements retain visible Word fallback. Inline hosts preserve separate block-height reservations.

`bun test --timeout 30000 --preload ./scripts/test/preload-dom.ts packages/sciobot/src/react-blocks` passes 26 cases and 1062 assertions. These cover geometry, registration, cleanup, provider identity, zoom, fallback restoration, and independent block reservations. Existing measurement tests emit React act warnings; the new inline cases emit none. Log: `/tmp/solo-b20-react-blocks-final.log`.

`SCIOBOT_BROWSER_STABLE=1 bunx playwright test --config e2e/sciobot-inline-host.config.ts` passes five Chromium cases. Four shared cases cover both client orders and both reconnect receivers. They verify different local widths, concurrent payload/native edits, peer-preserving history, exact remote-caret rectangles, deletion, and saved convergence. The source-chip case verifies focused local input, React state, native typing around the chip, and provider/zoom geometry. Saved DOCX reopens with equal complete semantics and measured widgets. Additional mouse assertions pass in the focused source-chip rerun. Native clicks select exact chip boundaries before and after zoom. Clicks remain 1.5 pixels outside widget controls because Chromium rounds subpixel hit points. Logs: `/tmp/solo-b20-chromium-final.log`, `/tmp/solo-b20-mouse-final.log`. Fixtures: `inlineWidgetDocument()`, `test/browser/inline-host.tsx`.

Private types pass 1996 files; browser types pass 1720 files. The private build verifies 205 modules with external core and no production Pro imports. Root formatting, root lint, line caps, scoped formatting, and diff checks pass. Compatibility policy passes 133 decisions with zero format fields changed. Decision: `.collaboration/changes/sciobot-inline-react-host.json`. All 198 production source hashes and feature mappings match application revision `0ea531f1e52e84db4cb613b9bc2ca61e1d3ea0e9`. Coverage records the inline foundation; B21 retains whole-widget navigation, deletion, and clipboard behavior. The test transport uses the EigenPal Pro License.

## B21

- [x] Edit inline widgets as native atoms

**Depends:** [B20](02-react-blocks.md#b20), [B31](02-react-blocks.md#b31), [B08](02-react-blocks.md#b08), [B10](02-react-blocks.md#b10), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/react-blocks/inline-input.ts`; `fork:e2e/sciobot-inline-input.spec.ts`

**Read:** [S10](sources.md#s10), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Connect arrow navigation, selection, deletion, copy, and paste to the inline content-control boundary.

**Close when:** Backspace deletes one chip. Undo restores its payload. Copy/paste retains attributes with a fresh ID and no split atom.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `123d52ff`, completed local source. `mountInlineInput` connects canonical inline selection, native atom navigation/deletion, and payload clipboard handling. Partial forward and reverse selections expand before replacement, cut, or copy. Collapsed carets retain native buffered typing. Shared clipboard leases install one handler, preserve independent callbacks, and reject conflicting ID factories. Listeners follow the editor's current container and document. Widget form input and composition events retain their existing boundaries.

The new Chromium command `SCIOBOT_BROWSER_STABLE=1 bunx playwright test --config e2e/sciobot-inline-input.config.ts` passes five cases. Real arrows cross one complete chip. Real Shift+ArrowRight selects its complete native range. Ctrl+C/Ctrl+V preserves attributes with a fresh ID. Partial replacement and cut remove the whole chip. One Ctrl+Z restores exact package semantics, including payload parts. Widget input cut/paste leaves Word content unchanged. Four real-session cases exercise both client orders and reconnect receivers. Offline paste and native peer typing converge. Paste undo preserves peer text. Deletion during peer typing, undo/redo, exact remote caret geometry, and complete saved digests converge. Log: `/tmp/b21-input-final-verified.log`.

Regression checks find and fix two concrete defects. The test transport adopted an empty retained custom XML record after paste undo. It now keeps that unlisted empty shell in the registry without adding it to the package. Listed nodes and authored orphan payloads remain intact. Per-keystroke normalization also interrupted native buffered typing beside a chip. Collapsed carets now stay in the native lane. The original six Chromium host cases pass unchanged, including providers, zoom, rapid typing, Backspace/Delete, peer payload edits, history, and Word reopen. Log: `/tmp/b21-host-final-verified.log`. The transport remains test-only under the EigenPal Pro License.

`bun test packages/sciobot/src/react-blocks packages/pro/src/collaboration/__tests__/document-custom-nodes-concurrent.test.ts` passes 41 cases and 1293 assertions across 13 files. `bun test packages/sciobot/src/blocks` passes 73 cases and 1465 assertions across nine files. Unit checks include cross-document reattachment, independent clipboard disposal, partial ranges, and buffered-word undo. The tracked shared paste check restores the entire peer-only package after undo. Logs: `/tmp/b21-affected-final-verified.log`, `/tmp/b21-payload-regressions.log`. Private types pass 2035 files; browser types pass 1757. Pro types pass 2273 and 1831 files. Private and Pro builds pass. The private boundary verifies 206 modules, one external core, and no production Pro imports. Pro API checks scan ten entries without errors or signature changes. Root formatting, root lint, scoped formatting, and diff checks pass. Decision: `.collaboration/changes/sciobot-inline-input.json`; no shared format fields change. Patch changeset: `empty-custom-xml-undo-shells`. Native desktop IME remains B23; application date and mention renderers remain X21/X22.

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

## B30

- [x] Publish inline widget measurements as surface view inputs

**Depends:** [B19](02-react-blocks.md#b19), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/editor/surface-inline-widget-metrics.ts`; `fork:packages/core/src/editor/paginated-surface.ts`; `fork:packages/core/src/editor/__tests__/surface-inline-widget-metrics.test.ts`; `fork:packages/sciobot/src/react-blocks/surface-inline-metrics-collaboration.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Publish immutable inline metric snapshots without replacing the editor, changing native content, or entering document history. Preserve measurements across font changes and zoom. Clear host measurements on disposal and keep canonical Word save calculations on native fallback metrics.

**Close when:** A live native editor changes chip width after publication, preserves native text, selection, history, and saved content, retains the measurement after zoom and font-measurer replacement, and restores native fallback when cleared.

**Check:** Add native surface regressions and the required real two-session scenario. Run affected font/zoom/layout tests, types, public build/API checks, changeset, and compatibility policy.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `cdca9134`, completed local source. `setSurfaceInlineWidgetMetrics` publishes or clears an immutable local snapshot without replacing the surface or its session. Measurement wrappers preserve original method receivers, optional shaping methods, and old pass snapshots. Zoom and font replacement reapply the active host snapshot. Identical provider/revision pairs do not repaint. Body and furniture geometry refresh together; header/footer break keys include metric identity/revision. Canonical save calculations unwrap host metrics, retaining Word fallback pagination and cached PAGEREF values. Native IME composition retains its live DOM; publications coalesce and apply after composition readback. Only local measurement state changes.

`bun test --timeout 30000 --preload ./scripts/test/preload-dom.ts packages/core/src/editor/__tests__/surface-inline-widget-metrics.test.ts packages/sciobot/src/react-blocks/surface-inline-metrics-collaboration.test.ts` passes nine cases and 156 assertions (`/tmp/solo-b30-ime-focused.log`). Five native cases prove width updates, selection/content/history retention, zoom, real font-measurer replacement, prototype receivers, immutable prior snapshots, native Word save pagination, header cache invalidation, and IME readback. Four real two-session cases cover both client orders and receivers, distinct local widths, exact remote caret coordinates, concurrent text, peer-preserving undo/redo, deletion during peer editing, reconnect, clearing, and saved convergence. The affected font, surface layout, furniture, reservation, save-field, caret, widget, and lane suite passes 98 cases and 1177 assertions across ten files before the isolated IME guard; the focused suite then passes with that guard. Core types pass 1549 files and private types pass 1992 files. Core build and public API checks pass; the editor snapshot adds only the publication function and its snapshot type. Root lint and line caps pass. Compatibility policy passes 132 decisions with zero format fields changed. Minor changeset `.changeset/surface-inline-widget-metrics.md`; decision `.collaboration/changes/sciobot-surface-inline-metrics.json`. Backlog validation passes 271 leaves with 165 closed; all 198 production hashes and feature mappings match after auditing application revision 7b61c6cb07510d21d14a4ec2eb753958a7870401. The shared test transport uses the EigenPal Pro License. B20 retains React mounting, browser identity, providers, and interactive geometry verification.

## B31

- [x] Delete measured inline widgets as native atoms

**Depends:** [B19](02-react-blocks.md#b19), [B30](02-react-blocks.md#b30), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/editor/surface-range-edit.ts`; `fork:packages/core/src/editor/surface-locked-control-delete.ts`; `sibling native inline-widget input tests`

**Read:** [S25](sources.md#s25)

**Change:** Recognize measured inline control boxes in native adjacent and range deletion plans. Preserve native locks, tracking, and history. Keep measurement-driven atom behavior local to the surface.

**Close when:** Backspace and Delete remove one complete measured inline control. Selected ranges remove complete controls with surrounding text. Undo and redo restore exact native content and tags. Cleared measurements restore ordinary native control behavior. Locked and tracked controls retain their native edit rules.

**Check:** Add focused native tests and the required two-session history, geometry, deletion, reconnect, and save/reopen scenario.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`; implementation follows `640cbbf3`. The reproduction changed `Source` to `Sourc` after one native Backspace. Native input now classifies admitted measured controls as atoms. Adjacent deletion removes an unlocked control as one node. Selected deletion removes complete controls plus surrounding text. Native pointer/keyboard selection expansion includes measured controls without changing their canonical lock or type. Clearing measurements restores ordinary native character deletion.

The focused native/shared command passes 14 cases and 120 assertions (`/tmp/solo-b31-focused-final.log`). Native cases verify Backspace, Delete, exact history, Word roundtrip, selected ranges, selection expansion, cleared metrics, and three lock modes. Suggestion mode retains native unlocked SDT wrappers and tracks the complete fallback inside them. Native locks still determine selection-before-removal and removal refusal. Four real shared scenarios cover both client orders and reconnect receivers. Deletion preserves concurrent peer payload updates and surrounding text. Undo restores the control with the peer's latest payload. Local widths retain exact independent remote-caret coordinates. Redo and save/reopen converge.

The affected native input, pointer, range, measurement, isolation, and private React-block command passes 204 cases and 1629 assertions. Log: `/tmp/solo-b31-affected.log`. All six Chromium cases pass in `e2e/sciobot-inline-host.config.ts`. Real keyboard Backspace and Delete each remove one chip. One real Ctrl+Z restores its fallback and exact payload in both sessions. Existing mouse, zoom, provider, native input, concurrent history, remote geometry, deletion, and saved-convergence cases remain green. Log: `/tmp/solo-b31-chromium-final.log`. The first browser run overlapped CSS regeneration and failed to load its stylesheet. The verified rerun follows the completed core build.

Core types pass 1550 files; private types pass 1999 files; browser types pass 1721 files. Core and private builds pass. The private boundary verifies 205 modules and external core. Core API checks scan 15 entries with no signature changes or errors. Root lint, line caps, scoped formatting, and diff checks pass. Compatibility policy passes 135 decisions with zero format fields changed. Patch changeset: `measured-inline-native-deletion`. Decision: `.collaboration/changes/sciobot-inline-native-deletion.json`. Test transport uses the EigenPal Pro License. B21 retains private input and clipboard integration.

## B32

- [x] Move native elements without replacing identity

**Depends:** [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/package/ooxml-edit.ts`; `fork:packages/core/src/store/__tests__/move-node.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Expose an atomic canonical element move over the existing shared move primitive. Preserve subtree identities and inherited namespace meaning.

**Close when:** Native moves preserve descendant IDs and text, support same-parent ordering, reject cycles without writes, and journal a move rather than subtree replacement. QName-valued opaque metadata survives conflicting namespace scopes and save/reopen.

**Check:** Add native regressions and real-session movement, peer text, undo/redo, reconnect, and save/reopen checks. Run native builds, public API checks, types, and compatibility policy.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `ddb95987`, completed local source. The store exports `moveNode(part, nodeId, destinationParentId, destinationIndex, options)`. Its index refers to destination children after source removal. The edit rewrites only source and destination ancestor paths. Descendant objects and IDs stay intact. Native level checks reject incompatible block/inline placement. Cycles, roots, missing nodes, and invalid indices refuse before journal effects. Same-position moves retain the original part.

A normal move records exactly one existing `moveNode` effect. Inherited namespace differences add existing `setNamespaceBinding` effects. The edit preserves complete source scope, including bindings used only in opaque QName-valued metadata. Four native cases prove literal expected Word structure, immutable input, identical moved objects, exact journal effects, ordering, atomic refusal, and conflicting namespace roundtrip. The shared B13 cases prove real-session text/payload preservation, peer-only undo, redo, reconnect, deletion convergence, remote caret geometry, and saved convergence. A focused shared namespace case verifies the resolved URI before and after movement, history, and save/reopen. Import can replace the original prefix with a stable alias; the required namespace meaning remains exact.

`/tmp/b13-focused-final.log` records 17 passing native/private cases and 199 assertions. The affected run and unchanged isolation rerun are recorded under B13. Core and private builds pass. Core types pass 1583 files, private types pass 2039, and browser types pass 1757. Public API checks scan 15 entries without errors. The store snapshot adds only the native move signature. Logs: `/tmp/b32-core-build.log`, `/tmp/b32-core-types.log`, `/tmp/b13-types-final.log`, `/tmp/b13-browser-types.log`, `/tmp/b32-api-check.log`. Root formatting, root lint, and diff checks pass. Minor changeset: `native-identity-preserving-moves`. Decision: `.collaboration/changes/sciobot-native-block-moves.json`. Existing move and namespace journal encodings remain unchanged.

## B33

- [x] Notify paint hosts when a native surface ends

**Depends:** [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/editor/surface-paint-observer.ts`; `fork:packages/core/src/editor/__tests__/surface-paint-observer.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Restore borrowed native nodes and notify paint observers on surface destruction or document replacement. Make observer removal idempotent.

**Close when:** Replacement and repeated destruction notify each active host once. Removed observers receive no destruction callback. One failing host cleanup does not prevent other resources from releasing.

**Check:** Add native lifecycle regressions and B16 real-session/browser lifetime checks. Run affected types, native build/API checks, formatting, lint, and compatibility policy.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `c7629552`, completed local source. `SurfacePaintObserver.destroyed` notifies local hosts when their native surface ends. Destruction restores borrowed paint before releasing roots, measurements, and listeners. The dispatcher clears subscriptions before invoking callbacks. Each restoration and destruction callback has an independent error boundary, so one failed cleanup cannot prevent the others. Explicit observer removal restores borrowed content once and receives no later destruction event.

Native lifecycle regressions verify document replacement, repeated editor destruction, idempotent removal, and two throwing cleanup callbacks beside a healthy host. The combined native/retention command passes five cases and 77 assertions. Log: `/tmp/b16-focused2.log`. The affected React host and native window-move command passes 74 cases and 1376 assertions across 15 files. Log: `/tmp/b16-affected.log`. A later focused retention command adds twenty inline replacement cycles and passes both lifecycle cases with 103 assertions. Log: `/tmp/b16-inline-retention.log`. B16 records real-session virtualization, deletion, history, resource counts, and Word convergence.

Core and private builds pass. Private output verifies 209 modules with one external core. Core types pass 1583 files; private types pass 2041; browser types pass 1762. The public API checker scans 15 entries without errors and adds only the optional destruction callback. Root lint and line caps pass. Native surface composition grows by one line; cleanup dispatch stays in the observer module. Logs: `/tmp/b33-types.log`, `/tmp/b33-core-build.log`, `/tmp/b16-build-final.log`, `/tmp/b33-api-check.log`, `/tmp/b16-lint.log`. Minor changeset: `surface-paint-observer-lifetime`. Decision: `.collaboration/changes/sciobot-react-retention.json`. Shared operations, schemas, and payload encodings remain unchanged.
