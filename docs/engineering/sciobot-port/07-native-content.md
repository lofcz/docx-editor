# Native Word editing controls

Reuse existing native commands. Close a leaf with verified production controls when the engine already provides the behavior.

Read the [dispatch and completion rules](../../../SCIOBOT_PORT_TODO.md) before claiming a leaf.

Each named test or artifact is a deliverable. These specifications do not claim those checks already exist.

## N01

- [x] Create a native command capability fixture

**Depends:** None.

**Primary files:** `fork:packages/sciobot/src/__tests__/fixtures/native-commands.json`

**Read:** [S01](sources.md#s01), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** List the production command IDs and expected before/after semantics. Map each to native commands or a named missing implementation leaf.

**Close when:** Every BasicBlocks, BasicMarks, Font, List, Align, LineHeight, Link, and Table action has an explicit mapping.

**Check:** Validate fixture schema, source coverage, and independently captured expected results.

**Claim:** astra N01, branch `codex/react-block-proof`. **Evidence:** Fork baseline `3278fa04`; application revision `21b9d47b79d4e6c29d6b07d1952e6a269b72c13e`. `scripts/sciobot-port/capture-native-commands-reference.mjs` executes installed Plate transforms with the production kit configurations. It extracts the actual production block conversion and insertion functions; unrelated container dispatch entries are excluded. Render imports are inert, while editing transforms and configured input rules execute unchanged. `native-command-runner.mjs` captures 158 before/after document, selection, and mark states. Cases cover all eight required families, list styles/check state, mark toggles at ranges and carets, font sizes/colors/clearing, every configured alignment/line-height value, links, table insertion/deletion/rows/columns/merge/split/appearance, and typed input rules. One identity map spans each before/after pair, preserving replacement versus retained identity. Explicit command IDs and resolved inputs accompany the results. 57 source/package hashes, capture/runner hashes, UI consumers, input declarations, and resolved keyboard shortcuts record provenance. Mappings identify N02–N14, N22 for table margin, and native `deleteTable`; combined row/column growth names both N11 and N12. Logical alignment and internal-link records state their unresolved adapter requirements. `bun scripts/sciobot-port/capture-native-commands-reference.mjs --check` independently replays all 158 cases and matches the committed artifact. `bun test packages/sciobot/src/__tests__/native-command-fixture.test.ts` passes three schema/coverage tests with 12250 assertions. Recursive text/element and selection checks reject malformed states; required action IDs and keyboard presence are explicit. Scoped Prettier and oxlint pass. This fixture records actual Plate behavior, including rule precedence and default no-ops. It does not establish native command parity, browser shortcut delivery, or UI/network behavior. Those remain the mapped implementation leaves. No production or collaboration code changes.

## N02

- [x] Connect paragraph and heading controls

**Depends:** [N23](07-native-content.md#n23), [C03](01-contracts.md#c03), [N01](07-native-content.md#n01), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/chrome/block-type.ts`; `sibling block-type browser test`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Connect paragraph, H1, H2, H3, blockquote, and horizontal-rule actions to native commands.

**Close when:** Each N01 block fixture matches its expected structure. Toolbar use retains the native caret and one-step undo.

**Check:** Include Enter and paragraph-style controls in a recovered source whose removed mark had a different heading style. Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `1597e5cd`, completed local source. The private adapter exports paragraph/H1-H3/quote conversion, native block insertion, a DOM toolbar, disposable input rules, and pre-room document preparation. Conversion removes list membership and merges only the requested paragraph style, preserving native text, caret ranges, and imported properties. Missing built-in definitions come from the core blank document; existing definitions remain unchanged, and conflicting identities refuse immutable preparation. Infrastructure preparation preserves original story and opaque bytes and is byte-idempotent. Insertions use native paragraph allocation followed by style assignment within the same transaction. Heading/quote prefixes, dash/underscore rules, heading shortcuts, and empty-heading Enter follow the captured N01 actions. Native quote paragraphs represent Plate quote containers; the toolbar retains selection despite the captured Plate quote setter clearing it, as this leaf requires. Rule paragraphs remain empty rather than retaining Plate's hidden void-child dashes. Rule conversion retains the original editable paragraph as the follower and allocates a preceding rule, so offline peer insertions retain their native text owner. No text relocation protocol is added.

`bun test --timeout 30000 --preload ./scripts/test/preload-dom.ts packages/sciobot/src/chrome` passes 49 tests and 639 assertions (`/tmp/solo-n02-all-final.log`). It covers ten plain/list conversions, six insertion kinds, seven heading/quote/Enter inputs, two rule inputs, five preparation cases, eighteen real shared scenarios, and a recovered-source heading regression. Shared cases cover both client orders, offline text during controls and prefix conversion, reconnect, peer-preserving undo/redo, deletion during peer editing, remote selection geometry, and saved convergence. A recovered source from a rejected H1 insertion accepts H2 and splits into two H2 paragraphs without restoring the removed H1 mark. Saved rule checks compare complete native fingerprints rather than assuming transient store IDs survive export. `bunx playwright test --config e2e/sciobot-block-controls.config.ts` passes two Chromium cases (`/tmp/solo-n02-browser-final.log`), proving pointer toolbar selection, one-step undo, continued typing, real heading input, Enter, heading shortcut toggles, rule input, native caret placement, and shared saved convergence. Private types pass 1986 files; the build verifies 202 private modules with external core. Scoped lint and formatting pass. Browser types pass 1710 files. Backlog validation passes 270 leaves with 163 closed and verifies all 198 production source hashes. Compatibility policy passes 130 decisions with zero format fields changed. The coverage audit confirms application revision 83e60516bfd3b506dd56604d96b8656c12c615f6; all source hashes and mappings remain unchanged. Compatibility decision: `.collaboration/changes/sciobot-native-block-controls.json`. The test transport uses the EigenPal Pro License; production controls use Apache core operations and private adapters. Manual suggestion-mode routing remains R23.

## N03

- [x] Connect all basic text marks

**Depends:** [C03](01-contracts.md#c03), [N01](07-native-content.md#n01), [V01](11-verification.md#v01), [N18](07-native-content.md#n18)

**Primary files:** `fork:packages/sciobot/src/chrome/marks.ts`; `sibling marks browser test`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Connect bold, italic, underline, strike, subscript, superscript, and inline code. Keep mixed-selection state accurate.

**Close when:** The seven mark fixtures match expected properties. Keyboard shortcuts and toolbar actions preserve selection.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `adc99b2c`, completed solo source on the same branch. `marks.ts` connects all seven captured marks. Native commands handle six marks; code uses the explicit Word character-style identity from N18. `marks-toolbar.ts` preserves native selection and reports on/off/mixed states. Application shortcuts retain the production key bindings. Native caret character styles lower through the existing `characterStyle` operation field, retain imported character styles, and participate in local undo/redo. Character styles remain outside replaceable run and paragraph-mark property bags. Rectangular cell selections affect only selected cells.

`bun test --timeout 30000 --preload ./scripts/test/preload-dom.ts packages/sciobot/src/chrome/marks.test.ts packages/sciobot/src/chrome/marks-collaboration.test.ts` passes 25 cases and 566 assertions (`/tmp/sciobot-n03-controls-final.log`). Seven captured on/off/caret fixtures, mixed selection, reversed and multi-paragraph selection, imported-style restoration, toolbar state, disposal, and rectangular cells pass. Fourteen real-session cases cover both client orders, concurrent text, independent undo/redo, deletion during peer editing, reconnect, remote selection geometry, canonical convergence, and saved reopening. The affected 142-case integration batch found one imported-style guard defect; the corrected 25-case control batch passes without changing its expected output. All other 131 integration cases passed. The final native caret/history batch passes 29 cases and 131 assertions, including the added Enter regression (`/tmp/sciobot-n03-core-final.log`).

`bunx playwright test --config e2e/sciobot-marks.config.ts` passes seven Chromium cases (`/tmp/sciobot-n03-browser-verified.log`). Each checks a real toolbar click, retained native selection, one-step undo, the application shortcut, mixed state, caret typing, and both sessions saved native XML. Private typecheck passes 2306 files; private build passes 294 modules with external core. Core builds and scoped lint/formatting pass. Compatibility decision: `.collaboration/changes/sciobot-basic-marks.json`; published core patch: `.changeset/caret-character-style.md`. Transport remains test-only under the EigenPal Pro License. Manual suggestion-mode routing remains R23.

The backlog validator runs but reports external application-source drift. Captured application revision `5650bdf550b7987174d1b61fe2a925b86fb24f15` differs from current `416fbad785716b5b16cfcc522ea1f06765227340`. Only `WorkspaceExportModal.tsx` and `workspaceExportArchive.ts` differ among its 198 captured source hashes. Basic mark sources remain unchanged. No source reference, fixture, API snapshot, or coverage baseline was recaptured. This is the validator's only reported error; application export integration remains open in its existing leaves.

## N04

- [x] Connect font family and size controls

**Depends:** [C03](01-contracts.md#c03), [N01](07-native-content.md#n01), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/chrome/fonts.ts`; `sibling fonts browser test`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Connect font family and size with packaged font loading and mixed-selection values.

**Close when:** Selected text changes family and size. Save/reopen preserves both. Native text paints with loaded fonts.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** `chrome/fonts.ts` connects packaged font families, exact CSS-pixel sizes, mixed values, and native clears. `fonts.test.ts` verifies all fourteen captured sizes, full history, Word reopen, caret typing, and character-style inheritance. `font-link-collaboration.test.ts` verifies concurrent edits, font resets, independent history, reconnect, deletion recovery, native selection geometry, and canonical saved convergence in both client orders. Chromium verifies actual packaged font requests, native paint, real typing, mixed values, and clears. The completed batch passes 296 native cases and three Chromium cases. Private and browser type checks pass.

## N05

- [x] Connect text color and highlight controls

**Depends:** [C03](01-contracts.md#c03), [N01](07-native-content.md#n01), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/chrome/colors.ts`; `sibling colors browser test`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Connect foreground and background colors, including clearing a color and retaining theme references.

**Close when:** Apply and clear each color independently. Unchanged imported theme colors retain their canonical values.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `260fd5f7`, completed solo source on `codex/react-block-proof`. `colors.ts` connects literal RGB foreground and arbitrary RGB backgrounds, clear/inherit/transparent values, mixed values, native caret preferences, and a host-supplied palette. Backgrounds use native character shading, rather than approximating the application palette with Word's limited fixed highlights. An explicit background overrides a fixed highlight; clearing removes direct shading/highlight and restores inheritance. Foreground and background actions remain independent. Native color removal emits no empty color elements. Property rewrites retain foreign attributes, bindings, children, and unchanged theme references. The new `shd` capability receives one explicit manifest entry and its matching API tuple item; no capture or baseline-generation command ran.

`bun test --timeout 30000 --preload ./scripts/test/preload-dom.ts packages/sciobot/src/chrome/colors.test.ts packages/sciobot/src/chrome/colors-collaboration.test.ts packages/core/src/store/__tests__/run-background.test.ts packages/core/src/store/__tests__/tree-ops.test.ts packages/core/src/store/__tests__/authorable-mutation-manifest.test.ts packages/core/src/editor/__tests__/pending-caret-formatting.test.ts packages/core/src/editor/__tests__/pending-character-style.test.ts` passes 137 cases and 876 assertions (`/tmp/sciobot-n05-batch.log`). The final combined batch includes those files, native formatting-review cases, the oracle regression, and inherited-highlight Enter coverage; all 143 cases and 965 assertions pass across nine files (`/tmp/sciobot-n05-final-tests.log`). Explicit picker actions override a fixed highlight; carrying imported shading through Enter does not override inherited highlight styling. It covers apply/clear, arbitrary RGB values, imported themes, foreign metadata, caret typing/history, fixed-highlight override, malformed input refusal, native property vocabulary, and eight real-session cases. Those sessions cover both client orders, concurrent peer text, independent undo/redo, deletion during peer edits, reconnect, remote geometry, and canonical save/reopen convergence.

Four additional real-session native formatting-review cases and one oracle regression pass five cases and 83 assertions (`/tmp/sciobot-n05-review-final.log`). Accept/reject preserve collaborator text and complete independent native formatting semantics. Rejection can retain equivalent text-run splits, as existing Word formatting does. The test-only oracle permits only adjacent text-only runs with identical complete attributes and property subtrees to share a text token. It retains theme references, unknown properties, wrapper boundaries, and all other structure. Separate raw canonical fingerprints still match each editor exactly after save/reopen; no production normalization or peer-identity coalescing was added. The oracle regression rejects changed theme references, colors, language, run attributes, wrappers, breaks, and text. Manual suggestion-mode application routing remains R23.

`bunx playwright test --config e2e/sciobot-colors.config.ts` passes four Chromium cases (`/tmp/sciobot-n05-browser-final.log`). Pointer apply/clear preserves native selection and one-step undo; actual foreground/background paint is RGB(23,162,184). Independent colors, imported neighbor properties, real typing, and both sessions' saved native fingerprints pass. Core builds. Private typecheck passes 2317 files; private build passes 297 modules with external core. Scoped lint/formatting and file-size guards pass. Compatibility decision: `.collaboration/changes/sciobot-native-colors.json`; published minor change: `.changeset/native-run-backgrounds.md`. Transport remains test-only under the EigenPal Pro License. N03 records the unrelated application export-source drift that still fails the backlog audit.

## N06

- [x] Connect alignment and indentation controls

**Depends:** [C03](01-contracts.md#c03), [N01](07-native-content.md#n01), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/chrome/paragraph-position.ts`; `sibling paragraph-position browser test`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Connect alignment and paragraph indentation without changing list membership or neighboring paragraph properties.

**Close when:** N01 alignment and indentation fixtures pass. One undo restores the exact original paragraph properties.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `cdbcd919`, completed solo source on `codex/react-block-proof`. `paragraph-position.ts` connects all six captured alignments and both indentation actions. Physical left/right resolve independently for each paragraph direction; logical start/end retain native Word values. The application IndentKit step is 24 CSS pixels, mapped to 360 twips. Relative indentation uses each paragraph's native resolved indent. Canonical property plans preserve list membership, right and hanging indents, independent spacing, unknown properties, neighboring paragraphs, and selection. A toolbar action creates one native undo unit.

`bun test --timeout 30000 --preload ./scripts/test/preload-dom.ts packages/sciobot/src/chrome/paragraph-position.test.ts packages/sciobot/src/chrome/paragraph-position-collaboration.test.ts` passes 32 cases and 604 assertions (`/tmp/sciobot-n06-final.log`). Local cases cover all six alignments in both directions, plain/list indentation, mixed paragraph indents, canonical save/reopen, and exact original undo. Sixteen real-session cases cover both client orders, concurrent peer text, independent undo/redo, deletion during peer editing, reconnect, remote caret/selection geometry, imported exact-spacing retention, and saved canonical convergence.

`bunx playwright test --config e2e/sciobot-paragraph-position.config.ts` verifies all eight controls through pointer actions, native range selection, one-step undo, real typing, and both sessions' saved/reopened fingerprints. Seven cases pass in `/tmp/sciobot-n06-browser.log`; a source reload interrupted right alignment. Its unchanged assertions pass in the isolated rerun `--grep 'right preserves'` (`/tmp/sciobot-n06-browser-right.log`). Private typecheck passes 2312 files. Private build passes 296 modules with external core. Scoped lint and formatting pass. Decision: `.collaboration/changes/sciobot-paragraph-position.json`; existing native operations need no published changeset. The evaluated transport uses the EigenPal Pro License. N03 records the external application export-source drift that still fails the backlog audit; no reference baseline was recaptured. Manual suggestion-mode routing remains R23.

## N07

- [x] Connect line height controls

**Depends:** [C03](01-contracts.md#c03), [N01](07-native-content.md#n01), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/chrome/line-height.ts`; `sibling line-height browser test`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Map production line-height choices to native paragraph spacing rules and retain imported exact spacing.

**Close when:** Each existing line-height choice survives reopen. Editing a nearby paragraph does not normalize an imported exact line height.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `adc99b2c`, completed solo source on `codex/react-block-proof`. `line-height.ts` connects production choices 1, 1.2, 1.5, 2, and 3 through native `setLineSpacing`. The toolbar reads native multiple, exact, minimum, and mixed states without writing. Pointer actions retain selection. Native paragraph spacing merges preserve before/after settings, indentation, unknown properties, and nearby exact line heights. Plate omits its default 1.5 property; Word receives an explicit 1.5 multiple to preserve the same effective value.

`bun test --timeout 30000 --preload ./scripts/test/preload-dom.ts packages/sciobot/src/chrome/line-height.test.ts packages/sciobot/src/chrome/line-height-collaboration.test.ts` passes 16 cases and 372 assertions (`/tmp/sciobot-n07-final.log`). Five captured choices survive canonical save/reopen with exact original undo. Imported 12.55-point spacing remains unchanged. Ten real-session cases cover both client orders, concurrent text, independent undo/redo, deletion during peer editing, reconnect, remote geometry, saved native fingerprints, and convergence. Native defaults come from the actual room document; undo must restore its exact original properties and effective spacing.

`bunx playwright test --config e2e/sciobot-line-height.config.ts` passes five Chromium cases (`/tmp/sciobot-n07-browser-final.log`). Each uses a pointer toolbar action, native range selection, real typing, one-step undo, both sessions, and canonical saved/reopened output. Private typecheck passes 2309 files; build passes 295 modules with external native core. Scoped lint and formatting pass. Compatibility decision: `.collaboration/changes/sciobot-line-height.json`; existing native commands require no publishable change. The test transport uses the EigenPal Pro License. The external application source drift recorded under N03 also affects the backlog audit; no reference baseline was recaptured. Manual suggestion-mode routing remains R23.

## N08

- [ ] Connect list and checklist controls

**Depends:** [C03](01-contracts.md#c03), [N01](07-native-content.md#n01), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/chrome/lists.ts`; `sibling lists browser test`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Connect bullet, numbered, checklist, indent, and outdent actions through native list operations.

**Close when:** Nested numbering and check state survive save/reopen. Enter and Backspace match the captured list editing behavior.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Checkpoint from `be36f466`, branch `codex/react-block-proof`. N08 remains open. `chrome/lists.ts` connects all eight captured marker styles, explicit membership, style toggling, checkbox state, and list levels. New list styles author nine native level overrides on a fresh counter without changing imported definitions. Conflicting explicit styles on one new instance refuse before writes. `chrome/list-input.ts` adds the six captured prefixes and direct checklist continuation, empty-item outdent, and Backspace. Direct checklist Enter uses one native transaction and restores exact history.

Two native defects surfaced during this work. Guarded writes now accept an optional post-edit selection before native history/presence capture. The application session wrapper now shares the provider's existing undo authority. This restores native caret and caret-format history across shared redo, without a separate history stack or new shared data. The native surface stays at 6354 lines under its existing 6362-line cap. The API snapshot adds only the optional guarded-write callback. A local wrapper symbol forwards the existing native authority without importing core into the DOM-free protocol.

The completed focused integration command passes 82 cases and 1534 assertions across eleven files (`/tmp/sciobot-n08-checkpoint-final-native.log`). It covers controls, prefixes, exact local history, both client orders, concurrent peer text, independent history, deletion recovery, reconnect, remote geometry, saved convergence, wrapped authority, shared redo caret, manual AI isolation, font history, and mark history. The prior 98-case batch passed 96 cases; two new core fixtures lacked document bytes. The corrected final core/control batch passes 45 cases and 575 assertions (`/tmp/sciobot-n08-verified-native.log`). Existing tracked checklist and native selection-history regressions pass in that broader batch. All original deadlines remain unchanged. Native numbering allocation persistence remains the documented session policy; redo restores the exact complete numbered package.

`bunx playwright test --config e2e/sciobot-lists.config.ts` passes three Chromium cases in 41.0 seconds (`/tmp/sciobot-n08-browser-final.log`). Actual style controls, imported Roman numbering, typing, prefix conversion, nesting, unchecked continuation, empty-item outdent, Backspace, exact undo/redo selection, and equal saved packages pass. Core/private/browser type checks pass 1613/2358/1971 files. The private build verifies 310 modules with core external. Compatibility decision: `.collaboration/changes/sciobot-manual-list-controls.json`; core changeset: `.changeset/native-guarded-selection.md`. Transport remains test-only under the EigenPal Pro License.

The next checkpoint adds atomic tracked checklist continuation. Native split, selection deletion, checkbox insertion, and review ownership commit together. Detached native preflight measures shared history ownership before admission. One undo restores the original document and caret. One review group contains the proposed break and checkbox. Splits through pending AI text preserve separate AI review ownership. Native rejection keeps the second paragraph identity, as Word paragraph-mark removal requires. The complete native oracle preserves every other attribute, property, control, and content boundary. It permits only equivalent adjacent text-run splits.

Mixed plain/list indentation now uses one canonical plan and one history unit. Plain paragraphs retain their effective indentation and independent first-line settings. Lists retain their counters, levels, and checkbox state. Suggesting uses the existing native review decisions for each paragraph. All eight styles reuse compatible preceding or following counters without editing imported definitions. Plain gaps still start independent sequences.

The focused checks cover 91 scenarios across ten files. Eighty-eight pass in `/tmp/n08-continuation-final-native.log`. The corrected four-case mixed-review fixture passes in `/tmp/n08-continuation-mixed-final.log`, including the two remaining scenarios. The fixture retains an authored extension property during rejection. Twelve shared continuation cases cover both client orders, end/middle/expanded selections, offline peer text, accept/reject, independent histories, caret redo, and saved/reopened convergence. Four shared mixed-indentation cases pass. Sixteen style-continuity cases pass. Two pending-AI ownership cases pass. Eleven continuation cases pass in `/tmp/n08-continuation-refusal-final.log`, including refusal without partial writes or fallback. No original deadline changed.

Five Chromium cases pass in 40.4 seconds (`/tmp/n08-continuation-final-browser.log`). They include real controls for adjacent counters, mixed indentation, and tracked checklist Enter. Exact native history, caret redo, rejection, and both saved packages pass. Private and browser type checks pass 2364 and 1972 files. The private build verifies 311 modules with native core external. No published API or shared format changes in this checkpoint.

Remaining before closure: implement the captured `list.insert.disc`, `list.insert.decimal`, and `list.insert.todo` selection/insertion effects. Finish sequence bridges and sequence-boundary parity. Run complete N01 list capture parity and the corresponding browser/shared scenarios. N08 remains open. No unsupported branch establishes completion.

## N09

- [x] Connect manual table insertion

**Depends:** [C03](01-contracts.md#c03), [N01](07-native-content.md#n01), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/chrome/tables.ts`; `sibling tables browser test`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Connect table grid selection and insertion to native commands. Keep the caret in the first inserted cell.

**Close when:** Insert a 2x3 table at the caret. Native cell navigation works. One undo removes exactly the inserted table.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `51c2d4ce`, completed source in the current native-table batch. `tables.ts` connects the production eight-by-eight insertion picker and native model selection. Inserting a 2x3 table retains source text and adopts the first native cell. Tab/Shift+Tab navigate cells; final-cell Tab inserts one native row. The engine now rejects stale provisional caret hints and resolves the inserted table from the operation's exact canonical following-paragraph address. Explicit native history groups stop at their own step, so a retained row birth cannot skip into earlier cell typing. No document or journal schema changes.

The focused native/history batch passes 18 cases and 218 assertions across four files (`/tmp/sciobot-n09-history-tests.log`). Cases cover six real-session insertion/history/deletion/reconnect scenarios across both client orders, remote geometry, peer-preserving undo/redo, exact canonical save/reopen, and retained-row-birth undo after typing. Existing history-group capture and Editor.exec gesture tests pass. `bunx playwright test --config e2e/sciobot-tables.config.ts` passes both Chromium cases (`/tmp/sciobot-n09-browser-verified.log`), including pointer grid selection, actual cell typing, Tab/Shift+Tab, final-cell row growth, and exact undo/redo/save oracles. The shared integration batch passes 115 cases and 1396 assertions across eleven files (`/tmp/sciobot-table-batch-tests.log`), covering table plans, explicit history groups, typing history, and owned AI presence. All six combined Chromium cases pass (`/tmp/sciobot-table-batch-browser-final.log`). Browser fixture types pass 1897 files. Core and Pro builds pass; the private SDK build verifies 299 modules with core external. Decisions: `.collaboration/changes/sciobot-table-insertion-controls.json`, `.collaboration/changes/native-gesture-undo-boundaries.json`. Changesets: `.changeset/shared-table-insertion-caret.md`, `.changeset/native-gesture-undo-boundaries.md`. The transport change remains evaluation-only under the EigenPal Pro License. N11 completes manual row toolbar and recovered-source coverage; The backlog audit remains blocked by external application-source drift. At verification, application HEAD is `3e11da39332e846269141bc779b6e521fdb17ef3`; the captured revision remains `5650bdf550b7987174d1b61fe2a925b86fb24f15`. Changed captured files are `WorkspaceExportModal.tsx`, `workspace-lesson-serialize.tsx`, and `workspaceExportArchive.ts`. No capture command ran.

## N10

- [x] Connect hyperlink editing and activation

**Depends:** [C03](01-contracts.md#c03), [N01](07-native-content.md#n01), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/chrome/links.ts`; `sibling links browser test`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Connect link insertion, editing, removal, and activation. Keep document selection when the link dialog opens.

**Close when:** Edit and undo a link target without changing its label. Unsafe targets fail through the existing URL policy.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** `chrome/links.ts` connects native insertion, retargeting, removal, safe activation, selection pins, Ctrl+K, and native link-click chrome. `links.test.ts` verifies unchanged labels, exact undo/redo, Word reopen, unsafe-target refusal, no-op targets, view-mode refusal, and document replacement guards. `font-link-collaboration.test.ts` verifies concurrent edits, independent history, reconnect, peer-preserving deletion recovery, native geometry, and canonical saved convergence in both client orders. The persistent relationship allocation fix preserves reentrant local content history. The completed batch passes 296 native cases and three Chromium cases. Private and browser type checks pass.

## N11

- [x] Connect manual table row controls

**Depends:** [N09](07-native-content.md#n09), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/chrome/table-rows.ts`; `sibling table-rows browser test`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Connect row insertion and deletion using native table selection. Preserve neighboring row properties.

**Close when:** Insert and delete a middle row. Undo restores its cells, text, and header behavior.

**Check:** Include a caret in recovered source text displayed inside the first table cell. Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `51c2d4ce`, completed with N09 in the shared table-controls batch. `table-rows.ts` connects native insert-above, insert-below, and delete-row commands through selection-preserving controls. Shared row insertion resolves its live first-cell caret from the exact source row and insertion side when the operation hint contains a provisional identity. Neighboring rows retain complete native properties, header behavior, unknown attributes, and cell text.

Eleven focused cases pass with 201 assertions (`/tmp/sciobot-n11-tests.log`). Six real-session cases cover all three actions and both client orders, concurrent peer edits inside the affected row, independent own/peer history, deletion, reconnect, remote caret/selection geometry, and exact canonical save/reopen. Two further real-session cases reject an AI insertion in cell one while a peer edits it. They preserve recovered peer text through row insertion, independent typing, and exact full-document undo/redo. The shared integration batch passes 115 cases and 1396 assertions across eleven files (`/tmp/sciobot-table-batch-tests.log`). All six combined Chromium cases pass (`/tmp/sciobot-table-batch-browser-final.log`), including four row cases with pointer actions, real typing, exact undo/redo, and Word save/reopen. The recovered-source export oracle permits only removal of its one exact internal retired-paragraph marker; every other value remains exact, and save leaves the live tree unchanged. Browser fixture types pass 1897 files. Core and Pro builds pass; the private SDK build verifies 299 modules with core external. Compatibility: `.collaboration/changes/sciobot-table-row-controls.json` and the N09 history decision. Published fixes have patch changesets; transport changes remain evaluation-only under the EigenPal Pro License. N12 and later table leaves retain their own requirements. The backlog audit remains blocked by external application-source drift. At verification, application HEAD is `3e11da39332e846269141bc779b6e521fdb17ef3`; the captured revision remains `5650bdf550b7987174d1b61fe2a925b86fb24f15`. Changed captured files are `WorkspaceExportModal.tsx`, `workspace-lesson-serialize.tsx`, and `workspaceExportArchive.ts`. No capture command ran.

## N12

- [x] Connect manual table column controls

**Depends:** [N09](07-native-content.md#n09), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/chrome/table-columns.ts`; `sibling table-columns browser test`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Connect column insertion and deletion with native grid handling.

**Close when:** Insert and delete a middle column beside merged cells. Undo restores valid widths, spans, and content.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `0c498ce0`, completed solo on the same branch. `table-columns.ts` connects native left/right insertion and deletion through selection-preserving, native-can-gated controls. Imported nonuniform widths, horizontal spans, vertical merge chains, header behavior, foreign metadata, and retained paragraph content survive the operations and exact history. New shared cells adopt live canonical caret identities through retained grid-column addresses, skipping expanded spans and vertical continuation cells. Gridless insertion from later rows resolves the corresponding first-row source cell required by the existing native operation contract.

Six local cases and twenty real-session cases cover both client orders, horizontal/vertical merged neighbors, concurrent edits inside the affected column, deletion during peer editing, independent own/peer history, reconnect, changed-grid remote caret/selection geometry, typing in fresh cells, and exact Word save/reopen. Four gridless cases verify native typing and exact two-step restoration, including the original absence of a grid. Four further cases preserve peer text and the complete original-column geometry when the insertion author undoes. Peer undo then restores exact grid absence; redo preserves text and geometry. The evaluated transport now hides an undone synthesized grid only when no columns need it. Existing immutable birth-coverage frames preserve original columns when retained peer text needs a born cell. No shared fields, codecs, operation shapes, or document tuple changes.

The final integration command `bun test --timeout 30000 --preload ./scripts/test/preload-dom.ts packages/sciobot/src/chrome/table*.test.ts packages/core/src/editor/__tests__/table-command-plan.test.ts packages/pro/src/collaboration/__tests__/document-table-merge-split.test.ts packages/sciobot/src/review/owned-presence-shared.test.ts packages/sciobot/src/review/track-table-columns-collaboration.test.ts packages/sciobot/src/review/recompose-inserted-tables-collaboration.test.ts packages/sciobot/src/diff/table-grid-split-collaboration.test.ts packages/sciobot/src/diff/table-format-collaboration.test.ts` passes 143 cases and 4187 assertions across thirteen files (`/tmp/sciobot-n12-integration.log`). This includes previous manual table controls and AI table decisions, presence, split/merge history, and born-table recomposition. `bunx playwright test --config e2e/sciobot-table-columns.config.ts` passes all eight Chromium cases (`/tmp/sciobot-n12-browser-verified.log`), including actual pointer controls, keyboard input, exact undo/redo, merged widths/spans, gridless insertion, and full canonical save/reopen. Original deadlines remain unchanged. `bun check -p packages/sciobot/tsconfig.browser.json` passes 1901 files. Core and Pro builds pass. The private SDK build verifies 300 modules, external core, and its production dependency boundary. Decision: `.collaboration/changes/sciobot-table-column-controls.json`. Patch changesets: `.changeset/shared-column-insertion-caret.md` and `.changeset/gridless-column-birth-undo.md`. The transport remains evaluation-only under the EigenPal Pro License. N19 retains reference-cell addressing in arbitrary grid states; these controls continue using existing native operation shapes. N13/N14 retain merge/split and appearance controls. The backlog audit still detects external application-source drift; no capture command runs. At verification, application HEAD is `a8ec67cc6e3cf0daf5abafeaffba5e36dfd4e406`; changed captured files remain `WorkspaceExportModal.tsx`, `workspace-lesson-serialize.tsx`, and `workspaceExportArchive.ts`.

## N13

- [x] Connect manual table merge and split controls

**Depends:** [N09](07-native-content.md#n09), [N15](07-native-content.md#n15), [N16](07-native-content.md#n16), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/chrome/table-merges.ts`; `sibling table-merges browser test`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25), [S28](sources.md#s28)

**Change:** Connect valid cell selections to merge and split. Show native refusal reasons for invalid selections.

**Close when:** Merge and split horizontal and vertical selections. Undo restores cell content. Invalid overlap performs no write.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `e479455c`, completed with N14 in one private table-controls batch. `table-merges.ts` connects native rectangle merging and existing-grid unmerge. Unmerge derives the complete visual cell extent through the native selection model and executes the existing native split operation. Unequal original columns keep their boundaries and widths. Moved paragraph identities and content remain in the first restored cell, matching the captured application unmerge semantics. Explicit partial overlaps across an existing merged cell refuse atomically with `cells-not-rectangle`. The toolbar exposes native refusal reasons. Viewing refuses a valid unmerge. Suggesting refuses structural changes without changing the mode.

The shared validator uses native validation without allocating a last-operation preview. Execution uses `surface.applyTableCommandPlan`, which retains native mode/lock gates, story scope, selection and one gesture per undo step. No published API or shared operation changes. Local cases cover horizontal, vertical and rectangle merges, unequal-grid unmerge, exact full-document undo/redo and save/reopen, invalid overlaps, and a viewing editor. Six real-session cases cover both client orders, offline edits in moved paragraphs, native node identity, independent own/peer history, reconnect, changed-grid remote carets/selections, existing-grid split, deletion while a peer edits the moved text, and exact restoration/export. The final combined command `bun test --timeout 30000 --preload ./scripts/test/preload-dom.ts packages/sciobot/src/chrome/table-merges.test.ts packages/sciobot/src/chrome/table-appearance.test.ts packages/sciobot/src/chrome/table-style-collaboration.test.ts packages/sciobot/src/diff/table-grid-split-collaboration.test.ts packages/sciobot/src/diff/table-merge-index-collaboration.test.ts packages/sciobot/src/diff/table-format-collaboration.test.ts packages/sciobot/src/review/owned-presence-shared.test.ts` passes 36 cases and 7596 assertions across seven files (`/tmp/sciobot-table-style-integration.log`). It includes earlier payload/index/property and owned AI presence regressions. `bunx playwright test --config e2e/sciobot-table-style.config.ts` passes seven cases (`/tmp/sciobot-table-style-browser-final.log`): three merge/split actions, three appearance cases, and a real native rectangle drag consumed by the merge control. Real typing, exact history and canonical Word output remain required, with original deadlines unchanged. Private types pass 2336 files; browser types pass 1907 files. The private build verifies 303 modules and the external-core/production-import boundary. Compatibility: `.collaboration/changes/sciobot-table-style-controls.json`. This private adapter change requires no new published API snapshot or package release. Evaluation transport remains test-only under the EigenPal Pro License. The backlog audit retains its external application-source drift; no capture command runs.

## N14

- [x] Connect manual table appearance controls

**Depends:** [N09](07-native-content.md#n09), [N17](07-native-content.md#n17), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/chrome/table-appearance.ts`; `sibling table-appearance browser test`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Connect width, row height, border, background, and header options to native property commands.

**Close when:** Change one cell background, one column width, and one row height. Reopen preserves all three without changing neighboring cell content.

**Check:** Add a focused Chromium case with real input and native selection where applicable. Add the focused two-session scenario described in the shared completion rules.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Baseline `e479455c`, completed with N13. `table-appearance.ts` connects selected-cell RGB background/clear, selected-column widths, exact row heights, all/clear borders and header-row count. Native points convert explicitly to twips in the toolbar. Width changes retain other columns and recalculate merged cell widths through the existing native operation. Native transactions retain selection, neighboring paragraph subtrees, imported properties and one-step history. Invalid values and a viewing editor refuse before writes. No public API, shared fields, codecs or document tuple changes.

Local cases cover unmerged, horizontal-merged and vertical-merged tables with nonuniform imported widths. Independent native-property assertions require widths `1200/2100/2400`, background `1A2B3C`, exact row height `720`, four single borders with size `8` and color `123456`, and two header rows. Every original paragraph subtree stays exact. Full canonical fingerprints verify each undo and redo, followed by exact Word save/reopen. Six real-session cases cover both client orders, offline peer text during all five actions, independent own/peer history, changed-layout remote caret/range geometry, reconnect, deletion of the affected row during peer typing, exact restoration and native export. The shared integration command and browser command recorded in N13 pass 36 native cases/7596 assertions and seven Chromium cases. Three Chromium appearance cases use actual color/number inputs, pointer actions, native typing, exact per-gesture history, and Word output for all three merge fixtures. Private/browser types pass 2336/1907 files. The private build verifies 303 modules, external core, and production imports. Compatibility: `.collaboration/changes/sciobot-table-style-controls.json`. Existing native property operations and N17 supply the engine behavior; this private chrome needs no new published API snapshot or package release. Evaluation transport remains test-only under the EigenPal Pro License. The backlog audit retains external application-source drift; no capture command runs.

## N15

- [x] Support splitting along existing native grid columns

**Depends:** [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/store/tree-op-table-split.ts`; `fork:packages/core/src/store/store/tree-op-types.ts`; `fork:packages/core/src/store/__tests__/table-merge-split-ops.test.ts`; `fork:packages/sciobot/src/diff/table-grid-split-collaboration.test.ts`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25), [S28](sources.md#s28)

**Change:** Add an explicit existing-grid column distribution to native cell splitting. Preserve the default equal-width split. Reuse the native row-band split and content ownership logic. Preserve original grid widths, column IDs, unrelated cells, paragraphs, payload controls, and imported XML. Reject incompatible distributions and invalid geometry before writes. Keep native transaction and collaboration handling unchanged.

**Close when:** A merged cell across unequal original grid columns unmerges into those original slots without introducing cuts or changing widths. All text remains in the first cell. Horizontal, vertical, combined, and unknown-width fixtures retain intended structure through save/reopen. Existing equal-width split tests remain valid. Two real sessions converge through concurrent text edits, undo/redo, reconnect, deletion during peer editing, and remote selection geometry.

**Check:** Run focused native regressions and a two-session regression in both delivery orders. Rebuild core, check API snapshots, update feature/architecture docs, and record a changeset and collaboration decision.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local branch `codex/react-block-proof`. Native `splitTableCell` accepts optional `columnDistribution: existing-grid` and retains default equal-width behavior. Original unequal grid slots, widths, column IDs, owner content/paragraph IDs, and imported XML survive restoration. Combined horizontal/vertical, grouped-slot, narrow-slot, missing-grid, unknown-width, and invalid-count cases pass. Native file: 28 tests, 92 assertions. From the repository root, `bun test packages/core/src/store/__tests__/table-merge-split-ops.test.ts packages/core/src/store/__tests__/table-op-topology.test.ts packages/core/src/automation/__tests__/table-authoring.test.ts` passes 88 cases and 284 assertions. `bun test --preload ./scripts/test/preload-dom.ts packages/sciobot/src/diff/table-grid-split-collaboration.test.ts` passes both real-session delivery orders and 144 assertions. It proves concurrent edits, peer-preserving split undo/redo, reconnect, payload/control retention, remote geometry in retained and fresh cells, concurrent deletion, peer-text restoration through deletion undo, redo, and saved DOCX convergence. The evaluated transport remains test-only under the EigenPal Pro License. Final full private suite: 557 pass, the same two R29 failures, 559 tests, 53 files, 7526 assertions. Core/private typecheck/build, scoped formatting/lint, root lint, scoped core API check, compatibility, and diff checks pass. Core API: 15 entries, zero errors, baseline warnings only. Private build: 75 modules, one external core, React-free main entry. Minor changeset: `.changeset/existing-grid-cell-split.md`. Compatibility: `.collaboration/changes/existing-grid-cell-split.json`, 28 decisions; no shared format change. Architecture, table guide, feature matrix, and store API snapshot updated. All 196 application source records remain equal at `be76538c`. D10 retains pure plan composition; N13/R10 retain UI/review integration. See [implementation notes](evidence/implementation-notes.md).

## N16

- [x] Retain native indexes when subtrees move across parents

**Depends:** [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/package/ooxml-edit.ts`; `fork:packages/core/src/store/__tests__/ooxml-edit-index.test.ts`; `fork:packages/sciobot/src/diff/table-merge-index-collaboration.test.ts`

**Read:** [S25](sources.md#s25), [S28](sources.md#s28)

**Change:** Make native index updates finish old-location cleanup before registering new locations. Preserve moved subtree node/parent lookups, leaf reads, allocator reservations, duplicate detection, and bounded incremental behavior. Do not rebuild entire indexes for ordinary edits. Keep canonical output and shared primitive formats unchanged.

**Close when:** Vertical and rectangle cell merges retain every moved paragraph/run/control ID in live native reads. An independent tree walk matches node membership, objects, and parents after moves in both directions, preview reuse, history, and subsequent text edits. Two real sessions retain moved content through concurrent editing, undo/redo, reconnect, deletion during peer edits, remote geometry, and save/reopen.

**Check:** Add focused native index regressions and both two-session delivery orders. Run existing index, merge/split, and relevant editor regressions. Rebuild core, check API snapshots, update docs, and record a changeset and compatibility decision.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local branch `codex/react-block-proof`. Index updates now finish removals across the changed spine before registering new locations. Four new differential cases verify exact node/parent membership, moves in both directions, subsequent edits, preview reuse, history, duplicate refusal, source snapshots, and zero full index builds on warmed valid paths. Existing 2000-operation differential checks remain valid. From the repository root, `bun test packages/core/src/store/__tests__/ooxml-edit-index.test.ts packages/core/src/store/__tests__/ooxml-edit.test.ts packages/core/src/store/__tests__/table-merge-split-ops.test.ts packages/core/src/store/__tests__/tracked-edits.test.ts packages/core/src/store/__tests__/lossless-editing.test.ts packages/core/src/store/__tests__/table-op-topology.test.ts packages/core/src/automation/__tests__/table-authoring.test.ts` passes 183 cases and 19318 assertions. `table-merge-index-collaboration.test.ts` passes both real-session delivery orders and 4852 assertions. Live independent tree walks verify every node/parent after rectangle merge, history, and concurrent join/deletion. Peer edits, owned controls/payloads, moved selections, subsequent addressing, undo/redo, reconnect, and saved DOCX convergence pass. Core/private typecheck/build, scoped formatting/lint, root lint, scoped core API, compatibility, and diff checks pass. Full private suite: 574 pass, the same two R29 failures, 576 tests, 55 files, 12517 assertions. Core API retains 15 entries without drift. Private build: 76 modules, one external core, React-free main entry. Patch changeset: `.changeset/moved-subtree-indexes.md`. Compatibility: `.collaboration/changes/moved-subtree-indexes.json`, 30 total decisions; no shared format change. Architecture, table guide, and feature matrix updated. All 196 audited application source records remain equal at `be76538c`. See [implementation notes](evidence/implementation-notes.md).

## N17

- [x] Preserve imported metadata on native table property writes

**Depends:** [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/store/table-property-attributes.ts`; `fork:packages/core/src/store/store/tree-op-tables.ts`; `fork:packages/core/src/store/store/tree-op-table-cell-properties.ts`; `fork:packages/core/src/store/__tests__/table-property-fidelity.test.ts`; `fork:packages/sciobot/src/diff/table-format-collaboration.test.ts`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Patch operation-owned Word attributes on existing table layout, table width, cell width, row height, and cell vertical alignment elements. Keep imported foreign attributes, children, identities, sibling properties, and cell content. Preserve existing native validation and shared operation formats.

**Close when:** Divider, right-edge, row-height, and vertical-alignment changes retain imported property metadata and IDs. Full-grid changes retain their existing fidelity. Tests cover merges, content, no-op properties, and save/reopen. Two real sessions retain metadata through concurrent text edits, undo/redo, deletion, reconnect, geometry, and saved convergence.

**Check:** Run native resize and fidelity regressions and both collaboration delivery orders. Rebuild core, check APIs, update architecture/feature/table docs, and add a patch changeset and compatibility decision.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local branch `codex/react-block-proof`. Native resize and vertical-alignment writes now retain imported property IDs, foreign attributes, children, and sibling properties. Five fidelity cases join existing namespace, sizing, merge, border/fill, and automation regressions. From the repository root, `bun test packages/core/src/store/__tests__/table-property-fidelity.test.ts packages/core/src/store/__tests__/table-resize-ops.test.ts packages/core/src/store/__tests__/table-cell-properties.test.ts packages/core/src/store/__tests__/table-merge-split-ops.test.ts packages/core/src/automation/__tests__/table-authoring.test.ts` passes 136 cases and 592 assertions. Two real sessions in `table-format-collaboration.test.ts` pass both clients 73/41 delivery orders. They retain all four changed properties and peer text through reconnect, undo/redo, merged geometry, concurrent row deletion, restoration, redo, and saved DOCX convergence. Remote range rectangles match and the resized-column caret moves right. The evaluated transport stays test-only under the EigenPal Pro License. Final focused command with `table-format.test.ts`, `table-format-collaboration.test.ts`, and `plan.test.ts` passes 31 cases and 1745 assertions. Full private suite: 594 pass, the same two R29 failures, 596 tests, 57 files, 14169 assertions. Final focused geometry adds twelve assertions after that full run. Core/private typecheck/build, scoped formatting/lint, root lint, diff, scoped core API, and compatibility checks pass. Core API: 15 entries, zero errors, baseline warnings only. Private build: 77 modules, one external core, React-free main. Patch changeset: `.changeset/table-property-metadata.md`. Compatibility: `.collaboration/changes/table-property-metadata.json`, 32 decisions; no shared format changes. Architecture/table/feature docs updated. All 196 audited source hashes still match `be76538c`. Application HEAD advanced externally to `9693712d`; a verified revision-only inventory refresh leaves every source and mapping record equal. See [implementation notes](evidence/implementation-notes.md).

## N18

- [x] Establish a native inline-code representation and projection

**Depends:** [C03](01-contracts.md#c03), [P14](03-agent-projection.md#p14), [V01](11-verification.md#v01), [N20](07-native-content.md#n20)

**Primary files:** `fork:packages/sciobot/src/native/inline-code.ts`; `sibling inline-code.test.ts`; `fork:packages/docx-to-markdown/src/markdown-inline.ts`; `focused code-span projection regression`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Define a Word-preserving native representation for an application-authored code span. Apply and remove it through canonical operations. Project that explicit identity as Markdown code without treating ordinary monospaced text as code. Preserve imported character styles and unknown run properties.

**Close when:** Code survives native formatting, exact Markdown parsing, undo/redo, and DOCX save/reopen. Backticks, XML characters, surrogate pairs, empty transitions, links, and neighboring marks preserve their intended text and addresses. Ordinary imported monospaced runs stay ordinary text.

**Check:** Add native semantic and Markdown AST oracles. Use two real sessions for concurrent peer edits, review visibility, history, geometry, reconnect, and saved convergence. Keep public Apache APIs and the single-core boundary.

**Claim:** codex, branch `codex/react-block-proof` (continuing Claude). **Evidence:** Native baseline `9770e78` plus private/application snapshot `0287a034937b7b98294415098bdeaa410cd5736f`. Application-authored `SciobotCode` identity uses canonical `setRunProperties.characterStyle`. Derived styles retain imported character style inheritance and restore its exact original ID on removal. Unknown run properties remain intact. Read-only native Markdown uses explicit code style IDs, literal code content, exact UTF-16 source mapping, and AST oracles. Ordinary imported monospaced runs stay ordinary text. Backticks, XML characters, emoji, empty transitions, links, adjacent marks, typing, stale refusals, tracked formatting decisions, and saved reopen pass. The 66-case/1030-assertion N20 command includes native semantics, converter source maps, core character-style writes, direct planner execution, and 32 real-session cases. Those sessions cover concurrent text/code, imported styles, both client and delivery orders, review decisions, undo/redo, deletion during peer editing, remote caret/selection geometry, reconnect, DOCX and shared-room reopen. `/tmp/codex-n20-native-final.log`. N20 closes infrastructure history through production pre-initialization; `/tmp/codex-n20-browser.log` adds 13 passing application Chromium cases. Private/application typecheck/build, scoped format/lint, core/converter API checks, and compatibility pass. Published changesets: `.changeset/run-character-style-writes.md` and `.changeset/markdown-code-character-styles.md`; existing API snapshots match. Decisions: `.collaboration/changes/run-character-style-writes.json`, `.collaboration/changes/markdown-code-character-styles.json`, `.collaboration/changes/sciobot-inline-code.json`, and N20 initialization decision. No production Pro runtime, parallel editor state, or Markdown reconstruction. N03 retains toolbar/shortcut wiring; R38 retains tracked paragraph lowering. Actual desktop Word evidence remains V06.

## N19

- [ ] Address native table columns by reference cell in any grid state

**Depends:** [C03](01-contracts.md#c03), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/core/src/store/store/tree-op-tables.ts`; `fork:packages/core/src/store/store/tree-op-types.ts`; `focused core column operation test`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Let `insertTableColumn` and `deleteTableColumn` address a column through a first-row reference cell whether `w:tblGrid` is absent, incomplete, or complete. Synthesize or repair the canonical grid inside the operation.

**Close when:** Several reference-cell insertions and deletions apply in sequence to gridless and partially gridded tables, survive undo/redo and save/reopen, and keep surviving cell identities.

**Check:** Add core operation tests, two real sessions for concurrent peer edits and reconnect, a core changeset, API snapshot updates, and a compatibility decision for the shared operation change.

**Claim:** unclaimed. **Evidence:** pending.

## N20

- [x] Keep concurrent auxiliary-part creation undoable

**Depends:** [C03](01-contracts.md#c03), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/native/inline-code.ts`; `fork:packages/sciobot/src/native/inline-code-collaboration.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Make concurrent first code writes safe when the original DOCX lacks a styles part. Preserve both actors' run formatting and shared style definitions through history. Keep the evaluated transport test-only; implement the production fix through Apache operations or private application initialization.

**Close when:** Both peers create code offline, reconnect in either order, and retain one valid shared styles part. One undo removes only its actor's code and preserves the peer's code. Redo, further edits, and saved reopen converge with both sessions ready.

**Check:** Run the `concurrent first code writes` regressions in `packages/sciobot/src/native/inline-code-collaboration.test.ts`. Extend them through redo, subsequent writes, both client orders, and saved reopen.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, private/application snapshot `refs/backup/codex-n20-seeded-vendor` (`0287a034937b7b98294415098bdeaa410cd5736f`), application branch `claude/docx-native-editor`. `prepareNativeDocument` creates the code style and derived imported-character-style definitions before document or room initialization. Application raw imports and blank documents use `prepareWorkspaceDocument`; the room creator initializes shared bytes before bootstrap. Live handles pass through unchanged. Preparation preserves every unrelated part byte-for-byte and never runs Word export cleanup. Case-insensitive occupied filenames allocate a fresh styles path. Conflicting imported identities stay openable and untouched. Eight real-session first-write cases cover both client orders, reconnect receivers, and undo owners. Four derived-style cases reproduce and prevent a second defect: later deduplication followed by peer undo removed the surviving derived definition. Preparation gives these definitions no actor history. History starts empty; code edits undo independently, redo, accept subsequent writes, and reopen as DOCX or complete shared-room state with both sessions ready. Command: `bun test --preload ./scripts/test/preload-dom.ts --timeout 60000 packages/sciobot/src/native/inline-code.test.ts packages/sciobot/src/native/prepare-document.test.ts packages/sciobot/src/native/inline-code-collaboration.test.ts packages/sciobot/src/diff/target-inline-code.test.ts packages/core/src/store/__tests__/run-character-style.test.ts packages/docx-to-markdown/test/markdown-code-spans.test.ts`; 66 pass, 1030 assertions, `/tmp/codex-n20-native-final.log`. Application Chromium command: `DOCX_NATIVE_SNAPSHOT=/home/lofcz/.cache/sciobot-docx-native/fork-n20 DOCX_HARNESS_PORT=8098 DOCX_HARNESS_URL=http://127.0.0.1:8098 bun test --timeout 600000 scripts/docx-native/harness/inline-code.test.ts scripts/docx-native/harness/workspace.test.ts scripts/docx-native/harness/placeholders.test.ts`; 13 pass, 93 assertions, `/tmp/codex-n20-browser.log`. Five new cases use real native selection, buttons, undo/redo, save, and reopen. Private/application typechecks pass (1733/509 files); scoped formatting, lint, diff, build, and compatibility checks pass. Build: 132 private modules, one external core. Decision: `.collaboration/changes/sciobot-native-document-initialization.json`; 92 decisions. The evaluated transport remains unchanged and test-only under the EigenPal Pro License. Its low-level auxiliary-part defect still reproduces when initialization is bypassed; production uses the explicitly permitted private initialization fix.

## N21

- [x] Prepare native documents without a document relationship part

**Depends:** [N20](07-native-content.md#n20), [R44](05-suggestions.md#r44), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/native/prepare-document.ts`; `sibling prepare-document.test.ts`; `sibling prepare-document-collaboration.test.ts`

**Read:** [S07](sources.md#s07), [S18](sources.md#s18)

**Change:** Initialize native review infrastructure before code styles when a valid DOCX lacks the main document relationship part. Retain original body/furniture/media bytes. Create the missing auxiliary relationship part during preparation, before editor or room bootstrap. Preserve refusal for malformed or opaque relationship bytes. Keep initialization idempotent and outside user history.

**Close when:** Full preparation accepts a valid package without `word/_rels/document.xml.rels`. Code styles and the group carrier resolve through its new relationships. Native story bytes stay exact. Both actors can edit and review independently after bootstrap. Undo, redo, reconnect, deletion, geometry, and saved reopen preserve both actors and their infrastructure. The installed bundle smoke uses the original absent-part fixture and passes.

**Check:** Add exact-byte and relationship-discovery regressions. Exercise two real sessions in both client orders and reconnect directions. Verify installed bundles, package hashes, and application regressions after the immutable vendor refresh.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Full preparation now initializes the existing group carrier before code styles. The carrier's native helper creates missing owner relationships. The code then adds styles through that valid part. Original body, media, and unrelated bytes stay exact; opaque existing relationships still refuse without replacement. The original installed smoke fixture stays unchanged and passes full preparation without a workaround. Preparation is idempotent and creates no user history. Focused preparation and code checks pass 44 cases with 936 assertions across three files; `/tmp/codex-n21-focused-final.log`. Four new real-session cases cover both client orders and reconnect directions, simultaneous code/review edits, peer-preserving undo/redo, deletion during peer input, restored peer text, remote geometry, save/reopen, and retained initialized parts. Two new local cases cover absent and opaque owner relationships. Full private suite passes 1466 cases, zero failures, 29863 assertions, 126 files; `/tmp/codex-n21-private-full.log`. Private typecheck passes 1755 files; the 140-module build retains one external core and a React-free main. Scoped format/lint, compatibility, source coverage, and backlog pass. Decision: `.collaboration/changes/sciobot-native-owner-relationship-initialization.json`; 99 decisions, with the earlier one-field format change unchanged. No published API, shared schema, provider, evaluated implementation, or database changes. The app pins immutable `refs/backup/codex-n21-runtime`, commit `f8a2d15e1a686fd504d8ae6255cab3d5dfca68a1`, tree `44ba09b1edb1bb7922015b354febf1476006bde0`. The adapter archive is `sciobot-docx-fdf6552c585c.tgz`. Published sources, shared build scripts, manifests, and dependency locks are byte-identical to R14 runtime `b69b8e8b714a7818e12ee8f636d6b75c2f8c18bb`; its four verified published archives are reused unchanged. The reuse guard rejects R13 artifacts because native build inputs differ. Core archive: `docx-editor.dev-core-0247acf90551.tgz`. Three other published archives also retain their R13 hashes. Installed-bundle smoke passes exact native preparation bytes, stable replacement groups, native accept/reject, retained-ID reservation, undo/redo, and DOCX reopening; `/tmp/codex-n21-installed-smoke.log`. App regressions pass 53 unit cases and 16 browser cases with 169 assertions. App/harness typechecks pass 6629/2195 files. Production build passes with the existing MapLibre warning. All five archive hashes, manifest/override/lock references, and single-core resolution match. App changes are restricted to the manifest, lock, and vendor artifacts. All 196 audited production source hashes remain unchanged.

## N22

- [x] Connect manual table indentation

**Depends:** [N09](07-native-content.md#n09), [N17](07-native-content.md#n17), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/chrome/table-indent.ts`; `sibling table-indent browser test`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Map the production table left-margin action to native table indentation. Convert captured CSS pixels into Word twips. Preserve imported table properties and neighboring content.

**Close when:** The N01 `table.margin` fixture maps 24 pixels to 360 twips. The table moves by 18 points without changing its columns or text. Reset restores zero indentation. Undo and save/reopen preserve the expected indentation and imported metadata.

**Check:** Add a focused native regression and Chromium control test. Verify two-session concurrent cell edits, reconnect, peer-preserving history, table deletion, remote selection geometry, and saved convergence. Reuse the native table property writer where it supports `w:tblInd`.

**Claim:** codex solo, branch `codex/react-block-proof`. **Evidence:** Implemented from `fef530df` on `codex/react-block-proof`. `chrome/table-indent.ts` maps captured 24 CSS pixels to 360 Word twips through native table-property transactions. Reset writes zero. The core writer preserves imported attributes, extension children, table widths, merged cells, and content identities. `test/support/table-indent.ts` supplies independent complete Word source oracles. Native suggesting now attributes all supported table-property operations and honors the formatting-tracking setting.

`bun test --preload ./scripts/test/preload-dom.ts` with the eight focused files in `/tmp/sciobot-n22-final-verified-native.log` passes 118 cases and 1589 assertions. `table-indent.test.ts` verifies exact source fingerprints, positive/negative indentation, 18-point movement, reset, canonical Word reopen, unchanged selection, complete history, refusal, and tracked decisions. `table-indent-collaboration.test.ts` covers both client orders, plain/merged tables, offline peer edits, independent history, reconnect, table deletion recovery, presence geometry, tracked accept/reject, decision history, and saved convergence.

`bunx playwright test --config e2e/sciobot-table-indent.config.ts` passes two Chromium cases (`/tmp/sciobot-n22-final-browser.log`). Real controls move painted paragraphs by 24 pixels, preserve native selection, support typing, reset, undo/redo, and reopen equal Word documents. Core, private, and browser type checks pass 1613, 2348, and 1920 files. Core/private builds pass; the private package keeps core external across 306 modules. Scoped lint and formatting pass. The public store API snapshot adds only `indentTwips`. Compatibility decision: `.collaboration/changes/native-table-indentation.json`; changeset: `.changeset/native-table-indentation.md`. Transport remains test-only under the EigenPal Pro License.

## N23

- [x] Define explicit native horizontal-rule identity and projection

**Depends:** None.

**Primary files:** `fork:packages/sciobot/src/native/horizontal-rule.ts`; `fork:packages/docx-to-markdown/src/markdown.ts`; focused representation and projection tests

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Represent application rules as empty Word paragraphs with the reserved `SciobotHorizontalRule` style and matching `bottom` and `between` single borders. Word paints one closing line per marker, including adjacent grouped paragraphs. Add opt-in Markdown projection with exact generated syntax extents. Preserve ordinary imported borders and refuse conflicting style initialization.

**Close when:** The explicit native paragraph and style survive save/reopen. Only opted-in empty paragraphs with a visible single closing bottom or between border project as `---`. Ordinary imported borders and conflicting nonempty/boxed paragraphs stay ordinary. Source maps bind generated syntax to the zero-length native paragraph.

**Check:** Run immutable representation and converter tests, option validation, scoped types/lint/format, converter build/API check, changeset, and compatibility policy. N02 retains live command, browser, and collaboration verification; D34 retains thematic-break insertion planning.

**Claim:** astra N23, branch `codex/react-block-proof`. **Evidence:** Immutable native helper adds the reserved paragraph style without body writes and refuses conflicting definitions or unresolved style relationships. The exact empty marker has matching 0.75-point bottom/between edges; unknown content and changed properties fail closed. Private Markdown projection validates the supplied native package before opt-in; `projectLists` forwards that package through working and committed reads. General converter callers opt in explicitly and own source identity validation. Native helper checks pass 7 cases, including save/reopen, untouched imported metadata, mixed invalid markers, generated-only syntax, and both review readers. Final converter command `bun test --preload ./scripts/test/preload-dom.ts packages/docx-to-markdown/test/markdown-horizontal-rule.test.ts` passes 6 cases and 36 assertions (`/tmp/n23-final-converter.log`). Independent native paint assertions cover one marker, adjacent grouped markers, and markers separated by a page break. A detached two-page continuation fixture proves one logical owner and closing stroke; it does not claim that an authored empty one-line paragraph naturally splits. Earlier adjacent regressions passed all 38 unchanged code-span, logical source-map, list-projection, and Markdown-projection cases. Nine converter evolution-boundary checks pass. Scoped lint and formatting pass. Converter build and API extraction/check pass. The API snapshot adds only horizontalRuleParagraphStyles. Private types pass across 1975 files; the private build verifies 196 modules with external core. Repository lint passes. Compatibility policy passes 129 decisions with zero format fields changed against the branch merge-base fc28dba928b073c2beddc5672cdc7b42ea3af0f9. Minor changeset `.changeset/explicit-horizontal-rule-projection.md`; compatibility decision `.collaboration/changes/sciobot-horizontal-rule-representation.json`; package README and API documentation describe the opt-in contract. This leaf adds no live command or planner insertion. N02 and D34 retain those behaviors and their browser/two-session checks.
