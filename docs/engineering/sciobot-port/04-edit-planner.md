# Small canonical edit plans

Produce pure plans first. Direct and suggestion executors consume the same plans.

Read the [dispatch and completion rules](../../../SCIOBOT_PORT_TODO.md) before claiming a leaf.

Each named test or artifact is a deliverable. These specifications do not claim those checks already exist.

## D01

- [x] Define a pure canonical edit plan

**Depends:** [C07](01-contracts.md#c07), [P10](03-agent-projection.md#p10), [P06](03-agent-projection.md#p06)

**Primary files:** `fork:packages/sciobot/src/diff/plan.ts`; `sibling plan.test.ts`

**Read:** [S03](sources.md#s03), [S18](sources.md#s18), [S25](sources.md#s25)

**Change:** Define text, format, sibling, table, and payload operations with expected revisions. Separate planning from transaction application.

**Close when:** An empty diff emits no operations. Serializing a plan contains no live editor references or reconstructed document tree.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `diff/plan.ts` defines versioned mode-independent text, format, sibling, table, and payload edits with expected package revisions. Missing old properties remain distinct from explicit null values. No-op text/property/sibling changes emit no operations. Plans preserve canonical addresses and unchanged sibling anchors; insertions contain only new target AST fragments. Native table commands include atomic row/content authoring without review attribution. Schema checks reject whole-story replacement commands, canonical trees, live editors, getters, toJSON, cycles, invalid anchors, hidden root state, unknown versions, and excessive resources. Construction clones and deeply freezes data without freezing caller inputs. `plan.test.ts` and payload-schema regressions pass 18 focused tests and 114 assertions. The full private suite passes 212 tests and 1,990 assertions across 24 files. Package typecheck, external-runtime build, scoped Prettier, repository lint, and diff checks pass. Shared JSON validation preserves existing payload behavior. This pure plan constructor applies no mutations. Executors still validate live revision, addresses, old values, protection, geometry, and supported review modes before application.

## D02

- [x] Plan edits within one paragraph

**Depends:** [D01](04-edit-planner.md#d01), [P06](03-agent-projection.md#p06)

**Primary files:** `fork:packages/sciobot/src/diff/paragraph-text.ts`; `sibling paragraph-text.test.ts`

**Read:** [S03](sources.md#s03), [S18](sources.md#s18)

**Change:** Compute minimal text insert, delete, and replacement ranges. Keep unchanged runs and preserve UTF-16 offsets.

**Close when:** One changed word retains surrounding run IDs. Emoji edits never split surrogate pairs.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `paragraph-text.test.ts` and `paragraph-text-collaboration.test.ts` pass 16 tests and 789 assertions. Minimal insert/delete/replacement plans use original UTF-16 coordinates in descending order and the single native grapheme strategy. One hundred generated Unicode cases reconstruct their targets exactly. Emoji, ZWJ sequences, combining marks, mapped fragments, native strategy changes, malformed boundaries, and resource limits have explicit checks. Direct lowering inserts before deleting old text and retains insertion bias; prefix, suffix, and longer-word regressions preserve neighboring run IDs, unknown properties, and one-step undo/redo. Native tabs, manual breaks, and hyphens retain their model semantics. Two real sessions verify synchronization, remote caret offsets/geometry, concurrent offline edits, peer-preserving undo/redo, reconnect, deletion during peer editing, stale-plan refusal, and save/reopen convergence. The full private suite passes 228 tests and 2,780 assertions; the final focused run adds one malformed-strategy assertion. Package typecheck, external-runtime build, scoped formatting, repository lint, compatibility policy, and diff checks pass. Compatibility decisions live in `.collaboration/changes/sciobot-paragraph-text-planning.json` and the paint-observer decision linked to its existing minor changeset. No shared format changes come from this private planner. A01/A02 retain production preflight and application; review executors consume the same semantic plan through their dedicated lowering.

## D03

- [x] Plan inline formatting and hyperlink edits

**Depends:** [D02](04-edit-planner.md#d02)

**Primary files:** `fork:packages/sciobot/src/diff/run-format.ts`; `sibling run-format.test.ts`

**Read:** [S03](sources.md#s03), [S18](sources.md#s18)

**Change:** Plan changed mark properties and link targets independently from text. Keep unchanged run properties and relationships.

**Close when:** Bold-to-italic and URL-only fixtures change only their intended properties. Direct and tracked executors receive the same semantic plan.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `packages/sciobot/src/diff/run-format.test.ts` passes 10 tests and 64 assertions. `planRunFormat` and `planHyperlinkFormat` return immutable D01 property deltas without text, relationship, or document writes. Bold-to-italic changes only b/i; native lowering retains neighboring run IDs, fonts, colors, character styles, unknown XML, and text. The same semantic plan produces identical direct and accepted tracked native fingerprints; rejection restores the original. Native helpers preserve font slots/themes, compound attributes, complex-script companions, clipped UTF-16 ranges, and Original/Proposed visibility. Property removal preserves absence. URL-only plans address one wrapper, preserve its tooltip and marks, and leave shared relationships unchanged. Pure native retargeting changes only the selected link. Unsafe URLs, invalid properties, getters, split graphemes, missing targets, and unsupported field-result ownership refuse without source writes. Full private suite: 272 tests, 3114 assertions, 29 files. Private typecheck, external-runtime build with 33 private modules, scoped Prettier, root lint, diff check, and compatibility policy pass. Compatibility record: `.collaboration/changes/sciobot-inline-format-planning.json`. This pure planner requires no collaboration mutation test. D15 retains Markdown-to-native composition; A01/R06 retain guarded execution. Native tracked hyperlinks lack attribution, so R26/R27 explicitly retain that required capability and adapter. See [implementation notes](evidence/implementation-notes.md).

## D04

- [x] Plan paragraph property changes

**Depends:** [D01](04-edit-planner.md#d01), [P06](03-agent-projection.md#p06)

**Primary files:** `fork:packages/sciobot/src/diff/paragraph-format.ts`; `sibling paragraph-format.test.ts`

**Read:** [S03](sources.md#s03), [S18](sources.md#s18)

**Change:** Plan heading, alignment, indentation, and spacing changes. Preserve unrelated styles and section properties.

**Close when:** A heading change keeps text and block identity. A spacing change retains existing borders and tabs.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `packages/sciobot/src/diff/paragraph-format.test.ts` passes 6 tests and 31 assertions. `planParagraphFormat` returns immutable D01 deltas over authored paragraph values without native document writes. Heading/style changes retain paragraph and run identities, text, and unrelated XML. Direct and accepted tracked native fingerprints match; rejection restores the original fixture. Spacing changes retain line/after settings, tab-stop children, borders, section properties, unknown XML, and style by identity/fingerprint. Attribute patches preserve unrelated settings; null removes one attribute, such as a conflicting hanging indent. Removal retains explicit property absence. Alignment, direction, toggles, no-ops, invalid values, unsupported structural properties, missing targets, and getters have explicit checks. Full private suite: 278 tests, 3145 assertions, 30 files. Private typecheck, external-runtime build with 34 private modules, scoped Prettier, root lint, diff check, and compatibility policy pass. Compatibility record: `.collaboration/changes/sciobot-paragraph-format-planning.json`. This pure planner needs no collaboration mutation test. Callers resolve native style IDs and units; D15 retains Markdown composition. A01/R06 retain guarded execution and collaboration. Numbering and tab-stop authoring remain separate leaves. See [implementation notes](evidence/implementation-notes.md).

## D05

- [x] Plan sibling block insertion and deletion

**Depends:** [D01](04-edit-planner.md#d01), [P06](03-agent-projection.md#p06), [C08](01-contracts.md#c08)

**Primary files:** `fork:packages/sciobot/src/diff/sibling-blocks.ts`; `sibling sibling-blocks.test.ts`

**Read:** [S03](sources.md#s03), [S18](sources.md#s18)

**Change:** Pair stable siblings and emit bounded insertion or deletion operations. Preserve unchanged neighbors and nested parent identity.

**Close when:** C01 section insertion and deletion pairs leave surrounding canonical node IDs unchanged.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `packages/sciobot/src/diff/sibling-blocks.test.ts` supplies 13 focused cases. The captured C01 section insertion and deletion cases retain surrounding canonical IDs, shared native objects, unknown XML fingerprints, and parent identity through pure native edit oracles. The planner validates contiguous native windows, retains exact siblings and compatible changed content pairs, and emits bounded structural gaps with original anchors in reverse order. Changed pairs remain explicit work for D15 content/property planning. Nested containers, partial windows, opaque surrounding nodes, terminal sections, duplicates, edge insertions, complete removals, multiple gaps, detached immutable results, invalid identities, getters, native-tree rejection, and resource limits pass. One hundred generated cases reconstruct target order exactly. Shared target-fragment validation retains all D01 plan regressions. The full private suite passes 291 tests and 4020 assertions across 31 files. Private typecheck, external-runtime build with 36 private modules, scoped Prettier, root lint, diff check, and compatibility policy pass. Compatibility record: `.collaboration/changes/sciobot-sibling-block-planning.json`; no shared formats changed. This pure planner needs no collaboration mutation test. D12 retains domain identity pairing; A01/R05 retain execution, native structural invariants, collaboration, and review persistence. See [implementation notes](evidence/implementation-notes.md).

## D06

- [x] Plan list membership changes

**Depends:** [D04](04-edit-planner.md#d04), [D05](04-edit-planner.md#d05)

**Primary files:** `fork:packages/sciobot/src/diff/lists.ts`; `sibling lists.test.ts`

**Read:** [S03](sources.md#s03), [S18](sources.md#s18)

**Change:** Plan bullet, numbered, checklist, and nesting changes. Treat automatic ordinal renumbering as derived behavior.

**Close when:** Inserting item two produces one content change. Unchanged later list items do not become AI changes.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, application fixture `da9081ad`, local working tree. `planListMembership` emits detached D01 list intents for bullet/ordered membership, removal, levels 0-8, shared new-instance keys, explicit authored starts, and owned checklist creation/state/nesting/removal. Native style and numbering readers preserve inherited membership and relationship-selected definitions. Derived ordinal/listStart inputs are excluded. The captured C01 item-two case produces one D05 structural change, no existing-item membership changes, native ordinals 1/2/3, and identical later paragraph objects. Native numbering oracles retain text and properties; accepted tracked results match direct edits, and rejection restores originals. Checklist state writing verifies both native metadata and the exported glyph; ordinary form controls remain separate. Proposed wrapper ownership, malformed inputs, conflicts, limits, accessors, immutable serialization, and live revision/selection/history preservation pass. Focused lists/plan/sibling command: 33 cases, 1046 assertions. Full private suite: 423 pass and the same two R29 failures across 425 tests, 43 files, and 6254 assertions. Private typecheck/build, scoped Prettier/oxlint, root lint, diff checks, and compatibility policy pass. The build verifies 59 private modules and one external core. No-impact decision: `.collaboration/changes/sciobot-list-planning.json`. This pure planner needs no shared mutation scenario. D15/A01 retain composition, new-resource allocation, and guarded execution. P21 and R30 explicitly retain task-syntax mapping and tracked checklist decisions; planning does not establish those capabilities. The production coverage snapshot changed only List/Suggestion task mappings; all 196 source hashes remain equal. See [implementation notes](evidence/implementation-notes.md).

## D07

- [x] Plan text and marks inside table cells

**Depends:** [D02](04-edit-planner.md#d02), [D03](04-edit-planner.md#d03), [P03](03-agent-projection.md#p03)

**Primary files:** `fork:packages/sciobot/src/diff/table-cells.ts`; `sibling table-cells.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Resolve each cell diff within its table and row. Preserve grid, merges, and neighboring cells.

**Close when:** Editing one cell in the C01 table fixture changes no other cell IDs or properties.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `packages/sciobot/src/diff/table-cells.test.ts` passes 9 tests and 54 assertions. The pure planner verifies same-revision P03 metadata, native table/row/cell ancestry, and nearest paragraph cell ownership before composing D02 text and D03 run-format deltas. The C01 cell replacement subcase preserves every other native cell object, ID, and fingerprint, plus table properties/grid, row properties, and affected cell properties. Formatting precedes text replacement and preserves requested marks. Multi-paragraph cells, nested tables, content-control children, visible merged cells, empty cells, Unicode/tabs/manual breaks, and no-ops pass. Nested paragraphs cannot enter outer-cell plans even through forged metadata. Stale projections, duplicate patches, cross-cell targets, hidden merge continuations, overlapping/invalid ranges, getters, and ambiguous identities refuse without source writes. Full private suite: 300 tests, 4074 assertions, 32 files. Private typecheck, external-runtime build with 37 private modules, scoped Prettier, root lint, diff check, and compatibility policy pass. Compatibility record: `.collaboration/changes/sciobot-table-cell-planning.json`; no shared formats changed. This pure planner needs no collaboration mutation test. D15 retains target decomposition; A01/R07 retain guarded execution, collaboration, and review persistence. D08-D11 retain structural changes. See [implementation notes](evidence/implementation-notes.md).

## D08

- [x] Plan table row insertion and deletion

**Depends:** [D05](04-edit-planner.md#d05), [D07](04-edit-planner.md#d07), [R31](05-suggestions.md#r31)

**Primary files:** `fork:packages/sciobot/src/diff/table-rows.ts`; `sibling table-rows.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Plan native row insertion/deletion with plain cell values and inherited row properties. Keep table identity and existing row references. [D16](04-edit-planner.md#d16) retains rich inserted cell fragments and explicit new row properties.

**Close when:** The C01 added row retains the two original cells. Deleting a row produces no operations on another table.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `planTableRow` returns existing native row insertion/deletion intents, including optional plain cell values and inherited row/cell properties. It verifies canonical table/row/reference-cell ownership and runs an immutable native preview for geometry, protection, and limits. No allocated preview IDs enter serialized plans. The C01 Three/Four row retains both original cell objects, row IDs, grid, properties, and an unrelated table. Native blank insertion extends supported vertical merges; deletion repairs surviving merge heads without editing unrelated tables. Empty/Unicode/tab values, both insertion sides, inherited header and paragraph-mark formatting, nested table scope, frozen detached plans, invalid targets, getters, limits, locks, and binding refusals pass. Planning preserves live bytes, revision, selection, and existing history. R31 fixes the native direct/accepted formatting discrepancy found by these tests. From `packages/sciobot`, `bun test --preload ../../scripts/test/preload-dom.ts src/diff/table-rows.test.ts src/diff/table-rows-collaboration.test.ts src/diff/table-cells.test.ts src/diff/plan.test.ts` passes 35 cases and 360 assertions, including 13 new pure row cases and both real-session row cases. Native core coverage passes 196 cases and 757 assertions across seven files. Full private suite: 515 pass and the same two R29 failures, 517 tests, 49 files, 6,891 assertions. Core/private typecheck/build, scoped Prettier/oxlint, root lint, diff checks, core API check, and compatibility policy pass. The private build verifies 67 source modules with one external core. Compatibility record: `.collaboration/changes/sciobot-table-row-planning.json`; this pure planner changes no shared state. R31 records the actual native mutation scenario and changeset. D16 retains rich new cell fragments and explicit property overrides; D15 and R08 depend on D16. All 196 application source hashes remain equal after adding those required mappings. A01/R08 retain guarded execution and complete review lifecycle. R29 remains unchecked. See [implementation notes](evidence/implementation-notes.md).

## D09

- [x] Plan table column insertion and deletion

**Depends:** [D07](04-edit-planner.md#d07)

**Primary files:** `fork:packages/sciobot/src/diff/table-columns.ts`; `sibling table-columns.test.ts`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Plan column changes with grid widths and affected spans. Preserve surviving cell content.

**Close when:** Insert and delete a middle column in a three-column fixture. Surviving content and total grid consistency match explicit expected results.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local branch `codex/react-block-proof`. `planTableColumn` validates canonical anchors through the public immutable native dispatcher. It emits existing native insertion/deletion commands without preview IDs or reconstructed subtrees. The explicit three-column fixture proves middle insertion/deletion, inherited widths, totals 7200/9600/4800, unchanged surviving cell/grid IDs, row/table properties, unknown XML, unrelated tables, and insert/delete fingerprint restoration. Merged fixtures widen/narrow horizontal spans, retain vertical merges and all surviving text, and preserve paragraph IDs. Implicit regular grids, nested tables, resource limits, malformed targets, locks, bindings, detached codec roundtrips, and live-state preservation pass. Native save/reopen retains canonical fingerprints and grid consistency. From the repository root, `bun test --preload ./scripts/test/preload-dom.ts packages/sciobot/src/diff/table-columns.test.ts packages/sciobot/src/diff/table-rows.test.ts packages/sciobot/src/diff/inserted-table-rows.test.ts packages/sciobot/src/diff/plan.test.ts` passes 50 cases and 471 assertions. D09 adds 15 cases and 182 assertions. Full private suite: 541 pass, the same two R29 failures, 543 tests, 51 files, 7200 assertions. Private typecheck/build, scoped formatting/lint, root lint, compatibility, and diff checks pass. Build: 70 modules, one external core peer, React-free main entry. Compatibility: `.collaboration/changes/sciobot-table-column-planning.json`, 26 decisions. Pure planning changes no live/shared state. D17 retains rich inserted-column content and explicit new widths; D15/R09 depend on it. A01/R09 retain guarded direct/review execution and two-session verification. Application baseline `be76538c`; all 196 source records match the prior inventory. See [implementation notes](evidence/implementation-notes.md).

## D10

- [x] Plan table merge and split changes

**Depends:** [D07](04-edit-planner.md#d07), [D09](04-edit-planner.md#d09), [N15](07-native-content.md#n15), [N16](07-native-content.md#n16)

**Primary files:** `fork:packages/sciobot/src/diff/table-merges.ts`; `sibling table-merges.test.ts`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25), [S28](sources.md#s28)

**Change:** Plan merged cell boundaries and split restoration. Reject invalid span geometry before creating operations.

**Close when:** Horizontal and vertical merge fixtures retain all intended text. An overlapping span fails without a partial plan.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local branch `codex/react-block-proof`. `planTableMerge` emits existing native merge/split commands with original anchors and requested parameters. Native previews own rectangle geometry, visual chains, grid cuts, existing-grid restoration, locks, and limits. Fifteen pure cases prove horizontal/vertical/rectangle content retention, stable paragraph/control IDs, original unequal grid restoration, native equal-width behavior, inserted rows, blank-cell semantics, payload-part retention, foreign XML, nested ownership, overlap refusal without partial plans, and save/reopen fingerprints. Plans detach/freeze caller data and contain no preview IDs or copied native subtrees. Planning preserves live bytes, revision, selection, and history. N15 supplies existing-grid splitting; N16 fixes the moved-address defect exposed by the new vertical oracle. From the repository root, `bun test --preload ./scripts/test/preload-dom.ts packages/sciobot/src/diff/table-merges.test.ts packages/sciobot/src/diff/table-merge-index-collaboration.test.ts packages/sciobot/src/diff/table-grid-split-collaboration.test.ts packages/sciobot/src/diff/plan.test.ts` passes 30 cases and 5216 assertions. N16 supplies both focused real-session cases. Full private suite: 574 pass, the same two R29 failures, 576 tests, 55 files, 12517 assertions. Core/private typecheck/build, scoped formatting/lint, root lint, scoped core API, compatibility, and diff checks pass. Build: 76 modules, one external core, React-free main entry. Compatibility: `.collaboration/changes/sciobot-table-merge-planning.json`; this pure planner changes no shared state. Native fixes have separate changesets/decisions. D15 retains target/content composition; A01/R10 retain guarded direct/review execution. All 196 audited application source records remain equal at `be76538c`. See [implementation notes](evidence/implementation-notes.md).

## D11

- [x] Plan table and cell property changes

**Depends:** [D07](04-edit-planner.md#d07), [N17](07-native-content.md#n17)

**Primary files:** `fork:packages/sciobot/src/diff/table-format.ts`; `sibling table-format.test.ts`

**Read:** [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Plan widths, borders, backgrounds, and header properties independently from table content.

**Close when:** A cell background change retains its text and neighboring borders. A width change retains merge boundaries.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local branch `codex/react-block-proof`. `planTableFormat` emits one independent native table/cell property intent after immutable preview validation. Core owns units, colors, borders, physical selections, merge geometry, headers, and default protection. Accepted authored no-ops emit no operations; unchanged native drag proposals retain native refusal semantics, while full-grid target assignment is idempotent. Eighteen pure cases verify isolated backgrounds, shading patterns/themes, symmetric border edges, styles/repeated headers, unequal merged widths, exact row heights, imported XML/IDs, complete vertical selections, no-op/refusal boundaries, nested ownership, bounded plain data, detachment, and save/reopen. Source/live bytes, revision, selection, and existing history stay unchanged. From the repository root, `bun test --preload ./scripts/test/preload-dom.ts packages/sciobot/src/diff/table-format.test.ts packages/sciobot/src/diff/table-format-collaboration.test.ts packages/sciobot/src/diff/plan.test.ts` passes 31 cases and 1745 assertions. N17 separately verifies the required native fidelity fix through both real-session delivery orders. Full private suite: 594 pass, the same two R29 failures, 596 tests, 57 files, 14169 assertions. The final focused run adds twelve geometry assertions after that full run. Core/private typecheck/build, scoped formatting/lint, root lint, scoped core API, compatibility, and diff checks pass. Private build: 77 modules, one external core, React-free main. Compatibility: `.collaboration/changes/sciobot-table-property-planning.json`; pure planning changes no live/shared state. D15 retains complete-target composition and Plate header mapping; A01/R11 retain guarded direct/review execution. N14 retains UI controls. All 196 audited application source hashes still match `be76538c`. Application HEAD advanced externally to `9693712d`; a verified revision-only inventory refresh retains every source and mapping record. See [implementation notes](evidence/implementation-notes.md).

## D12

- [x] Pair domain containers without replacing parents

**Depends:** [C08](01-contracts.md#c08), [D05](04-edit-planner.md#d05), [P04](03-agent-projection.md#p04)

**Primary files:** `fork:packages/sciobot/src/diff/domain-containers.ts`; `sibling domain-containers.test.ts`

**Read:** [S09](sources.md#s09), [S10](sources.md#s10), [S18](sources.md#s18)

**Change:** Apply registered identity rules during recursive pairing. Diff children when wrapper identity matches.

**Close when:** Changing one nested question preserves its phase, activity, and source excerpt wrappers. A phaseType change produces a structural replacement.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `planDomainContainers` validates canonical child windows and recursively applies application identity rules. A nested question edit retains phase/activity/excerpt control IDs, content parents, labels, untouched paragraph objects, and other package parts. A phaseType change emits one structural replacement. Citation identity changes replace only the excerpt within retained outer containers. Atomic payloads never recurse into Word fallback. Changed content/property/atomic pairs remain explicit work for D13-D15; no document writes occur. `fixtures/domain-pairing.json` captures 20 decisions by executing the actual application strategies and installed Plate diff engine at application revision `da9081ad`, with source hashes. Regenerate with `bun scripts/sciobot-port/capture-domain-pairing-reference.mjs`. The focused command `bun test --preload ../../scripts/test/preload-dom.ts src/diff/domain-containers.test.ts src/diff/domain-strategies.test.ts src/diff/sibling-blocks.test.ts` from `packages/sciobot` passes 51 cases and 988 assertions. Cases include native nested edit/insertion oracles, source property quirks, inert attributes, exact/unmapped subtrees, opaque IDs, immutable outputs, invalid mappings, getters, cycles, resource limits, and live document/selection/history preservation. The full private suite reports 468 pass and the same two R29 failures across 470 tests, 45 files, and 6,440 assertions. Private typecheck/build, scoped Prettier/oxlint, root lint, diff checks, and compatibility policy pass. The build verifies 63 private source modules with one external core. Compatibility record: `.collaboration/changes/sciobot-domain-pairing.json`; no shared formats changed. This pure planner needs no new collaboration mutation scenario. R29 remains open. See [implementation notes](evidence/implementation-notes.md).

## D13

- [x] Plan granular application property updates

**Depends:** [D12](04-edit-planner.md#d12)

**Primary files:** `fork:packages/sciobot/src/diff/domain-properties.ts`; `sibling domain-properties.test.ts`

**Read:** [S04](sources.md#s04), [S09](sources.md#s09)

**Change:** Plan activity name and duration updates with old-value presence. Separate metadata properties from child text. Preserve all requested serialized attributes on a retained pair, including resource when another update keeps the activity paired.

**Close when:** Renaming an activity produces one property update. Changing duration and one sentence does not replace the activity wrapper.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `planDomainProperties` resolves canonical block/control/part IDs and emits one payload delta for an activity rename. Every delta records exact old-value presence, including absence, null, and zero. Separate duration and sentence plans retain the activity wrapper and native child IDs through an immutable native save/reopen oracle. Registered serialization rules determine property ownership. All requested serialized values on retained pairs receive deltas, including resource beside a name/duration update. Untargeted metadata and unknown payload fields remain intact. Identity changes remain structural; atomic blocks and derived/internal properties use their own leaves. `fixtures/domain-properties.json` captures eight actual application Working/Committed cases at revision `da9081ad`, with the source hash. Regenerate with `bun scripts/sciobot-port/capture-domain-properties-reference.mjs`. From `packages/sciobot`, `bun test --preload ../../scripts/test/preload-dom.ts src/diff/domain-properties.test.ts src/diff/domain-containers.test.ts src/diff/plan.test.ts` passes 59 cases and 284 assertions, including 16 new property cases. Tests cover native codec accept/revert oracles, canonical names, preserved hidden data, semantic no-ops, detached plans, inert nested keys, invalid patches, getters, cycles, limits, and live selection/revision/history preservation. The full private suite reports 484 pass and the same two R29 failures across 486 tests, 46 files, and 6,554 assertions. Private typecheck/build, scoped Prettier/oxlint, root lint, diff checks, and compatibility policy pass. The build verifies 64 private source modules with one external core. Compatibility record: `.collaboration/changes/sciobot-domain-properties.json`; no shared formats changed. This pure planner requires no new collaboration mutation scenario. D15 retains target reconciliation; A01/R12 retain actual revision/old-value checks, fallback updates, transactions, and property review. R29 remains open. See [implementation notes](evidence/implementation-notes.md).

## D14

- [x] Plan atomic domain block replacements

**Depends:** [C08](01-contracts.md#c08), [D12](04-edit-planner.md#d12)

**Primary files:** `fork:packages/sciobot/src/diff/atomic-blocks.ts`; `sibling atomic-blocks.test.ts`

**Read:** [S09](sources.md#s09), [S10](sources.md#s10), [S12](sources.md#s12)

**Change:** Treat visuals, listening exercises, and source chips as atomic changes. Reuse unchanged assets and opaque payloads.

**Close when:** A changed visual produces one replacement. Editing nearby text preserves listening audio references and pristine image identifiers.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `planAtomicBlock` resolves canonical part/control IDs and returns one complete root payload replacement for a changed visual, listening exercise, source chip, or opaque MDX block. Empty path `[]` distinguishes whole-block replacement from granular property edits. Native block ID/type remain stable; type changes require structural planning. Equal targets emit no operations and retain original payloads, parser locations, assets, and opaque metadata. Complete targets come from canonical data and application codecs; D15/X09 retain target decoding and listening reconciliation. The source-executed `fixtures/atomic-assets.json` captures five actual application reconciliation/atomic-diff cases at revision `da9081ad`, with 14 hashes. Regenerate with `bun scripts/sciobot-port/capture-atomic-reference.mjs`. From `packages/sciobot`, `bun test --preload ../../scripts/test/preload-dom.ts src/diff/atomic-blocks.test.ts src/diff/sibling-blocks.test.ts src/diff/domain-containers.test.ts src/diff/plan.test.ts` passes 72 cases and 1,157 assertions, including 16 new atomic cases. Cases cover native codec apply/revert and save/reopen, pristine image ID/source changes, inline citation fields, unhydrated and hydrated audio, script invalidation, complete opaque ASTs, inert expressions/spreads, detached outputs, source identity, getters, syntax, cycles, and limits. A nearby native sentence edit retains both atomic controls, unchanged payloads, audio/portrait references, pristine image IDs, and all other package parts. Shared comparison now ignores only actual AST locations and preserves typed opaque position data. Planning preserves live bytes, revision, selection, and existing history. The full private suite reports 500 pass and the same two R29 failures across 502 tests, 47 files, and 6,666 assertions. Private typecheck/build, scoped Prettier/oxlint, root lint, diff checks, and compatibility policy pass. The build verifies 66 private source modules with one external core. Compatibility record: `.collaboration/changes/sciobot-atomic-block-planning.json`; no shared formats changed. This pure planner needs no new collaboration mutation scenario. A01/R13 retain actual native fallback and review execution; X08/X10 retain asset services. R29 remains open. See [implementation notes](evidence/implementation-notes.md).

## D15

- [x] Plan complete target documents without whole-tree replacement

**Depends:** [D02](04-edit-planner.md#d02), [D03](04-edit-planner.md#d03), [D04](04-edit-planner.md#d04), [D05](04-edit-planner.md#d05), [D06](04-edit-planner.md#d06), [D07](04-edit-planner.md#d07), [D08](04-edit-planner.md#d08), [D09](04-edit-planner.md#d09), [D10](04-edit-planner.md#d10), [D11](04-edit-planner.md#d11), [D12](04-edit-planner.md#d12), [D13](04-edit-planner.md#d13), [D14](04-edit-planner.md#d14), [D16](04-edit-planner.md#d16), [D17](04-edit-planner.md#d17), [P12](03-agent-projection.md#p12), [P21](03-agent-projection.md#p21), [D18](04-edit-planner.md#d18), [D19](04-edit-planner.md#d19), [D20](04-edit-planner.md#d20), [D21](04-edit-planner.md#d21), [D22](04-edit-planner.md#d22), [D23](04-edit-planner.md#d23), [D24](04-edit-planner.md#d24), [D25](04-edit-planner.md#d25)

**Primary files:** `fork:packages/sciobot/src/diff/plan-target.ts`; `fork:packages/sciobot/src/diff/target-inline.ts`; `sibling plan-target.test.ts`; `fork:packages/sciobot/test/support/target-document.ts`; `fork:packages/sciobot/src/diff/target-lists.ts`; `sibling target-lists.test.ts`; `sibling target-structure.test.ts`; `fork:packages/sciobot/src/diff/inserted-list-types.ts`; `fork:packages/sciobot/test/support/target-plan-oracle.ts`

**Read:** [S03](sources.md#s03), [S18](sources.md#s18), [S25](sources.md#s25)

**Change:** Compose the existing planners for set_document. Reuse every unchanged canonical subtree and preserve non-body parts. Carry list instance and checklist intents into inserted fragments without renumbering existing items. Treat P21 empty-task placeholders as generated syntax.

**Close when:** All seven C01 pairs produce explicit operations. Unchanged imported headers, footers, media, and section settings retain their fingerprints.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, then claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. D18-D22 supply parsed tokens, native ownership, logical streams, normalization coordinates, and physical bindings. D15 composes retained paragraphs/containers/source atoms, anchored sibling gaps, and explicit inserted list/checklist intentions. D23 now adds retained native table cell streams, original row deletion commands, rich new rows, and the seventh captured Markdown pair's native planning case. Unique rows anchor matching; repeated rows retain positional identities. Existing numbering instances remain authoritative. Immutable native oracles verify table/link edits, formatting/list insertion, section insertion, deletion, and complete rewrite through save/reopen. Table and structural fixtures preserve imported non-body parts and section settings. The Plate table fixture stores GFM text in a paragraph; the native adaptation uses a Word table. This does not prove identical storage trees or complete parity. Merge dispatch, new property destinations, tables without complete declared grid anchors, richer insertion marks, broader inline/model ownership, continuation/mixed list ownership, new-row lists, composite/barrier streams, empty cell/container seeds, application codecs, and complete target/application parity remain required. A01 and the R leaves retain guarded live execution and two-session lifecycle evidence. D23 owns retained cell/row composition. D24 adds declared-grid column composition and combined row/column targets, including rich new cells and merged-span handling. D25 adds explicit retained property composition and source-proven generated cell separator handling. Merged owner promotion still requires native source reconciliation. See [implementation notes](evidence/implementation-notes.md). Closure (claude): `packages/sciobot/src/diff/target-c01-matrix.test.ts` plans all seven captured C01 pairs against one imported document with a header, footer, footnote, image part, background, unknown XML/binary parts, and a final section; the new `protectedTargetInput` helper in `test/support/target-document.ts` embeds each C01 body there. Each case first plans the captured before Markdown to an empty plan, proving the native source is the captured before state. Explicit expected operation kinds: formatting/lists `text,siblings`; tables/links `text,insert-row,format`; section insertion `siblings`; nested MDX `payload,text,text`; source content `payload,text×4`; deletion `siblings`; complete rewrite `text×5,siblings`. No plan contains `replaceStoryBlocks`. After the immutable native oracle and save/reopen, header, footer, footnotes, media, and unknown parts keep exact bytes; only `document.xml`, the hyperlink relationship part (tables/links), or the payload custom XML item (lesson cases) change. Final section and background keep node identity and fingerprints; retained pair IDs survive; replanning the after Markdown against the reopened package yields no operations; native Markdown ASTs equal the targets (list looseness ignored). Focused run: 8 matrix cases, 222 assertions; the 82 existing composer/structure/table/list/property cases still pass. The remaining set_document capability gaps listed above are not C01 requirements; they are separate required leaves D26-D31, and I05 now depends on them. Direct/review executors (A01, R05) still own live execution and two-session evidence.

## D16

- [x] Plan rich inserted table rows and explicit row properties

**Depends:** [D08](04-edit-planner.md#d08), [P06](03-agent-projection.md#p06), [D03](04-edit-planner.md#d03), [D04](04-edit-planner.md#d04)

**Primary files:** `fork:packages/sciobot/src/diff/inserted-table-rows.ts`; `sibling inserted-table-rows.test.ts`; `fork:packages/sciobot/src/diff/plan.ts`

**Read:** [S03](sources.md#s03), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Extend row intents to carry new cell AST fragments and explicit row properties. Preserve inline marks, links, manual breaks, multiple paragraphs, nested blocks, and owned application payloads. Resolve new native row/cell/paragraph IDs during execution rather than storing IDs from a planning preview. Reuse Apache native row creation and targeted content operations. Add a small prerequisite if a required native operation is absent.

**Close when:** A new row with bold text, a link, a manual break, multiple paragraphs, and an application block retains those values and explicit row properties through a native materialization oracle and save/reopen. Existing row/cell IDs and unrelated tables remain unchanged. Plan serialization contains only new fragments and canonical anchors.

**Check:** Add focused pure plan tests and native materialization evidence. Actual document mutation code also needs the focused two-session scenario from the shared rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local branch `codex/react-block-proof`. `planInsertedTableRow` emits immutable `insert-row` intents with existing anchors, rich new cell AST streams, exact height, and complete owned payload bindings. Native previews validate physical cell counts, inherited merges, locks, height, and storage. New native IDs remain outside plans. The test-only immutable oracle resolves IDs after native row creation and verifies bold text, a referenced external link, a manual break, multiple paragraphs, a nested table, application/opaque React data, exact row height, existing row/cell identities, unrelated-table fingerprints, and save/reopen. Source/live bytes, revision, selection, and history remain unchanged during planning. Focused command: `bun test --preload ./scripts/test/preload-dom.ts packages/sciobot/src/diff/inserted-table-rows.test.ts packages/sciobot/src/diff/plan.test.ts packages/sciobot/src/diff/table-rows.test.ts`; 35 pass, 289 assertions. D16 adds 11 cases and 127 assertions. Full private suite: 526 pass, the same two R29 failures, 528 tests, 50 files, 7018 assertions. Private typecheck/build, scoped formatting/lint, root lint, compatibility, and diff checks pass. Build: 69 modules, one external core peer, React-free main entry. Compatibility: `.collaboration/changes/sciobot-rich-row-planning.json`, 25 decisions. Pure planning adds no live/shared state or published API. D15 retains complete-target composition; A01/R08 retain guarded direct/review execution and rich-row two-session verification. Application baseline `be76538c`; all 196 audited source hashes match. See [implementation notes](evidence/implementation-notes.md).

## D17

- [x] Plan rich inserted table columns and explicit new widths

**Depends:** [D09](04-edit-planner.md#d09), [D16](04-edit-planner.md#d16), [P06](03-agent-projection.md#p06)

**Primary files:** `fork:packages/sciobot/src/diff/inserted-table-columns.ts`; `sibling inserted-table-columns.test.ts`; `fork:packages/sciobot/src/diff/plan.ts`

**Read:** [S03](sources.md#s03), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Carry row-anchored AST streams and owned application payloads for newly created column cells. Include an explicit new grid width when requested. Resolve new grid/cell/paragraph IDs after native column creation. Reuse native span geometry: an insertion inside a spanning cell widens that existing cell instead of creating another cell. Preserve its existing content. Require empty streams for hidden vertical continuations. Reuse D16 fragment and payload rules. Add a small prerequisite if a required native operation is absent.

**Close when:** A new middle column retains bold text, a link, a manual break, multiple paragraphs, and an application block. Its explicit width and content survive a native materialization oracle and save/reopen. Merged neighbors retain their text and valid spans. Existing row/cell IDs and unrelated tables remain unchanged. Serialization contains only new fragments, payloads, requested properties, and canonical anchors. Native allocation changes between preview and materialization cannot redirect content.

**Check:** Add focused pure plan tests and native materialization evidence. Actual document mutation code also needs the focused two-session scenario from the shared rules.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local branch `codex/react-block-proof`. `planInsertedTableColumn` emits immutable `insert-column` intents with original anchors, row-addressed rich streams, optional explicit widths, and complete owned payload bindings. Native previews resolve new cells versus widened spans and hidden continuations. Serialized plans contain no allocated native IDs. The native immutable oracle verifies both insertion sides, bold text, a referenced hyperlink, a manual break, multiple paragraphs, application/opaque React data, explicit width, merged content, stable existing row/cell IDs, foreign XML, unrelated tables, and save/reopen fingerprints. Runtime allocation changes cannot redirect content or width. Shared D16 validation retains all 127 D16 assertions. Planning preserves live bytes, revision, selection, and undo/redo history. From the repository root, `bun test --preload ./scripts/test/preload-dom.ts packages/sciobot/src/diff/inserted-table-columns.test.ts packages/sciobot/src/diff/inserted-table-rows.test.ts packages/sciobot/src/diff/table-columns.test.ts packages/sciobot/src/diff/table-rows.test.ts packages/sciobot/src/diff/plan.test.ts` passes 64 cases and 653 assertions. D17 adds 14 cases and 182 assertions. Full private suite: 555 pass, the same two R29 failures, 557 tests, 52 files, 7382 assertions. Private typecheck/build, scoped formatting/lint, root lint, compatibility, and diff checks pass. Build: 75 modules, one external core peer, React-free main entry. Compatibility: `.collaboration/changes/sciobot-rich-column-planning.json`, 27 decisions. Explicit sizing follows native fixed-grid limits and preferred-width semantics. Pure planning adds no live/shared state. D15 retains full-target composition; A01/R09 retain guarded direct/review execution and rich-column two-session verification. All 196 audited application source hashes match baseline `be76538c`. See [implementation notes](evidence/implementation-notes.md).

## D18

- [x] Compile inline target data without flattening opaque atoms

**Depends:** [P06](03-agent-projection.md#p06), [D02](04-edit-planner.md#d02), [D03](04-edit-planner.md#d03)

**Primary files:** `fork:packages/sciobot/src/diff/inline-target.ts`; `sibling inline-target.test.ts`

**Read:** [S03](sources.md#s03), [S18](sources.md#s18), [S25](sources.md#s25)

**Change:** Compile one parsed paragraph, heading, or GFM table cell into detached text/mark tokens and explicit hyperlink owners. Keep images, mentions, math, references, unknown inline types, and MDX expressions as complete inert atoms. Retain the full input AST and original child paths. Use native hyperlink authorability rules. Keep native address mapping and mutations outside this compiler.

**Close when:** Explicit fixtures prove nested Markdown/native HTML marks, literal code, manual breaks, Unicode text, link identity, tooltips, bookmark targets, and opaque atom retention. C01 inline content compiles with explicit expected text/mark results. Unknown metadata stays available. Empty strings and authored NBSP remain exact; D15 alone identifies generated checklist placeholders from mapped source metadata. Invalid/plain-data/resource inputs refuse before evaluation. Compilation preserves live source bytes, revision, selection, and history.

**Check:** Run focused Bun compiler/parser/property tests, private typecheck/build, scoped formatting/lint, root lint, compatibility policy, and the backlog check. This pure compiler needs no collaboration mutation scenario.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local branch `codex/react-block-proof`. `compileInlineTarget` returns detached immutable text/mark tokens, distinct hyperlink owners/targets/titles, complete inert atoms, and the full source AST. Paths address one parsed target and contain no native IDs. Opaque MDX, expressions, images, references, math, mentions, unknown inline data, and metadata retain their complete subtrees; their fallback never becomes text. Authored NBSP remains exact. Native hyperlink authorability validates external targets. Invalid model/XML text, nested links, getters, coercion-shaped type objects, cycles, hidden fields, sparse arrays, and byte/node/depth excess refuse before evaluation or partial output. Twenty compiler cases include explicit inline expectations for all seven C01 pairs and preserve live bytes, revision, selection, imported XML, and history. From the repository root, `bun test --preload ./scripts/test/preload-dom.ts packages/sciobot/src/diff/inline-target.test.ts packages/sciobot/src/projection/parse-target.test.ts packages/sciobot/src/diff/run-format.test.ts` passes 63 cases and 502 assertions. Full private suite: 614 pass, the same two R29 failures, 616 tests, 58 files, 14393 assertions. Private typecheck/build, scoped formatting/lint, root lint, compatibility, and diff checks pass. Build: 78 modules, one external core peer, React-free main. Compatibility: `.collaboration/changes/sciobot-inline-target-compilation.json`, 33 decisions; no native/shared formats change. This pure compiler needs no collaboration mutation test or published changeset. D15 retains native ownership/offset mapping, generated task handling, complete-target composition, all seven explicit operation cases, and non-body fidelity. All 196 audited source hashes and application revision `9693712d` match the baseline. See [implementation notes](evidence/implementation-notes.md).

## D19

- [x] Index native ownership of projected source constructs

**Depends:** [P04](03-agent-projection.md#p04), [P05](03-agent-projection.md#p05), [P12](03-agent-projection.md#p12), [P21](03-agent-projection.md#p21), [D05](04-edit-planner.md#d05)

**Primary files:** `fork:packages/sciobot/src/diff/target-source.ts`; `fork:packages/sciobot/src/diff/target-source-types.ts`; `sibling target-source.test.ts`; `fork:packages/sciobot/src/projection/blocks.ts`; `fork:packages/sciobot/src/projection/markdown.ts`; `fork:packages/sciobot/src/data/json.ts`

**Read:** [S03](sources.md#s03), [S18](sources.md#s18), [S25](sources.md#s25)

**Change:** Index a current MDX projection against its canonical body. Return original paragraph/table/control owners, direct parent positions, nearest table scopes, exact source extents, checklist bindings, and complete opaque AST records. Preserve each block's actual registry kind. Build contiguous sibling windows without crossing labels, atomic fallback, native metadata, opaque XML, or unprojected owners. Keep normalization, target pairing, planning, and mutations outside this read-only index.

**Close when:** Explicit fixtures prove nested containers, inline atoms, registered custom kinds, merged/nested tables, generated labels, empty tasks, source ranges, protected gaps, and zero-length insertion windows. Every window names actual consecutive native children. Existing sibling planners consume a returned window with stable anchors. Stale, conflicting, invalid, or excessive source metadata refuses without a partial index. Headers, footers, notes, media, section settings, source bytes/revision/selection/history, and caller data remain unchanged.

**Check:** Run focused native ownership, source projection, sibling planning, inline compilation, and protection tests. Run private typecheck/build, scoped formatting/lint, root lint, compatibility policy, source coverage, and the backlog check. This immutable reader needs no collaboration mutation scenario.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local branch `codex/react-block-proof`. `indexTargetSource` returns detached immutable original body owners, actual parent/child positions, nearest cell scopes, full extents/source slices, checklist bindings, complete opaque AST records, and consecutive native sibling windows. Registry kinds remain explicit, including empty containers, custom types, and atomic host overrides. Labels, opaque XML, property holders, ordinary Word wrappers, unprojected owners, inline SDT content, and atomic fallback retain their boundaries. Hidden continuations get no invented text. Empty windows retain physical anchors before a final section property; geometry and writability remain executor checks. Valid native fallback slices now omit only an undefined optional boundary field in the private JSON DTO while retaining precision flags. Eighteen cases verify mapping, a returned-window sibling operation, three imported review views, stale/conflicting metadata refusal, resource/identity limits, detachment, and unchanged live bytes/revision/selection/history. From the repository root, `bun test --preload ./scripts/test/preload-dom.ts packages/sciobot/src/diff/target-source.test.ts packages/sciobot/src/projection/blocks.test.ts packages/sciobot/src/projection/markdown.test.ts packages/sciobot/src/diff/inline-target.test.ts packages/sciobot/src/diff/sibling-blocks.test.ts packages/sciobot/src/projection/protected-content.test.ts` passes 75 cases and 1686 assertions. Full private suite: 632 pass, the same two R29 failures, 634 tests, 59 files, 14669 assertions. Private typecheck/build, scoped formatting/lint, root lint, compatibility, source coverage, and diff checks pass. Build: 79 modules, one external core, React-free main. Compatibility: `.collaboration/changes/sciobot-target-source-index.json`, 34 decisions; no native/shared format changes. This immutable reader needs no new collaboration mutation test or published changeset. D15 retains AST binding, full planner composition, actual C01 operation/application evidence, insertion list/checklist intents, and generated task handling. All 196 audited source hashes and application revision `9693712d` match the baseline. See [implementation notes](evidence/implementation-notes.md).

## D20

- [x] Normalize parsed targets into logical sibling block streams

**Depends:** [P06](03-agent-projection.md#p06), [D06](04-edit-planner.md#d06), [D18](04-edit-planner.md#d18), [D19](04-edit-planner.md#d19)

**Primary files:** `fork:packages/sciobot/src/diff/logical-blocks.ts`; `fork:packages/sciobot/src/diff/logical-block-types.ts`; `sibling logical-blocks.test.ts`; `fork:scripts/sciobot-port/capture-logical-reference.mjs`; `fork:packages/sciobot/src/__tests__/fixtures/logical-blocks.json`

**Read:** [S03](sources.md#s03), [S18](sources.md#s18), [S25](sources.md#s25)

**Change:** Normalize a parsed Markdown target into paragraph, container, table, and atomic logical blocks. Flatten list wrappers while retaining item/continuation roles, nesting, task state, group paths, and syntax starts. Normalize registered inline/flow container forms and native HTML table structure. Follow captured recap list rules. Keep the entire original AST, opaque atoms, unknown attributes, and original child paths. Produce matching hints without native IDs, writes, or list renumbering.

**Close when:** Explicit expectations cover all seven C01 before/after streams, nested/multiparagraph lists, empty tasks, native HTML tables, inline domain containers, recap rules, custom registry overrides, opaque subtrees, and unknown metadata. Independent executed application fixtures support recap/source normalization. Both sides use the same normalizer. Matching keys alone never authorize a no-op; list and property intentions remain explicit. Refusals preserve input data and existing native bytes/revision/selection/history.

**Check:** Run focused normalizer/parser/inline/source ownership tests, private typecheck/build, scoped formatting/lint, root lint, compatibility policy, source coverage, and backlog checks. This pure target transform needs no collaboration mutation scenario.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local branch `codex/react-block-proof`. `normalizeLogicalBlocks` returns detached immutable paragraph/container/table/atomic sibling streams while retaining the full original AST. Lists keep input group/item paths, nesting, item/continuation roles, task state, and syntax starts without native IDs or numbering writes. Registered inline/flow containers, native HTML table rows/cells/groups, nested cells, paragraph wrappers, and host overrides retain explicit structure and attributes. Dynamic tables and opaque React blocks remain complete inert atoms. Twelve independently executed application fixtures support recap suppression, legacy prefix handling, title normalization, and empty-source wrapping. The captured continuation case records Plate's collapsed list metadata separately; native streams retain parsed nesting. Matching keys never authorize no-op or write decisions. Twenty-one normalizer cases include explicit streams for all seven C01 pairs, metadata, marks, links, lists, empty tasks, native HTML, opaque data, bounded refusals, detachment, and unchanged native bytes/revision/selection/history. From the repository root, focused normalizer/parser/inline/source ownership tests pass 92 cases and 990 assertions. Full private suite: 653 pass, the same two R29 failures, 655 tests, 60 files, 14945 assertions. Private typecheck/build, scoped formatting/lint, root lint, compatibility, source coverage, backlog, and diff checks pass. Build: 80 modules, one external core, React-free main. Compatibility: `.collaboration/changes/sciobot-logical-block-normalization.json`, 35 decisions; native/shared formats remain unchanged. This immutable transform needs no collaboration mutation scenario or published changeset. D15 retains logical/native binding, full composition, actual C01 operation/application evidence, insertion intents, and generated task handling. All 196 source hashes and application revision `9693712d` match the baseline. See [implementation notes](evidence/implementation-notes.md).

## D21

- [x] Retain exact source offsets through target string normalization

**Depends:** [P06](03-agent-projection.md#p06), [D19](04-edit-planner.md#d19), [D20](04-edit-planner.md#d20)

**Primary files:** `fork:packages/sciobot/src/projection/target-source-map.ts`; `fork:packages/sciobot/src/projection/normalize-target.ts`; `fork:packages/sciobot/src/projection/parse-target.ts`; `sibling target-source-map.test.ts`

**Read:** [S03](sources.md#s03), [S18](sources.md#s18), [S25](sources.md#s25)

**Change:** Record exact UTF-16 coordinate changes during the existing HTML-break and ordered-list string normalizers. Return that immutable trace with parsed targets. Map normalized AST offsets and ranges back to original projection Markdown. Preserve normalizer output and reference ASTs. Keep rewritten tag interiors unmappable, mark generated separators explicitly, and leave native resolution to existing source maps.

**Close when:** Explicit source/normalized coordinates cover no-op text, length-changing and multiline breaks, inserted list separators, composed changes, code/fence exclusions, Unicode boundaries, and all C01/reference parsing cases. A real native source projection resolves text after a normalized break to its original paragraph/UTF-16 offsets. Generated syntax never becomes authored text ownership. Parsing and mapping preserve live native bytes, revision, selection, and history. No text diff guesses source identity.

**Check:** Run focused source-map/parser/normalizer/source ownership tests, private typecheck/build, scoped formatting/lint, root lint, compatibility policy, source coverage, and backlog checks. This immutable read trace needs no collaboration mutation scenario.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local branch `codex/react-block-proof`. Existing HTML-break and ordered-list normalizers capture their actual changes without text-diff identity guesses. Parsed targets retain an immutable UTF-16 trace. `targetSourceOffsetAt` and `targetSourceRangeAt` return original source coordinates, explicit rewritten/generated status, complete tag edges, and no rewritten interior or surrogate-split coordinates. Generated separators retain zero-length original spans. Seventeen cases cover identity/empty text, length changes, equal-length and multiline tags, multiple/composed edits, code/fence exclusions, existing expression/math handling, CRLF/Unicode, invalid ranges, immutable data, 24 independent parser references, and all seven C01 streams. A real native table resolves text after a rewritten break to its original paragraph and model offsets 2-7; a prior user insertion moves the same source to offset 7. Nested table text retains its own cell and offsets 0-8. Exact table syntax still has no native text owner. Parsing/mapping/refusals preserve live bytes, imported XML, revision, selection, and undo/redo history. Focused source-map/parser/logical/source ownership tests pass 89 cases and 1272 assertions. Full private suite: 670 pass, the same two R29 failures, 672 tests, 61 files, 15439 assertions. Private typecheck/build, scoped formatting/lint, root lint, compatibility, source coverage, backlog, and diff checks pass. Build: 81 modules, one external core, React-free main. Compatibility: `.collaboration/changes/sciobot-target-normalization-map.json`, 36 decisions; native/shared formats remain unchanged. No new collaboration mutation scenario, published changeset, or public API extraction is needed. D15 retains logical/native binding, full composition, actual C01 operations/application, insertion intents, and generated task handling. All 196 source hashes and application revision `9693712d` match the baseline. See [implementation notes](evidence/implementation-notes.md).

## D22

- [x] Bind logical source blocks to native owners and physical sibling windows

**Depends:** [P03](03-agent-projection.md#p03), [D19](04-edit-planner.md#d19), [D20](04-edit-planner.md#d20), [D21](04-edit-planner.md#d21)

**Primary files:** `fork:packages/sciobot/src/diff/bind-source.ts`; `fork:packages/sciobot/src/diff/bound-source-types.ts`; `sibling bind-source.test.ts`; `fork:packages/sciobot/src/projection/tables.ts`; `fork:packages/sciobot/src/projection/mdx.ts`

**Read:** [S03](sources.md#s03), [S18](sources.md#s18), [S25](sources.md#s25)

**Change:** Parse and normalize the current native projection, then bind logical source paths through D21 coordinates to D19 owners. Retain native P03 table/row/cell identities in the composed MDX projection. Identify physical cells, span placeholders, hidden continuations, composite paragraphs, inline block atoms, and generated empty-task syntax without guessing identities from text. Keep original source ASTs and actual native sibling windows. Record unrepresented native owners explicitly.

**Close when:** Repeated text, nested containers/tables, empty cells, merged placeholders, ordinary Word wrappers, labels, opaque atomic fallbacks, custom registry kinds, inline controls, and generated task NBSP have explicit ownership evidence. Composite logical content never gains one invented native paragraph. Window members keep actual canonical child IDs and anchors. A one-to-one returned binding drives an existing sibling plan. Invalid/stale/conflicting metadata refuses without native bytes/revision/selection/history changes. Protection and geometry remain executor checks. Full D15 planning and application requirements remain open.

**Check:** Run focused binding/source/table/MDX/logical/source-map tests, private typecheck/build, scoped formatting/lint, root lint, compatibility policy, source coverage, and backlog checks. This immutable reader needs no collaboration mutation scenario.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local branch `codex/react-block-proof`. `bindLogicalSource` connects logical source paths through exact normalization coordinates to original native owners and physical sibling windows. Composed MDX retains P03 native table/row/cell metadata and final cell slices without a second Markdown export. Physical cells, span placeholders, hidden continuations, missing slots, nearest nested cells, empty paragraphs, labels, and opaque fallbacks remain distinct. Composite or competing paragraph claims receive no invented single owner. Inline atoms retain actual model ranges and complete ASTs; native atomic policy suppresses container-shaped opaque fallbacks. Registry mismatches refuse. Empty native source excerpts retain their canonical wrapper despite parsed inline shape. Generated task NBSP retains exact source proof and insertion offset; authored NBSP stays text. Unrepresented owners and actual window anchors remain explicit. A returned binding drives a real existing sibling plan. Twenty-seven binding cases include native ownership, registry/metadata ambiguity, immutable output, three imported review views, and unchanged live bytes/revision/selection/history. Focused binding/source/table/MDX/logical/source-map tests pass 140 cases and 2228 assertions. Full private suite: 697 pass, the same two R29 failures, 699 tests, 62 files, 15711 assertions. Private typecheck/build, scoped formatting/lint, root lint, compatibility, source coverage, backlog, and diff checks pass. Build: 82 modules, one external core, React-free main. Compatibility: `.collaboration/changes/sciobot-logical-source-binding.json`, 37 decisions; native/shared formats remain unchanged. This immutable reader needs no collaboration mutation scenario, published changeset, or public API extraction. D15 retains full composition, native token/format targeting, actual C01 operations/application, insertion intents, and generated task handling during edits. All 196 source hashes and application revision `9693712d` match the baseline. See [implementation notes](evidence/implementation-notes.md).

## D23

- [x] Compose retained table cells and row edits

**Depends:** [D07](04-edit-planner.md#d07), [D08](04-edit-planner.md#d08), [D16](04-edit-planner.md#d16), [D18](04-edit-planner.md#d18), [D22](04-edit-planner.md#d22)

**Primary files:** `fork:packages/sciobot/src/diff/target-table.ts`; `sibling target-table.test.ts`; `fork:packages/sciobot/src/diff/plan-target.ts`; `fork:packages/sciobot/src/projection/target-mentions.ts`; `fork:packages/sciobot/test/support/target-plan-oracle.ts`; `fork:packages/sciobot/test/support/inserted-fragment-oracle.ts`

**Read:** [S03](sources.md#s03), [S18](sources.md#s18), [S25](sources.md#s25)

**Change:** Compose D22 table ownership with retained cell streams, original row deletion commands, and D16 inserted row fragments. Keep exact source positions for text that needs no mention transformation. Retain table, row, cell, paragraph, and imported property identities. Keep new native IDs deferred. D15 retains column/merge/property dispatch and remaining complete-target ownership.

**Close when:** The C01 table/link case plans the existing cell text, one inserted row, and the existing hyperlink target. Multiple row insertions preserve target order at every edge. Repeated rows retain their native positions. Visible merged cells and nested tables retain their nearest native owners. Rich row marks/links, row deletion, source bytes, imported non-body parts, and section settings have native save/reopen evidence. Unsupported metadata and new-row list destinations return no partial plan. These refusal checks do not establish support for those D15 requirements.

**Check:** Run focused table/composer/binding/rich-fragment/parser tests, the private suite, private typecheck/build, scoped formatting/lint, root lint, compatibility, source coverage, and backlog checks. Pure composition needs no new collaboration mutation scenario. A01 and the review leaves retain guarded live execution and two-session evidence.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local branch `codex/react-block-proof`. `planTargetTable` retains D22 native table/row/cell bindings and composes existing paragraph/cell, row deletion, and rich inserted-row intentions. Unique rows anchor alignment; repeated rows pair by position. Retained merged slots and nested tables keep exact owners. New row fragments retain original nested table ASTs and deferred native IDs. Combined immutable previews validate topology. The C01 table/link adaptation produces cell text, one inserted row, and an existing hyperlink target change. Thirteen table cases verify edge/middle row order, deletion, marks/links, nested fragments, repeated values, merges, native property identity, source bytes, imported non-body parts, section settings, and save/reopen. A parser regression preserves untouched text positions for GFM binding and escaped text. The test-only fragment oracle removes only the generated host tail after a final authored paragraph and declares all new nested grid columns. Focused tests pass 115 cases and 974 assertions across six files. Full private suite: 752 pass, the same two R29 failures, 754 tests, 67 files, 16165 assertions. Commands and logs are in the implementation notes. Private typecheck/build, scoped formatting/lint, root lint, compatibility, source coverage, backlog, and diff checks pass. Build: 87 modules, one external core, React-free main. Compatibility: `.collaboration/changes/sciobot-target-table-composition.json`, 40 decisions; this leaf changes no native/shared formats. The coverage inventory adds AIKit/TableKit mappings only; all 196 source hashes and application revision `9693712d` remain unchanged. D15 retains column/merge/property dispatch, new-row list destinations, composite/empty ownership, application codecs, and complete target parity. A01 and the R leaves retain guarded live execution and collaboration lifecycle checks. This pure private component needs no published changeset or public API extraction. See [implementation notes](evidence/implementation-notes.md).

## D24

- [x] Compose target edits over declared native table grid columns

**Depends:** [D09](04-edit-planner.md#d09), [D17](04-edit-planner.md#d17), [D23](04-edit-planner.md#d23)

**Primary files:** `fork:packages/sciobot/src/diff/target-table-columns.ts`; `sibling target-table-columns.test.ts`; `fork:packages/sciobot/src/diff/target-table.ts`; `fork:packages/sciobot/src/diff/target-sequence.ts`; `fork:packages/sciobot/src/diff/target-table-fragments.ts`; `fork:packages/sciobot/test/support/target-table-fixture.ts`; `fork:packages/sciobot/test/support/target-plan-oracle.ts`

**Read:** [S03](sources.md#s03), [S18](sources.md#s18), [S25](sources.md#s25)

**Change:** Compose D09 deletion and D17 rich insertion with original native grid anchors. Map retained cells across column changes and align rows through retained columns. Validate combined row/column topology without copying a canonical table or exposing preview IDs. Keep D15 property/merge dispatch and remaining source ownership requirements open.

**Close when:** Left, middle, right, repeated, and multiple column changes preserve explicit target order and original surviving native IDs. Rich new cells retain marks/links/fragments. Combined row/column changes preserve original row anchors and target content through native application and save/reopen. Existing grid widths, imported property metadata, unrelated parts, and source bytes retain explicit expected evidence. Unsupported destinations refuse atomically. Tables without declared grid anchors remain an explicit D15 requirement.

**Check:** Run focused target table/column/row/planner tests, the private suite, private typecheck/build, scoped formatting/lint, root lint, compatibility, source coverage, and backlog checks. Pure planning needs no new collaboration mutation scenario. A01 and the review leaves retain actual guarded writes and two-session lifecycle evidence.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local branch `codex/react-block-proof`. `target-table-columns.ts` composes original D09 grid-column deletion and D17 rich insertion. Complete vectors, header text, unique body evidence, and positional fallback retain explicit native cell bindings. Shared sequence matching preserves D23 row behavior. Rows align through retained columns; column previews precede row topology plans. Combined insert/delete cases retain original row/grid anchors and carry no preview identities. Seventeen column cases verify left/middle/right/multiple changes, simultaneous insertion/deletion, row/column composition, repeated headers, header mark changes, explicit marks/links, nested new tables, merged-span growth/restoration, foreign grid metadata, nonuniform widths, source bytes, unrelated parts, background, section settings, and save/reopen. Unsupported paragraph attributes and non-authorable new links refuse atomically after earlier planning. The former column refusal expectation now has positive native application coverage. Focused tests pass 178 cases and 1764 assertions across ten files. Full private suite: 769 pass and the same two R29 failures, 771 tests, 68 files, 16354 assertions. Private typecheck/build, scoped formatting/lint, root lint, compatibility, source coverage, backlog, and diff checks pass. Build: 90 modules, one external core peer, React-free main. Compatibility: `.collaboration/changes/sciobot-target-column-composition.json`, 41 decisions; this leaf changes no native/shared formats. The inventory adds only AIKit/TableKit D24 mappings. All 196 source hashes and application revision `9693712d` stay unchanged. Tables without complete declared grid anchors, property/merge dispatch, and remaining D15 ownership remain required. A01 and review leaves retain guarded live writes and two-session lifecycle evidence. This pure private component needs no published changeset or public API extraction. See [implementation notes](evidence/implementation-notes.md).

## D25

- [x] Compose explicit retained table property intentions

**Depends:** [D04](04-edit-planner.md#d04), [D11](04-edit-planner.md#d11), [D24](04-edit-planner.md#d24)

**Primary files:** `fork:packages/sciobot/src/diff/target-table-properties.ts`; `sibling target-table-properties.test.ts`; `fork:packages/sciobot/src/diff/target-property-values.ts`; `fork:packages/sciobot/src/diff/target-table.ts`; `fork:packages/sciobot/test/support/target-document.ts`

**Read:** [S03](sources.md#s03), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Decode explicit GFM alignment and static HTML table/row/cell property requests into existing native D04/D11 plans. Preserve omitted imported properties, original native ownership, text, and independent formatting. Keep structural matching separate from property intentions. Reject unmapped metadata and dynamic expressions before returning a plan.

**Close when:** Alignment, cell fill/clear, borders, row height, and table sizing have explicit native property outcomes through save/reopen. Combined content/property changes retain original IDs and imported metadata. Whole merged selections retain native validation. Repeated requests become authored no-ops. New destinations, unsupported units/values, conflicts, and unknown properties refuse without partial source writes. D15 retains remaining new-destination and complete-document requirements.

**Check:** Run focused property/table/composer/native-format tests, the private suite, private typecheck/build, scoped formatting/lint, root lint, compatibility, source coverage, and backlog checks. This pure composer needs no new collaboration mutation scenario. A01/R11 retain guarded live execution and two-session lifecycle evidence.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Static property decoding separates native intentions from table structure matching. Existing D04/D11 planners retain original table, row, cell, paragraph, and imported property identities. Eighteen property cases cover GFM alignment, HTML/CSS inheritance, static numeric literals, absolute units, fill/clear, borders, widths, heights, combined text changes, native merged selections, nested tables, generated separators, authored hard breaks, no-ops, and atomic refusals. Native save/reopen preserves requested properties, foreign metadata, unrelated parts, background, and section settings. Source-proven separator removal retains canonical offsets. Changed merge ownership refuses until native source reconciliation exists. Focused tests pass 121 cases and 1808 assertions across eight files. Full private suite: 787 pass and the same two R29 failures, 789 tests, 69 files, 16485 assertions. Private typecheck passes 1626 files. Build passes with 95 private modules, one external core peer, and a React-free main. Scoped formatting/lint, root lint, compatibility, source coverage, backlog, and diff checks pass. Compatibility: `.collaboration/changes/sciobot-target-property-composition.json`, 42 decisions. This pure private composer changes no native/shared formats and needs no published changeset or public API extraction. All 196 application source hashes and revision `9693712d` remain unchanged. D15 retains complete-target support, merge reconciliation, and new property destinations. A01/R11 retain guarded live execution and two-session lifecycle evidence. See [implementation notes](evidence/implementation-notes.md).

## D26

- [x] Plan formatting of inserted text in retained paragraphs

**Depends:** [D15](04-edit-planner.md#d15)

**Primary files:** `fork:packages/sciobot/src/diff/target-inline.ts`; `fork:packages/sciobot/src/diff/insertion-marks.ts`; `sibling insertion-marks.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** When target text inserts or replaces characters whose bold, italic, or strike marks differ from the inherited neighbor run, add explicit run-format intentions for the inserted range. Address post-text coordinates with old values read from a pure native preview of the same plan. Keep inherited fonts, styles, and unrelated properties.

**Close when:** `Tail.` to `**New** Tail.`, `One two.` to `One *big* two.`, and replacements that drop or add marks plan explicit operations. The native oracle lands the exact target Markdown through save/reopen, and unrelated run properties keep their fingerprints.

**Check:** Add a focused Bun test with explicit expected results. Run the C01 matrix and composer regressions.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `packages/sciobot/src/diff/insertion-marks.ts` adds `planMarkDeltas` and `previewTextEdits`. When inserted or replacing text would inherit marks that differ from the target, `planTargetParagraph` applies its own text edits to an immutable native preview and plans D03 run-format intentions in post-text coordinates after the text edits, with old values read from that preview. Same-text mark changes reuse the same helper over the original part. Focused run: `insertion-marks.test.ts`, 8 tests, 55 assertions. The exact plan for `Tail.` to `**New** Tail.` is one insertion and one `b` format over offsets 0-3 with explicit absence before. Bold prefix, italic word mid-sentence, plain replacement of a bold word, bold replacement of a plain word, bold-italic insertion, and plain prefix plus struck suffix each land the exact target AST through the native oracle and save/reopen, and through the live A01 executor in one undo unit; every landed run keeps the inherited unknown property and color. Inline code still refuses (D31). All 429 `src/diff` tests pass, including the updated composer refusal test and the C01 matrix. Limit: tracked executors must accept post-text formatting of their own inserted runs; R06/R07 own that lowering.

## D27

- [x] Plan hyperlink structure changes in retained paragraphs

**Depends:** [D15](04-edit-planner.md#d15), [D26](04-edit-planner.md#d26)

**Primary files:** `fork:packages/sciobot/src/diff/target-inline.ts`; `fork:packages/sciobot/src/diff/target-links.ts`; `sibling target-links.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Plan text edits inside and beside existing native hyperlink wrappers, link insertion around existing text, link removal, and link boundary changes. Retain the original wrapper and relationship when its target survives. Never mint relationships in the plan.

**Close when:** Adding a link to existing words, removing one, changing link text, and editing text after a link each plan explicit operations. The native oracle lands the target Markdown through save/reopen with retained wrapper identities where the link survives.

**Check:** Add a focused Bun test with explicit expected results. Run the C01 matrix and composer regressions.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. D01 gains `link` (exact post-text range, external URL or bookmark target, title; never a relationship ID) and `unlink` (wrapper ID) edits. `packages/sciobot/src/diff/target-links.ts` adds `planLinkStructure`. When text or link spans change, `planTargetParagraph` applies its text edits to a native preview. Wrappers whose exact span survives keep their identity and are retargeted through D03; other wrappers are unlinked with their runs kept; target links over new spans are inserted. A tooltip removal becomes unlink plus link, because native retargeting cannot drop a tooltip. `linkAwareBias` points each pure insertion at a link boundary into or away from the link, following the target, so wrappers survive text edits beside them. Unchanged-text retargeting keeps its original path, so C01 tables/links still plans `text, insert-row, format`. A01 lowers `link` through the gated relationship mint (previewed with a scratch relationship) and `unlink` through native `removeHyperlink`. Focused run: `target-links.test.ts`, 7 tests, 58 assertions. Each case asserts exact operation kinds, codec round-trip, live direct landing of the exact target AST, empty replanning, one-step undo restoring original wrapper IDs, redo, and reopen. Cases: link around existing words (`link`), link removal (`unlink`), changed link text (`text×3, unlink, link`), text after a link (`text`, wrapper retained), text before a retargeted link (`text×3, format`, wrapper retained), tooltip removal (`unlink, link`). Link plans carry no relationship IDs; an unsafe `javascript:` target never commits. 475 diff/ai tests pass; private typecheck has no errors in these files; oxlint passes; compatibility decision `.collaboration/changes/sciobot-target-link-structure.json`. Limit: tracked lowering of `link`/`unlink` belongs to R27.

## D28

- [x] Plan inserted tables and block quotes

**Depends:** [D15](04-edit-planner.md#d15)

**Primary files:** `fork:packages/sciobot/src/diff/plan-target.ts`; `fork:packages/sciobot/src/diff/fragment-seeds.ts`; `sibling fragment-seeds.test.ts`

**Read:** [S03](sources.md#s03), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Accept new GFM tables and block quotes as validated sibling insertion fragments. Add one pure native fragment materializer that direct and suggestion executors share. Use native Word tables and the document's resolved Quote paragraph style; refuse unmapped metadata and unresolved styles before planning. Code blocks, thematic breaks, and block math have no native representation yet; D34 owns them.

**Close when:** A new table (with alignment, marks, and links in cells) and a new block quote each plan an explicit sibling insertion, and the shared materializer lands the target Markdown through save/reopen without changing retained siblings or non-body parts.

**Check:** Add a focused Bun test with explicit expected results. Run the C01 matrix and composer regressions.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `planTargetDocument` now accepts inserted GFM tables and block quotes, validated before planning: cells and quotes may hold mark/link paragraphs only, with no inline atoms, inline code, or nested list membership. Quotes get an explicit sibling `styles` binding to the document's own Quote paragraph style (or `options.quoteStyle`); a document without one refuses. The D01 codec validates the new binding. `packages/sciobot/src/diff/fragment-seeds.ts` is the shared materializer: `siblingFragmentSeeds` builds one native `insertParagraphs` seed list plus table seeds; `completeSiblingFragments` applies quote styles, links, and new tables after allocation (native `insertTable` before the following paragraph, equal grid widths from the governing section, cell paragraphs with column `jc`). Trailing and adjacent tables receive the empty paragraphs Word requires, which project to nothing. A01 uses this materializer. Focused run: `fragment-seeds.test.ts`, 4 tests, 43 assertions. A new table with center/right alignment, bold and linked cells plans one sibling insertion with the exact alignment and no list bindings; the plan round-trips through the codec. Applied directly, it lands the target AST, replans empty, keeps both neighbor paragraph IDs, yields one `w:tbl` with the expected cell texts and `jc` per column, and passes one-step undo, redo, and reopen. A trailing table plus an adjacent second table land as `p, tbl, p, tbl, p` with empty separators and the exact target. A one-paragraph quote with marks and a link gets `pStyle Quote` and lands exactly. Refusals: no Quote style, a list inside a quote, a multi-paragraph quote, code blocks, thematic breaks, and inline code in a cell. The diff, ai, and tracked-block suites pass 491 tests. Compatibility decision `.collaboration/changes/sciobot-inserted-tables-quotes.json`. Limits: explicit left alignment is Word's default and the native projection writes no `:---` marker; multi-paragraph quotes need P23; the tracked sibling executor (R05) still uses paragraph-only seeds.

## D29

- [x] Plan inserted registered containers and atomic blocks

**Depends:** [D15](04-edit-planner.md#d15), [D28](04-edit-planner.md#d28)

**Primary files:** `fork:packages/sciobot/src/diff/plan-target.ts`; `fork:packages/sciobot/src/diff/inserted-blocks.ts`; `sibling inserted-blocks.test.ts`

**Read:** [S04](sources.md#s04), [S09](sources.md#s09), [S25](sources.md#s25)

**Change:** Plan new lesson phases, activities, other registered containers, and atomic blocks inside retained or new parents. Carry the application payload, a deferred persistent-ID allocation key, generated-label role, and native child fragments. Seed empty containers explicitly.

**Close when:** A new activity in an existing phase, a new phase with an activity, a new empty container, and a new atomic block plan explicit operations. The materializer lands payload records and controls whose projection equals the target through save/reopen.

**Check:** Add a focused Bun test with explicit expected results. Run the C01 matrix and composer regressions.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `packages/sciobot/src/diff/inserted-blocks.ts` adds `planInsertedBlocks`. For every new registered container or atomic block it creates the application payload from target syntax (`domainNodeFromAst`, verified with D16's `payloadMatches`) under a new persistent ID. The host allocator `options.allocateBlockId` defaults to random UUIDs, and uniqueness is verified against controls and the payload store before planning. It also resolves list groups nested in new containers into nested `lists` paths. The D01 codec validates the new sibling `blocks` binding (`isInsertedBlockBindings`). The shared materializer (`fragment-seeds.ts`) renders each new block from target syntax as a tagged block content control (alias = registry tag) holding native paragraphs, headings, links, marks, nested containers, and tables; atomic blocks get a readable fallback paragraph; empty containers get one seed paragraph. It inserts the control after allocation with fresh native IDs, before the following new paragraph or the gap anchor, stores payload records in the same transaction, and returns every target paragraph path's new ID so A01 applies nested list membership. Generated `scio:label` fallbacks are not emitted; the projection does not need them. Focused run: `inserted-blocks.test.ts`, 6 tests, 72 assertions. A new activity in the existing C01 phase (bold text, link, nested bullet list) plans exact payload and list bindings, round-trips the codec, and lands through the live direct executor. The payload, parent, and tag are exact; texts are `Talk about it with notes.`, `One`, `Two`; existing block controls keep their IDs; replanning is empty; one undo removes block and payload; redo and reopen are exact. A new phase with a nested activity plus a new empty `recap_goals` land in one transaction with correct parentage and attributes. A new atomic `lesson_visual` stores its payload and fallback `Diagram`. Refusals: unknown tags (expression never evaluated), reused or invalid IDs, quotes inside new blocks. Two real sessions (both receivers): a new activity beside concurrent peer text converges, survives undo/redo, and matches saved digests. 468 diff/ai tests pass; private typecheck has no errors in these files; build verifies 118 modules; oxlint/prettier pass; compatibility decision `.collaboration/changes/sciobot-inserted-registered-blocks.json`. Limits: inline atomic sources and images are D30; opaque shared-MDX insertion remains unsupported; the tracked sibling executor (R05) still needs to adopt the materializer.

## D30

- [ ] Plan inserted inline atoms

**Depends:** [D15](04-edit-planner.md#d15), [D29](04-edit-planner.md#d29), [X23](08-domain-content.md#x23), [X38](08-domain-content.md#x38)

**Primary files:** `fork:packages/sciobot/src/diff/inserted-atoms.ts`; `sibling inserted-atoms.test.ts`

**Read:** [S03](sources.md#s03), [S04](sources.md#s04), [S25](sources.md#s25)

**Change:** Plan inline images, inline math, mentions, and inline source atoms in new and retained paragraphs. Carry asset references for executor-side media reconciliation and registered payloads for inline controls. Do not fetch or embed resources during planning.

**Close when:** Each atom kind plans explicit operations in a new and in a retained paragraph. The materializer lands target Markdown through save/reopen, with media parts added only for new images.

**Check:** Add a focused Bun test with explicit expected results. Run the C01 matrix and composer regressions.

**Claim:** codex solo, branch `codex/d30-inline-images`. **Evidence:** In progress. [Deferred descriptions](evidence/deferred-inline-atoms.json) preserve resource references without fetching. [Math projection](evidence/math-source-projection.json) retains exact LaTeX. [New-paragraph materialization](evidence/inline-atom-materialization.json) supplies native controls and payload records. Runtime `efa88073` adds [retained-paragraph atom operations](evidence/retained-inline-atoms.json). D01 decoding reconstructs complete intentions without canonical writes. Direct execution inserts math, mentions, and sources after ordinary text edits. Headings, table cells, paragraph properties, and native identities survive. Existing controls anchor independently edited text gaps; later text changes and new atoms retain original control IDs, exact control fingerprints, and payloads. Four real-session cases cover both client orders and reconnect receivers, peer text in the same paragraph, both histories, imported OMML, remote geometry, and DOCX/room reopening. A stronger isolated case preserves peer text inside newly inserted ordinary text during the inserting actor's undo/redo. Plan-wide validation rejects duplicate birth identities across separate operations before publication. The frozen baseline passes 2824 cases across 231 files. Final identity-guard verification passes 1046 affected cases across 107 files. The compiled SDK/core harness passes 16 cases and 820 assertions, including peer-text history and cross-operation identity refusal. Private types cover 2144 files; the build verifies 245 private modules with core external. Root formatting, lint, compatibility policy, and the 199-hash production audit pass. Runtime `f122f23a` adds [native inline image materialization](evidence/inline-image-materialization.json). Revision-bound preparation validates snapshotted bytes before execution. Pictures, media relationships, controls and exact source metadata land together. Standard and nonstandard main parts retain independent native media through shared histories. Managed images use native Markdown source spans, including beside hyperlinks; unregistered drawings keep their omitted/protected contract. All 1132 affected cases pass across 118 files, including C01 and composers. Compiled execution passes 16 local image cases, eight shared cases and two detached exporter cases. Private types cover 2155 files; exporter types cover 1520 files. The private build verifies 253 modules with core external. Runtime `44b384ec` adds [new-cell inline atoms](evidence/cell-inline-atoms.json) for retained table row/column insertion, native splits, combined topology and cell lists. Cell-prefixed D01 bindings reconstruct exactly; payload locks finalize after topology. All 1309 affected cases pass across 132 files, including C01 and composers. The compiled SDK/core/exporter passes 15 cell cases. Types cover 2159 files; the private build verifies 255 modules. A separate [failing regression](evidence/regressions/peer-born-cell-undo.test.ts) proves that row/column birth undo hides peer text entered inside the new cells. Both source and compiled cases fail; retained-cell peer checks do not establish this behavior. That failure, new nested/registered cell fragments, linked/code-marked atoms, retained hyperlink ownership, image removal/replacement, tracked ownership, and installed application checks remain open. No D30 closure is claimed.

**Isolated follow-up:** Branch `codex/d30-linked-atoms` follows the unmerged native history repair. [Table history evidence](evidence/born-cell-history-repair.md) records peer-owned cell placement, geometry and moves. Runtime `38f89183` adds [linked inline control execution](evidence/linked-inline-controls.md); `6f2cacea` verifies imported link identities, peer text, both histories and cold reopening. All 1063 affected cases pass across 103 files, including C01 and review composition. Types cover 2172 files; the private build verifies 256 modules with core external. Root format, lint and compatibility policy pass. All 199 pinned production source hashes match, but the stock inventory audit rejects the changed application revision. These checkpoints remain isolated. Linked drawings, code-marked atoms, changed retained link ownership, further nested cell fragments, tracked births and installed application verification remain open. The checkbox stays unchecked.

**Code follow-up:** Native repair `472ff8e5` and the private adapter support code-styled controls through local and shared histories. Broad and focused checks cover 1,930 distinct passing source cases; 24 compiled code fixtures also pass. Exact native provenance now passes the stronger 32-case middle-fragment peer matrix and all 56 compiled code fixtures. Logical fragment mapping repairs accepted paragraph-mark reads and edits; 369 tracking cases pass. Runtime `f7ca8c5e` repairs both inherited tracked-row acceptance fingerprints. All 657 core automation cases pass. Earlier readers preserve exact native states through edits and both histories before and after acceptance. See [table holder evidence](evidence/tracked-table-property-holders.md) and [joined paragraph evidence](evidence/joined-paragraph-automation.md). Runtime `956ba1e9` adds [retained atom mark changes](evidence/retained-atom-mark-changes.md). The four atom kinds retain native controls and media through bold, italic, strike and code changes; 22 focused cases pass, including both peer histories and reopening. Runtime `fc6cd155` adds [retained atom hyperlink changes](evidence/retained-atom-link-changes.md). Runtime `b8764ccc` preserves exact ordinary typing boundaries; all twelve shared link/history cases pass unchanged. Runtime `2b4fe004` resolves the four original linked image birth probes. Native reservation replacement retains links and target marks; 86 compiled cases pass. Runtime `d11c08c7` and follow-up `c37f8096` repair imported repeated-space link ownership while preserving existing tracked leading-link insertion. All 91 compiled cases pass. Runtime `827f0820` / `9c8a08c8` repairs the 64 recorded composer regressions. The final complete C01/composer gate passes all 333 unchanged cases; 159 compiled atom/code cases also pass. See [shared history repair evidence](evidence/composer-shared-history-repairs.md). Runtime `7234d5a0` / `7ea0a3d0` resolves both original fresh/retained linked-atom probes. See [mixed atom evidence](evidence/mixed-inline-atoms.md). See [hyperlink text ownership evidence](evidence/hyperlink-text-ownership.md). This is implementation progress, not D30 closure.

**Mixed atom follow-up:** Branch `codex/d30-mixed-linked-atoms`, runtime `7ea0a3d0`. Fresh atoms preserve linked native math, images, mentions, and sources. Forty-eight local cases verify serialized plans, repeated-text anchors, complete control fingerprints, media, history, and reopening. Thirty-two shared cases verify both actor histories, independent native peer typing, geometry, and DOCX/room reopening. The compiled SDK/core/exporter passes 239 cases; native checks pass 1,028 cases; all 333 unchanged C01/composer cases pass. Core/private types, builds, API checks, root format/lint, and compatibility checks pass. See [commands and limits](evidence/mixed-inline-atoms.md). The production inventory still rejects application revision drift; all 199 pinned source hashes match. D30 remains open.

**Image replacement checkpoint:** Runtime `42eaeaf0` preserves native image identity, geometry, payloads, and media through concurrent replacements and both actor histories. The [image replacement evidence](evidence/inline-image-replacement.md) records the combined 255-case native batch and 88-case built SDK batch. D30 remains open; this checkpoint does not substitute for its closure audit.

**Inline removal checkpoint:** Runtime `72c1e85c` adds explicit native removals for images, inline math, mentions, and sources. Compound edits retain source updates and surviving controls. Payload retirement and image cleanup share owned history. All 27 built removal cases pass, including both client orders, both histories, and DOCX/room reopening. The [removal evidence](evidence/inline-atom-removal.md) records corrected C01 and compound regressions, source timing limits, and remaining scope. D30 remains open.

Additional WIP evidence: [tracked atoms in born paragraphs](evidence/tracked-born-inline-paragraphs.md) records native materialization and the remaining creation-undo peer-text failures.

[Independent paragraph history](evidence/independent-paragraph-peer-history.md) records retained peer text through creation undo/redo and the remaining native rejection ownership failure.

[Committed insertion ownership](evidence/committed-insertion-ownership.md) resolves peer-text rejection with 352 source cases and 43 compiled cases passing. D30 remains open for the listed feature and integration work.

[Native math control deletion](evidence/tracked-math-control-removal.md) preserves equations and hyperlinks through AI review, shared histories, and reopening. The source batch passes 98 cases; compiled packages pass 59 cases. D30 remains open for the remaining feature and integration work.

[Tracked atom link placement](evidence/tracked-atom-link-placement.md) adds new links, bookmarks, and retained link boundaries without replacing source link identities. The source batch passes 120 cases; compiled packages pass 80 cases. D30 remains open for tracked images, new-cell ownership, mixed ownership, and integration gates.

[Interior peer typing regression](evidence/linked-source-peer-regression.md) exposes misplaced text during concurrent link proposals. Source and compiled checks each pass the four outside-range cases and fail the four interior cases. The verified writer remains unchanged; preserving source leaf identity alone hides peer text in a deleted branch and is not a fix.

[Native copy provenance WIP](evidence/tracked-text-copy-recipes.md) records exact source recipes for both hyperlink review copies. All 77 native cases pass. The isolated transport still fails the four interior peer cases; paired review copies and causal revision membership remain required.

[Paired review peer copies](evidence/paired-review-peer-copies.md) repairs the interior hyperlink regression through exact decisions, both histories, and reopening. The evaluated format changes to `1.6.6.1`. Four prior composer failures remain reproducible; D30 stays open.

[Tracked image staging](evidence/tracked-inline-image-staging.md) adds revision-bound media preparation and native image presence ownership. Verification remains in progress; the checkbox stays open.

[Tracked image removal](evidence/tracked-image-removal.md) includes live-reference recovery in `2c47facc` and helper extraction in `3e5784fa`. Source, compiled, history, fidelity, and resource-cost checks pass the recorded batches. D30 remains open for the complete integration gates.

[Replacement cell coverage](evidence/replacement-cell-coverage.md) repairs the eight mixed-split history regressions through exact native row and column identities. The source batch passes 147 cases; compiled checks pass 58 cases. Four prior committed-view regressions remain unresolved. D30 stays open.

[Tracked new-cell atoms](evidence/tracked-cell-inline-atoms.md) adds row, column, and split atom execution with persistent native topology ownership. Eighteen compiled cases pass both decisions and shared histories. Native history recovery uses evaluated format `1.7.9.1`. D30 stays open.

[Repeated cell waves and new-row peer retention](evidence/tracked-cell-atom-peer-retention.md) verifies both decisions and creation histories. The evaluated transport repair advances to `1.7.10.1`. D30 stays open.

[Tracked atoms in complete new tables](evidence/tracked-new-table-inline-atoms.md) adds cell staging and private sibling-gap ownership. Middle, terminal, replacement, mixed-table, and repeated-wave cases retain exact native decisions. D30 stays open.

## D31

- [x] Plan inline code marks through the native representation

**Depends:** [D26](04-edit-planner.md#d26), [N18](07-native-content.md#n18)

**Primary files:** `fork:packages/sciobot/src/diff/target-inline.ts`; `sibling plan-target.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Map inline code added to or removed from retained text, and inline code inside inserted text, to the canonical inline-code representation established by N18.

**Close when:** Adding and removing inline code in retained paragraphs and inserting new inline code plan explicit operations that land the target Markdown through save/reopen.

**Check:** Add a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof` (continuing Claude). **Evidence:** Native baseline `9770e78` plus private snapshot `0287a034937b7b98294415098bdeaa410cd5736f`. Reviewed `target-inline.ts`, `insertion-marks.ts`, and `target-inline-code.test.ts` after N18/N20 closure. Six focused planner cases explicitly assert retained code additions/removals, inserted code beside existing code, combined toggle/link changes, new paragraphs/table cells, and reversible imported character-style inheritance. Tests apply real direct canonical plans, assert exact target Markdown, replan to no operations, undo to the original semantic fingerprint, and save/reopen to the same target. The complete N20 command includes these cases: `bun test --preload ./scripts/test/preload-dom.ts --timeout 60000 packages/sciobot/src/native/inline-code.test.ts packages/sciobot/src/native/prepare-document.test.ts packages/sciobot/src/native/inline-code-collaboration.test.ts packages/sciobot/src/diff/target-inline-code.test.ts packages/core/src/store/__tests__/run-character-style.test.ts packages/docx-to-markdown/test/markdown-code-spans.test.ts`; 66 pass, 1030 assertions, `/tmp/codex-n20-native-final.log`. Native/private typecheck/build, scoped format/lint, public API snapshots, and compatibility checks pass. Existing published changesets belong to N18. This closes pure planner support and direct semantic oracles; R38 retains tracked paragraph execution.

## D32

- [x] Plan logical streams across native barriers and composite owners

**Depends:** [D15](04-edit-planner.md#d15)

**Primary files:** `fork:packages/sciobot/src/diff/plan-target.ts`; `fork:packages/sciobot/src/diff/target-barriers.ts`; `sibling target-barriers.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Split complete targets at native sibling barriers such as unprojected opaque blocks, mid-body section breaks, and bookmarks between blocks. Plan each physical window separately and keep barrier nodes in place. Handle continuation and mixed list ownership.

**Close when:** Targets around an opaque block, a mid-body section break, and a continued list plan explicit operations. Barrier nodes keep identity and fingerprints through save/reopen.

**Check:** Add a focused Bun test with explicit expected results. Run the C01 matrix and composer regressions.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. When a logical stream spans several D22 physical sibling windows (separated by unprojected native nodes), `planTargetDocument` now covers the stream with those windows in order. It aligns target blocks to source blocks by content keys (bounded LCS) and plans each window as its own stream with logical-path offsets. New blocks that fall between matched blocks of two windows stay in the earlier window, before the barrier. Mid-body section breaks and lists continued across a plain paragraph already planned inside one window; they are now covered by explicit tests. Focused run: `target-barriers.test.ts`, 3 tests, 26 assertions, each landing through the live direct executor with empty replanning, the barrier node kept by identity, one undo step, and reopen with an unchanged barrier fingerprint. An opaque `x:block` between paragraphs, with edits and insertions on both sides, plans `text, siblings, siblings` and lands in the exact native order `Alpha changed., Beta., Before barrier., block, Gamma., After barrier.`. Edits around a paragraph carrying a mid-body `sectPr` keep that section mark. A list continued across a plain paragraph gains `3. Three` with native numbering and a text edit. 481 diff/ai tests pass; private typecheck has no errors in these files; oxlint passes; compatibility decision `.collaboration/changes/sciobot-target-barriers.json`. Limit: logical blocks without a structural native owner (composite owners) still refuse; no captured fixture produces one.

## D33

- [x] Compose retained table merge and split intentions

**Depends:** [D15](04-edit-planner.md#d15), [D28](04-edit-planner.md#d28)

**Primary files:** `fork:packages/sciobot/src/diff/target-table.ts`; `fork:packages/sciobot/src/diff/target-table-spans.ts`; `fork:packages/sciobot/src/diff/target-table-merges.ts`; `sibling target-table-merges.test.ts`

**Read:** [S03](sources.md#s03), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Dispatch merge and split requests in complete targets through the existing D10 planner. D35, D36, D37, and D38 carry the remaining retained table structure intentions split from this leaf: multi-paragraph GFM cells, tables without complete declared grid anchors, new property destinations, and lists inside new rows.

**Close when:** Each merge and split case plans explicit operations and lands the target through save/reopen with original table, row, and cell identities where they survive.

**Check:** Add a focused Bun test with explicit expected results.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `expandTargetSpans` turns explicit HTML `colspan`/`rowspan` into a regular slot grid with empty covered slots. `sourceSlotGrid` reads native merge rectangles from D22 placeholder bindings. `targetGeometryRequests` compares the partitions: an empty target slot inside a native merge keeps it (projections never emit spans), content or a smaller explicit span inside it requests a regular split, and a target region covering several complete native cells requests a merge. Crossing, ragged, overlapping, out-of-table, and unequal-split targets refuse before a plan returns. Merges keep the trailing covered cells whose complete content the target retains as its suffix, empty the earlier covered cells, give the head the rest, and then emit native `mergeTableCells`; native reading-order merging then reproduces the target. Splits keep the head cell, plan its content in place, and emit `splitTableCell` with existing-grid distribution, or the new private `split-cell` intent when new cells carry target fragments; `applyDirectPlan` fills those cells in its after-allocation phase, located by native grid position. Emptying a cell now keeps one blank paragraph (`blankCellTarget`) instead of deleting its only block, which the direct executor had refused with `block-required`. `target-table-merges.test.ts`: 8 cases, 92 assertions, each landing through the live direct executor with explicit native layout (text, `gridSpan`, `vMerge`), empty replanning, one-step undo to the original table fingerprint, redo, and reopen with identical layout and empty replanning. Cases: colspan merge with new head text (`text,text,table`, the untouched neighbor keeps object identity); merge into a blank head that moves the covered paragraph by identity; rowspan merge dropping a blank covered cell; placeholder content splitting a 2x2 cell into its grid (`text,split-cell`, rows 2 by columns 2, head keeps its node); explicit rowspans splitting a 2x2 cell into two vertical bands; an emptied placeholder keeping its merge; an emptied cell keeping one blank paragraph; and three atomic refusals with unchanged bytes. `target-table-merges-collaboration.test.ts`: 4 real two-session cases, 68 assertions, both receivers, fixed clients 73/41: a merge while the peer edits the covered paragraph that moves (converges on `Two Bob` in the merged head), and a split while the peer edits a neighbor cell; each checks convergence, remote caret geometry, undo/redo topology, peer deletion during concurrent undo, and equal saved digests. Regression: `bun test --preload ../../scripts/test/preload-dom.ts src/diff src/ai` from `packages/sciobot` passes 490 cases, 12630 assertions, before the new files. Private typecheck passes; private build verifies 125 modules; scoped Prettier/oxlint and compatibility (`.collaboration/changes/sciobot-target-table-merges.json`, 71 decisions) pass. R10 must lower or refuse the new `split-cell` intent for tracked suggestions. Limits: combining merge/split with simultaneous row or column changes refuses; a logically empty covered cell whose native runs are not blank (for example empty runs with foreign properties) still contributes an empty paragraph to a native merge.

## D34

- [ ] Plan inserted code blocks, thematic breaks, and block math

**Depends:** [N23](07-native-content.md#n23), [D28](04-edit-planner.md#d28), [X16](08-domain-content.md#x16), [X23](08-domain-content.md#x23), [N02](07-native-content.md#n02)

**Primary files:** `fork:packages/sciobot/src/diff/fragment-seeds.ts`; `sibling fragment-seeds.test.ts`

**Read:** [S03](sources.md#s03), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Once X16, X23, and N02 define the native code block, equation, and horizontal rule representations and their Markdown projections, accept those constructs as inserted fragments through the shared materializer.

**Close when:** A new fenced code block with a language, a thematic break, and a block equation each plan an explicit sibling insertion that lands the target Markdown through save/reopen.

**Check:** Add a focused Bun test with explicit expected results.

**Claim:** unclaimed. **Evidence:** pending. The native Markdown projection currently emits code paragraphs as plain text, has no horizontal-rule syntax, and emits equations as fallback text, so these insertions cannot land their target yet.

## D35

- [x] Plan edits inside multi-paragraph GFM table cells

**Depends:** [D15](04-edit-planner.md#d15), [D23](04-edit-planner.md#d23)

**Primary files:** `fork:packages/sciobot/src/diff/target-table-content.ts`; `fork:packages/sciobot/src/diff/target-generated-separators.ts`; `sibling target-table.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** A GFM cell with several native paragraphs projects as one synthetic paragraph joined by generated `<br>` separators, which currently refuses as a composite owner. Use the source map to tell generated paragraph separators from authored manual breaks, and map text, mark, and paragraph insertion/removal edits back to the native cell paragraphs.

**Close when:** Editing the second paragraph, adding a paragraph, and removing a paragraph in a two-paragraph GFM cell plan explicit operations that keep surviving paragraph IDs and land the target through save/reopen. Authored manual breaks stay manual breaks.

**Check:** Add a focused Bun test with explicit expected results. Run the C01 matrix and table composer regressions.

**Claim:** codex D35 worker, branch `codex/react-block-proof`. **Evidence:** Native baseline `3278fa04`, local working tree. `target-cell-paragraphs.ts` splits GFM cell streams only at source-proven generated separators between canonical paragraph contributions. It maps native bindings and sibling windows to the resulting paragraphs. Existing text, mark, and sibling planners retain surviving paragraph IDs. Text-diff mappings keep carried authored breaks inside their original paragraphs; new separators in an existing multi-paragraph cell denote new paragraphs. Full-target equality returns the unchanged bound view. Body quote paragraph objects retain identity for D39. `bun test packages/sciobot/src/diff/target-cell-paragraphs.test.ts packages/sciobot/src/diff/target-table.test.ts packages/sciobot/src/diff/target-c01-matrix.test.ts packages/sciobot/src/diff/target-table-properties.test.ts packages/sciobot/src/diff/target-table-merges.test.ts` passes 57 cases and 581 assertions. The final owner, Unicode, repeated-text, and no-op additions pass with D39: `bun test packages/sciobot/src/diff/target-cell-paragraphs.test.ts packages/sciobot/src/diff/target-quotes.test.ts`, 26 cases and 334 assertions. D35 contributes 13 focused cases. Explicit native text, operation kinds, paragraph IDs, unchanged source bytes, save/reopen output, and repeated no-op plans supply the oracles. Scoped Prettier, oxlint, and diff checks pass. Coordinator verification passes private types across 1972 files and the 195-module build, with native core external. No-impact decision: `.collaboration/changes/sciobot-gfm-cell-paragraph-planning.json`. This private pure planner needs no new collaboration mutation scenario. Existing direct/review executors retain live mutation authority. Ambiguous or unsupported composite source mappings still refuse before writes; single-paragraph cells retain existing manual-break behavior.

## D36

- [ ] Compose column changes in tables without complete declared grid anchors

**Depends:** [D24](04-edit-planner.md#d24), [N19](07-native-content.md#n19)

**Primary files:** `fork:packages/sciobot/src/diff/target-table-columns.ts`; `sibling target-table-columns.test.ts`

**Read:** [S03](sources.md#s03), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Address column insertions and deletions through first-row reference cells when `w:tblGrid` is absent or incomplete, using the N19 native operations. Keep every planned address an original canonical ID.

**Close when:** Insertion and deletion of several columns in a table without `w:tblGrid` and in a table with an incomplete grid plan explicit operations and land the target through save/reopen with surviving cell identities.

**Check:** Add a focused Bun test with explicit expected results.

**Claim:** unclaimed. **Evidence:** pending. Native `insertTableColumn` accepts a reference cell only while the grid is absent, and `deleteTableColumn` requires a grid column ID, so a second change in the same plan would need preview-allocated grid IDs.

## D37

- [x] Plan property destinations for new rows, columns, and split cells

**Depends:** [D25](04-edit-planner.md#d25), [D16](04-edit-planner.md#d16), [D17](04-edit-planner.md#d17), [D33](04-edit-planner.md#d33)

**Primary files:** `fork:packages/sciobot/src/diff/target-table-properties.ts`; `fork:packages/sciobot/src/diff/inserted-row-types.ts`; `sibling target-table-properties.test.ts`

**Read:** [S03](sources.md#s03), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Carry fill, borders, vertical alignment, and paragraph alignment for new rows, new columns, and new split cells as deferred properties applied after native allocation. Keep GFM column alignment for new rows and columns.

**Close when:** A new row with a filled bordered cell, a new aligned column, and a split cell with a fill each plan explicit operations and land the target through save/reopen.

**Check:** Add a focused Bun test with explicit expected results. Extend the A01 executor and record the tracked-executor requirement for R08-R10.

**Claim:** codex D37 worker, branch `codex/react-block-proof`. **Evidence:** Native baseline `bd8418db`, local working tree. Private row, column, fixed-grid split, and compound split intentions carry optional validated `cellProperties` bindings by physical cell index. `new-cell-properties.ts` resolves table/column/row/cell alignment inheritance and reuses native fill, border, vertical-alignment, and paragraph-property operations after allocation. Plans retain original anchors and never expose preview-generated IDs. GFM alignment reaches both retained and new cell paragraphs. A01 applies properties in the allocation transaction. R08/R09 live executors and the shared R08-R10 detached table stage carry the same bindings; inserted row/cell revisions own the new properties. Legacy R10's merge-only entry point retains its existing scope; rich split intentions use the shared stage. Six explicit direct/staged cases cover new rows, columns, and split cells with exact fill `123456`, four red single borders of size 8, bottom vertical alignment, center paragraph alignment, unchanged source bytes, codec round trips, and save/reopen. Two cases cover GFM column inheritance, and codec tests reject malformed, duplicate, foreign, and invalid property destinations. The prior new-column alignment refusal now tests an unsupported new width after an earlier text edit, retaining atomic-refusal coverage. Six real two-session cases cover row/column/split changes, both client orders, opposite reconnect receivers, offline peer suffixes, explicit native properties, peer-preserving undo/redo, caret geometry, deletion during peer editing, and save/reopen convergence. Final focused command `bun test --timeout 30000 packages/sciobot/src/diff/target-new-cell-properties{,-collaboration}.test.ts packages/sciobot/src/diff/target-table-properties.test.ts packages/sciobot/src/diff/target-table-merges.test.ts packages/sciobot/src/diff/target-new-cell-lists.test.ts` passes 47 cases and 876 assertions (`/tmp/d37-focused.log`). `bun test --timeout 30000 packages/sciobot/src/diff/new-cell-properties-review.test.ts` passes two live tracked cases and 14 assertions (`/tmp/d37-review-final.log`): rejection restores the exact original native fingerprint; acceptance matches direct native XML after removing only independently allocated `w14:paraId`/`w14:textId` attributes, as the existing R08 oracle does; save/reopen retains the exact accepted fingerprint. Scoped formatting/lint and whitespace checks pass. The backlog command reaches the production-coverage guard and requests an audited capture; root owns that integration update (`/tmp/d37-backlog-final.log`). Root integration confirms private types across 1975 files, private build with 196 modules, and repository lint. Compatibility decision: `.collaboration/changes/sciobot-new-cell-property-destinations.json`. No public API, shared operation schema, or published package changes. Explicit null alignment follows retained-cell inheritance semantics. New width destinations and table-wide border destinations retain explicit refusal. Nested-table fragments remain unsupported; the R41 seam records that future outer-cell alignment must preserve explicit inner-table alignment.

## D38

- [x] Plan lists inside new table rows and columns

**Depends:** [D06](04-edit-planner.md#d06), [D16](04-edit-planner.md#d16), [D17](04-edit-planner.md#d17)

**Primary files:** `fork:packages/sciobot/src/diff/target-table-fragments.ts`; `fork:packages/sciobot/src/diff/inserted-row-types.ts`; `sibling target-table.test.ts`

**Read:** [S03](sources.md#s03), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Bind list and checklist targets for paragraphs in new cells, by cell and paragraph path, and resolve them after native allocation as sibling insertions already do.

**Close when:** A new row with a bullet list, an ordered list, and a checklist in its cells plans explicit operations and lands the target through save/reopen.

**Check:** Add a focused Bun test with explicit expected results.

**Claim:** codex D38 worker, branch `codex/react-block-proof`. **Evidence:** Native baseline `5b5fa695`, local working tree. Private row/column intentions carry optional validated `[physicalCellIndex, paragraphIndex]` list destinations. Native allocation resolves these paths to paragraph IDs; existing numbering and checkbox operations own the writes. Direct, live tracked, and detached tracked executors publish content and numbering together. Row/cell insertion revisions own new list paragraphs and checkbox controls. Checklist seed controls carry their own `w14` namespace, including detached package stages. One allocation context per plan execution shares repeated new-list keys and authored starts across intents; preview and commit contexts remain separate. Codecs reject malformed, duplicate, nested, out-of-range, and conflicting destinations. The former new-row-list refusal test now proves atomic refusal for an unsupported list continuation paragraph after an earlier text edit. The final seven-file command `bun test --timeout 30000 packages/sciobot/src/diff/new-cell-list-context.test.ts packages/sciobot/src/diff/target-new-cell-lists.test.ts packages/sciobot/src/diff/inserted-cell-lists.test.ts packages/sciobot/src/diff/target-new-cell-lists-collaboration.test.ts packages/sciobot/src/ai/apply-direct.test.ts packages/sciobot/src/review/track-table-rows.test.ts packages/sciobot/src/review/track-table-columns.test.ts` passes 56 cases and 906 assertions (`/tmp/d38-final-seven.log`). The earlier 12-file affected suite passed 126 cases and 1745 assertions before the allocation-context review fix. Six focused direct/tracked/detached cases cover bullet, ordered, and checklist cells, exact path bindings, reject/undo/accept, native numbering and checked state, and save/reopen. Four allocation cases prove one shared native instance and one start=3 override across separate row/column/ordinary intents. Four real two-session cases cover both client orders, offline peer text, opposite reconnect receivers, peer-preserving undo/redo, caret geometry, save/reopen convergence, and row/column deletion during peer editing. Scoped formatting, lint, and whitespace checks pass. Aggregate private types/build remain with the root integration review. Full target AST comparison ignores only Markdown `spread` metadata, matching C01: Word paragraphs do not retain tight/loose list metadata, and native export emits loose lists. Paragraph content, list nesting, numbering identity/start, and checkbox state remain explicit assertions. Compatibility decision: `.collaboration/changes/sciobot-new-cell-list-destinations.json`. No public API, native operation schema, or published package changes. D37 still owns new property destinations; merge/split list destinations retain their explicit refusal.

Coordinator integration: private types pass 1972 files; the private build verifies 195 modules with external core. D38 code is committed in `bd8418db`.

## D39

- [x] Plan edits inside retained native quote paragraphs

**Depends:** [D15](04-edit-planner.md#d15), [P23](03-agent-projection.md#p23)

**Primary files:** `fork:packages/sciobot/src/diff/plan-target.ts`; `sibling target-quotes.ts`; `sibling target-quotes.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** A projected blockquote has no native container: its children are Quote-styled body paragraphs. Bind the logical quote container to its native paragraph window so retained quote paragraphs plan text, mark, insertion, and removal edits, and so a paragraph can enter or leave a quote through a style change.

**Close when:** Editing one paragraph of a two-paragraph quote, adding a paragraph to it, removing one, and unquoting a paragraph each plan explicit operations that keep surviving paragraph IDs and land the target through save/reopen.

**Check:** Add a focused Bun test with explicit expected results. Run the C01 matrix and composer regressions.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Source baseline `3278fa04`, local working tree. A private planning view expands ordinary body quote children into their actual sibling paragraphs and retains original logical paths, exact source contributions, and native windows. Existing source bindings grant ownership; no quote container or native range is fabricated. Text/mark edits and paragraph insertion/removal retain surviving IDs. Entering/leaving a quote emits explicit native paragraph-style deltas. Adjacent new quote paragraphs share the existing fragment representation. Five explicit local cases preserve unrelated native objects and metadata, source immutability, plan codec round trips, exact target ASTs, and save/reopen. Eight real-session cases cover both client orders and reconnect receivers, offline peer suffixes, deletion during affected-content editing, exact native identities, peer-preserving undo/redo, caret geometry, and saved convergence. Final focused command `bun test --timeout 30000 --preload ./scripts/test/preload-dom.ts packages/sciobot/src/diff/target-quotes.test.ts` passes 13 cases and 269 assertions; `/tmp/codex-d39-integration-final.log`. Composite C01/composer verification covers 111 distinct cases across eight files: 93 existing composer/source cases, five fragment-seed cases, and these 13 quote cases. The fragment-seed file's unchanged multi-paragraph quote expectation passes after preserving grouped fragment shape. Private source types pass 1958 files; build verifies 189 private modules, one external native core, and separate optional runtime entries. Scoped formatting/lint, root lint, compatibility, reference coverage, and backlog checks pass. Compatibility passes 124 decisions against `upstream/main`, zero additional format fields; `.collaboration/changes/sciobot-retained-quote-planning.json`. No published code, API, or collaboration schema changes require a changeset or API extraction. The audit retains all 198 application source hashes; its reference revision alone advances to `aec740a325e199f0a46069130d70ee65c88c524f`. Nested or non-paragraph quote syntax retains its existing explicit refusal; D39 owns the projected ordinary body-quote representation. D35 separately owns composite GFM cell paragraphs.

**Tracked inline ownership checkpoint:** Runtime `44f01237` couples complete inline payloads to exact native review receipts. Proposal schema 3 preserves native ID reservations after decisions and tombstoning; older built readers refuse all three states. Local and shared decisions, owned inverses, history, source run boundaries, and source hyperlinks pass the recorded targeted checks. See [evidence and remaining scope](evidence/tracked-inline-ownership.md). D30 and R19 remain open.

**Tracked removal coordinator checkpoint:** Runtime `3b388abc` applies removal-only AI targets through the actual native transaction writer and captures schema 3 payload ownership. Eleven built-SDK cases pass, including locked controls, both decisions, exact rejection, history, reopening, and repeated AI waves. Four shared coordinator cases pass both client orders and peer histories. Dependent mixed coordinates, births, image resources, OMML, and final application verification remain open. See [checkpoint evidence](evidence/tracked-inline-ownership.md#ai-coordinator-removal-checkpoint). D30 and R19 remain unchecked.

**Mixed tracked removal checkpoint:** Runtime `7e93de92` reconciles text and format coordinates around native atom deletions. Native source scopes restore exact original run boundaries. Independent decisions and shared histories pass the recorded checks. Guard `cb97baea` rejects proposed grapheme splits before publication; all 22 built-SDK cases pass. See [evidence and verification status](evidence/mixed-inline-review.md). D30 and R19 remain unchecked.

**Tracked retained birth checkpoint:** Runtime `a7a2f968` stages inline math, mentions, and source births through native review with complete payload ownership. Nineteen local cases and four shared coordinator cases pass. All 43 built-SDK cases pass, including mixed removal and independent decisions. A broader owned-presence run exposed an inherited offline suffix placement failure, reproduced before the birth change. See [evidence and remaining work](evidence/tracked-retained-inline-births.md). D30 and R19 remain open.
