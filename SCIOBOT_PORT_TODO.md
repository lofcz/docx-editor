# Sciobot Plate replacement TODO

Port the production workspace editor to the DOCX fork. Keep AI editing, suggestions, and React blocks as required capabilities.

The detailed inventory below is traceability, not the cutover gate. The initial replacement is viable when the six active gates below pass in `sciobot-next`; the remaining leaves track product parity and stay open until their implementation and verification pass.

The React block proof supplies a tested starting point. It does not establish full application parity.

## Active cutover gates

Keep this list small and reviewable. Close a gate only with a command, fixture, or application commit recorded in the linked leaf evidence.

| Gate | Done when | Existing leaves |
| --- | --- | --- |
| **CUT-01 Integration shell** | `sciobot-next` can load, create, save, reopen, and export a DOCX through one private adapter; Plate is no longer on the DOCX path. | C03, C08, B01, B04, I01, I02, I03 |
| **CUT-02 Word-first editing** | Native paragraphs, lists, tables, images, links, breaks, equations, and tracked changes round-trip through the adapter with undo/redo. | N01, N02, N03, N04, N05, N06, N07, N08, D34 |
| **CUT-03 AI direct edits** | The existing AI tool contract can target a selection or block, stream progress, commit atomically, and survive save/reopen. | C01, C02, C07, D01, D02, D03, D04, D05, D06, D07, D08, D09, D10 |
| **CUT-04 Suggestions and review** | AI suggestions preserve user edits, support accept/reject/undo, property changes, and reload without losing pending state. | C05, S01, S02, S03, S04, S05, S06, S07, S08, S09, S10, S11, S12, S13, S14, S15, R19 |
| **CUT-05 Arbitrary React blocks** | A production React block can be inserted, edited, moved, persisted, restored, and collaborated on without Plate serializers. | C03, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17, B18, B19, B20, B21, B22, B23 |
| **CUT-06 Collaboration and release evidence** | Two real sessions converge for concurrent edits and review waves; runtime dependencies and production coverage are verified. | V01, V02, V03, V04, V05, V06, V07, V08 |

The workstream tables retain the remaining implementation work. Do not create a new leaf for a cosmetic or domain-specific port unless it blocks a gate.

Read the [implementation notes](docs/engineering/sciobot-port/evidence/implementation-notes.md) before extending the native adapter.

## Baseline and scope

- Work in [/run/media/lofcz/ssd_external/GitHub/docx-editor](/run/media/lofcz/ssd_external/GitHub/docx-editor), branch `codex/react-block-proof`.
- Preserve the existing proof changes. Commit each coherent batch and keep local uncommitted changes minimal.
- Resolve application and Plate paths through the [source map](docs/engineering/sciobot-port/sources.md).
- Follow the actual [production coverage ledger](docs/engineering/sciobot-port/coverage.md), including static and export consumers.
- Reuse Apache core operations and application-owned adapters. The production path needs no Pro or editor-api runtime.
- Treat the canonical OOXML store as authoritative. Never rebuild imported DOCX from whole-document Markdown.
- Keep public document APIs Office.js-shaped. Keep FileExecutor integration private and separate from that public model.

## First dispatch wave

These leaves have no task prerequisites. Read their checkboxes before claiming one. Claim distinct files when running workers in parallel.

| Leaf | Deliverable |
| --- | --- |
| [C01](docs/engineering/sciobot-port/01-contracts.md#c01) | Capture the seven AI parity fixtures |
| [C02](docs/engineering/sciobot-port/01-contracts.md#c02) | Capture document tool request and result contracts |
| [C03](docs/engineering/sciobot-port/01-contracts.md#c03) | Create the private application adapter package |
| [C05](docs/engineering/sciobot-port/01-contracts.md#c05) | Capture replacement matcher regression cases |
| [C06](docs/engineering/sciobot-port/01-contracts.md#c06) | Expand the production feature coverage ledger |
| [N01](docs/engineering/sciobot-port/07-native-content.md#n01) | Create a native command capability fixture |
| [V01](docs/engineering/sciobot-port/11-verification.md#v01) | Generalize the two-editor collaboration fixture |

A coordinator can start C03, C01, C02, and V01 independently. C05, C06, and N01 also have no prerequisites.

C03 opens C08, B01, and B04. C07 also needs C02. Each feature still needs its listed verification prerequisites.

## Workstreams

| Workstream | Leaves | Focus |
| --- | --- | --- |
| [Contracts and reference fixtures](docs/engineering/sciobot-port/01-contracts.md) | 8 | Freeze observable Plate behavior before replacing its implementation. These leaves provide independent expected results. |
| [Production React block host](docs/engineering/sciobot-port/02-react-blocks.md) | 33 | Build on examples/react-blocks. The proof remains evidence, rather than the production implementation. |
| [Agent document projection and targeting](docs/engineering/sciobot-port/03-agent-projection.md) | 23 | Keep the existing Markdown tool protocol. Map targeted changes back to canonical OOXML without rebuilding the document. |
| [Small canonical edit plans](docs/engineering/sciobot-port/04-edit-planner.md) | 39 | Produce pure plans first. Direct and suggestion executors consume the same plans. |
| [Suggestion state and review operations](docs/engineering/sciobot-port/05-suggestions.md) | 53 | Native revisions and application property proposals must form one review model. Required cases cannot close through refusal alone. |
| [Direct AI edits and streaming lifecycle](docs/engineering/sciobot-port/06-direct-ai.md) | 10 | Direct edits commit before animation. Presentation never owns document state. |
| [Native Word editing controls](docs/engineering/sciobot-port/07-native-content.md) | 23 | Reuse existing native commands. Close a leaf with verified production controls when the engine already provides the behavior. |
| [Application content and custom renderers](docs/engineering/sciobot-port/08-domain-content.md) | 37 | Reuse application services and content UI. Remove Plate-specific wrappers from each renderer separately. |
| [Editor interactions and review presentation](docs/engineering/sciobot-port/09-editor-ui.md) | 16 | Connect existing application controls to native selection, commands, and private review state. |
| [Application persistence and editor replacement](docs/engineering/sciobot-port/10-app-integration.md) | 25 | Work in isolated application branches. Keep unrelated dirty files intact. Preserve existing per-path tools and export workflows. |
| [Cross-feature evidence and cutover gates](docs/engineering/sciobot-port/11-verification.md) | 8 | These leaves collect required evidence. Passing one gate does not replace focused checks on each feature. |

The essential integration path runs through shared contracts, mapped projections, pure plans, review operations, and FileExecutor integration.

React block hosting can progress beside that path. Application renderers consume its stable contracts.

## Agent dispatch rules

1. Read the checkout instructions and the selected leaf sources before editing.
2. Claim one unchecked leaf with completed dependencies. Set its Claim field to your identifier and branch.
3. Confirm every dependency is closed with evidence. An unchecked dependency is not ready.
4. Keep one focused behavior and one reviewable change per leaf. Split work that exceeds one focused agent session.
5. Use the primary files as ownership boundaries. Coordinate shared contracts, package manifests, and lockfiles through one writer.
6. For parallel worktrees, first preserve the local proof and backlog baseline in each worktree. Upstream HEAD lacks that baseline.
7. Keep unrelated dirty files intact. Use one writer per checkout when baseline isolation is unavailable.
8. If a required core operation is missing, add a small implementation leaf and make the blocked feature depend on it.
9. Keep blocked leaves unchecked. Record the concrete blocker, child leaf, and next step in Evidence.
10. Update the existing checkbox when finished. Do not create a second status list or mark a workstream complete.

Task IDs remain stable. Append split leaves with unused IDs and update dependencies. Do not reuse deleted IDs.

Primary paths are proposed write locations. Referenced tests and artifacts become real deliverables when the leaf runs.

Run the backlog check after changing dependencies, evidence, or checkboxes:

```bash
bun docs/engineering/sciobot-port/check.mjs
```

The check lists eligible leaves. It rejects dependency cycles, broken links, missing evidence, and premature closure.

## Completion rules

- Close the stated behavior with explicit expected results. Prefer captured Plate fixtures or native semantic fingerprints as independent oracles.
- Do not use implementation snapshots alone as proof. Do not weaken expected results to match a failing implementation.
- Required supported behavior cannot close through an unsupported-operation error. Failure checks only prove atomic refusal.
- Keep source projections read-only. Unsupported edits fail before writes and never disable tracking automatically.
- For document or geometry changes, use two real collaboration sessions through V01. Test the specific changed behavior.
- Check synchronization, concurrency, undo/redo, deletion during peer editing, reconnect, and save/reopen convergence.
- Check remote selection geometry when editable regions or layout change. Pure functions and fixture-only leaves need no session test.
- Keep collaboration evaluation dependencies outside the production path. Record the repository compatibility decision when shared operations change.
- Run scoped formatting, lint, typecheck, and relevant tests for code changes. Update API snapshots and adapter parity when public contracts change.
- Add the required changeset for publishable code changes. Keep engine lane boundaries and the single-core rule intact.
- Follow application migration instructions if a later leaf needs SQL. Do not create migration filenames manually.
- Put the command, result, revision, fixture or artifact path, and any limits in the leaf Evidence field.
- Complete the relevant coverage cells with evidence links. A checked proof or audit leaf does not itself establish production support.

Use this dispatch prompt:

```text
Read AGENTS.md, SCIOBOT_PORT_TODO.md, and leaf <ID>.
Verify dependencies and claim only <ID>.
Read its mapped sources. Implement its bounded change.
Run its acceptance checks and applicable collaboration checks.
Preserve unrelated changes. Record evidence beside the checkbox.
If blocked, leave it open and add the smallest missing prerequisite.
```

## Cutover gates

- V02 proves direct edits and accepted suggestions match all seven captured Plate cases.
- V03 proves mixed review waves preserve user changes and pending state through reload.
- V04 proves the combined features converge between two real sessions.
- V05 records large-document performance and leaves regressions open until fixed.
- V06 requires actual desktop Word files. Browser tests do not replace this interoperability check.
- V07 verifies the shipped runtime dependency boundary.
- V08 requires every other leaf and every production coverage row to pass.

The existing proof evidence is in [proof-results.json](examples/react-blocks/proof-results.json). Its checks remain separate from port completion.
