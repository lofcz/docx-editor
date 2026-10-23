# D30 table birth history investigation

This work remains isolated on `codex/d30-born-cell-history` in `/tmp/sciobot-resolved-table-targets`.
Runtime `20eb0d7c` follows the work-in-progress checkpoints `015d5bea`, `a9a00a00`, `2a915a72`, `bd2727de`, `db8e74e7`, `d697de99`, `d3889299`, and `37b117d9`.
The primary branch remains at `75930cb4`. D30 remains open.

The evaluated transport uses the EigenPal Pro License. These changes add no private production imports.
They preserve native table ancestry through creation undo and hide empty retained scaffolds during projection.
Undo markers belong to the undo transaction, so redo reverses them with the native history frame.
Hidden children require journal splices to address observed identities instead of raw array indexes.
The repair retains only actual table, row and cell births and their dependent geometry.
Paragraph-only edits and native table resizing keep their existing history behavior.

## Verified checkpoint

- The focused row/column history matrix passes all 16 cases, including disconnected peer edits after creation undo.
- Those cases verify native cell placement, original cell fingerprints, both histories, remote caret geometry, and DOCX/room reopening.
- The original merged-header resize regression passes with the scoped retention hook.
- All 109 affected cases pass across 11 files, with 3275 assertions, in 142.51 seconds.
- The main Pro type check passes across 2286 files.
- Root lint passes with existing warnings.
- Root formatting passes. Collaboration compatibility policy passes with 169 decisions and one changed format field.
- Both original peer-text probes pass unchanged when run alone: 2 cases, 16 assertions, 2.71 seconds.

The earlier combined probe run reported one original probe timeout at its unchanged 5000 ms limit.
That run overlapped the affected suite. Its other failure was the merged-header span assertion.
The isolated rerun keeps the original deadline and all original assertions.

The affected command includes `born-cell-history-collaboration`, `inserted-cell-atoms-collaboration`,
`registry-undo-tombstone`, `document-format`, `document-table-merge-split`, `document-table-vmerge-rows`,
`document-table-span-resize`, `document-table-row-growth`, `document-concurrent-relocation`,
`document-typing-history`, and `document-undo-groups` tests.

## Geometry repair

Run `bun test ./docs/engineering/sciobot-port/evidence/regressions/born-cell-geometry-and-journal.test.ts` explicitly.
The merged-header case still expects a native span of 3 after a peer preserves the inserted middle column.
It now passes with that exact expectation unchanged.
The second case verifies that a later native row deletion preserves a hidden birth awaiting disconnected peer text.
That case passes with identity-based journal splices.

Runtime `2a915a72` captures exact native cell-to-column coverage before and after admitted table births.
It stores only native identities and coverage, with no paragraph content or canonical package snapshot.
The creating actor's undo retains this immutable evidence on the born cell record.
Projection derives spans from surviving column identities and new cell ownership.
This retains widened merged headers and shrinks original spans beside surviving split cells.
Later native span writes record their own ownership in the same journal transaction.
Their undo and redo restore the corresponding native values and projection ownership.
Capture reads written child lists inside the transaction because derived parent indexes have not updated yet.

Four stronger cases verify column births and cell splits, with connected and disconnected peer text.
They compare exact native topology and complete Word fingerprints through both actors' histories.
Both DOCX and room reopening match those fingerprints without writing shared state during reads.
The column cases apply a second native split, then verify its native spans, cold reads, undo and redo.
The six focused cases pass with 164 assertions at `2a915a72`.
All 113 affected cases pass across 12 files, with 3421 assertions, in 36.67 seconds.
Types cover 2287 files. Root lint and collaboration policy pass, with two changed format fields.

Runtime `bd2727de` stores topology as bounded structured objects instead of ordinary XML strings.
The ordinary 4096-character XML value limit cannot represent larger table histories.
The disconnected column case now includes ten rows and verifies all empty companion cells and later native split spans.
All four strengthened geometry cases pass, with 146 assertions, in 3.98 seconds.
Final checkpoint verification passes all 24 history and original regression cases across four files.
It includes 1180 assertions and completes in 30.27 seconds.
The final main Pro type check covers 2287 files; root formatting and root lint pass.

New immutable geometry evidence and native-write ownership require shared schema 6.
Together with repair version 3, the proposed format is `docx-collaboration:1.6.3.1`.
The changeset, compatibility decision, and migration guide require export and reseed.
No migration or deployment has occurred.

## Geometry survives deletion of its first row

Two new real-session cases expose a cold-read failure at `9b001c54`, with both client orders.
Alice creates a middle column below a merged header. Bob types inside two new cells in different rows.
Alice undoes the column birth, then deletes the first body row.
The live cached header still has span 3, but room reopening shows span 2.
The immutable frame lived on the first born cell, so row deletion hid it from frame discovery.
Both original cold-read assertions fail. Neither assertion changes during the repair.

Runtime `db8e74e7` places frames in an immutable directory on the retained native table record.
The first born cell identity remains the frame key, but its placement no longer owns frame visibility.
The undo delete filter retains both the directory and its immutable entries.
Native undo markers continue to determine when a retained frame contributes geometry.
Later native span writes also read the table-owned directory.
Frame traversal follows causal dependencies from native cell identities, with deterministic key ordering for independent frames.
It does not depend on Y.Map receipt order.

All six geometry cases pass with 212 assertions in 8.72 seconds.
The new cases verify explicit native topology, complete Word fingerprints, DOCX/room reopening, deletion undo/redo, and both peer histories.
They also apply and undo a later native split after the first row disappears.
The main Pro type check passes across 2287 files.
The private SDK type check passes across 2165 files.
Its older history oracle now narrows text nodes before reading `localName`; expected values and assertions remain unchanged.
All 115 affected cases pass across 12 files, with 3487 assertions, in 47.55 seconds.
The four original investigation cases pass unchanged, with 34 assertions, in 3.16 seconds.
Root formatting, root lint and collaboration compatibility policy pass.
The proposed schema remains the unreleased `docx-collaboration:1.6.3.1`; no deployed room changes.

## Native moves after creation undo

Two native move cases fail at `6b77e235`: inserting an existing row into the middle and composing two moves in one journal.
An append case passes at that checkpoint.
The failure comes from applying visible destination indexes directly to raw shared lists containing hidden births.

Runtime `d697de99` names visible destination neighbors during journal preflight.
Execution detaches the moving node before resolving those neighbors against the raw destination list.
Preflight also retains the source each journal prefix sees.
This removes an earlier placement when the same row moves across parents twice in one transaction.
Existing registry parent indexes update after that transaction and cannot name that intermediate source alone.
Moves into text records refuse before any shared write.

Four cases use real row creation, creation undo, and disconnected peer text in the new row.
A fresh room replica then applies native primitive journals against the visible Word tree.
The cases cover middle insertion, append clamping, two moves in one parent, and moves through another table.
An independent expected tree retains all native nodes and changes only observed row order.
Late peer text remains in its retained row after those moves.
Cold room readers use their own snapshot blob maps; DOCX readers match complete native Word fingerprints.
The initial four cases pass with 52 assertions in 4.59 seconds.
Main Pro types cover 2287 files; private SDK types cover 2167 files.
Root format, lint, and compatibility policy pass at the runtime checkpoint.

The wider 17-file check exposes four unchanged projected-journal tests with a read-contract failure.
Table birth visibility had required shared maps even for unrelated paragraph and split-run shape reads.
Runtime `d3889299` restricts that repair to actual native table ancestry, using existing kind and parent reads.
The original projection fixtures and all expected results remain unchanged.

A new bounded-room case also exposes a missing resource refusal.
Six raw table children include one hidden birth; moving another row into the table must exceed the limit of six.
Visible child counts alone incorrectly admit that move.
Projection now tracks shared child occupancy independently of visible indexes throughout the journal.
The visible split projection delegates occupancy to its raw shared projection.
The bounded move refuses with `too-many-children`, and complete shared-state bytes remain unchanged.

Move tests also compare complete Word fingerprints after undo and verify that cold reads do not write shared state.
The focused projection, move and merged-geometry check passes 28 cases with 405 assertions in 7.20 seconds.
Both package type checks pass: 2287 Pro files and 2167 private SDK files.
Final affected verification passes 158 cases across 17 files, with 3751 assertions, in 37.76 seconds.
It includes all direct move-effect journal, projection, repair and convergence fixtures.
Root formatting and root lint pass at the final runtime checkpoint.
The four original investigation cases pass unchanged, with 34 assertions, in 2.28 seconds.
Final compatibility policy passes with 169 decisions and two changed format fields.

## Concurrent native column coverage

Four new real-session cases fail at `f9c484b9`, with both client orders and reconnect receivers.
Alice and Bob each insert a middle column while disconnected, below the same merged header.
Both columns survive in the native grid, but the header span remains 3 instead of 4.
The merged native header also holds duplicate fixed-width properties.

The first repair reveals two further causes.
Lazy creation replaces the table's geometry map when two disconnected replicas each create it.
Tables now allocate that shared map during native record creation, before writers can disconnect.
Grid-span records also allocate the later-write ownership map at creation.
Native topology capture must read current shared attribute values inside the journal transaction.
The derived attribute index still describes the preceding state until that transaction finishes.
The before/after coverage reader now distinguishes those written shared values from pure projected reads.

Runtime `37b117d9` combines independent coverage by native column identity.
Causal dependency ordering includes new cell and grid-column identities from the source state.
Later observed edits replace prior coverage; independent edits retain each contributor's columns.
Surviving born cells reserve their recorded columns before original cells derive their spans.
Concurrent merged headers derive fixed widths from the surviving native grid widths.
Only one peer-created width property shows; imported duplicate properties remain unchanged when no peer-created copy exists.
Projection restores a required span value even when author undo removed its raw attribute.
Dependency invalidation compares both span and width projections, including transitions back to raw values.

The four cases verify complete header fingerprints, exact native cell identities, both authored atom fingerprints, and untouched original body cells.
Bob types inside Alice's new cell before Alice undoes.
Both histories preserve that foreign content, remove unused columns, and restore exact authored Word fingerprints on redo.
DOCX and room reopening match the full recovered Word fingerprint without shared-state writes during reads.
All ten concurrent and existing geometry cases pass, with 528 assertions, in 50.97 seconds.
Pro types cover 2288 files. Private SDK types cover 2169 files.
Root formatting and lint pass. Compatibility policy passes with 169 decisions and two changed format fields.
The proposed format remains unreleased `docx-collaboration:1.6.3.1`; the primary branch has not adopted this repair.

The first expanded 18-file run passes 161 cases and exposes one original native merge regression.
Active merge geometry must use its after coverage; undo repair must use before coverage plus surviving peer-owned columns.
Runtime `20eb0d7c` distinguishes those states before joining independent coverage.
The original `typing in a cell a peer merges away keeps the typing in the merged cell` expectation remains unchanged.
All 22 focused concurrent, existing geometry and native merge/split cases pass with 582 assertions in 25.14 seconds.
Final affected verification passes all 162 cases across 18 files, with 4067 assertions, in 107.18 seconds.
The final Pro type check covers 2288 files and root lint passes.
The four original regression probes pass unchanged, with 34 assertions, in 20.18 seconds.

## Remaining work

The target planner still refuses an unequal header split into spans 2 and 1.
The stronger history test uses the actual supported native split command and verifies its exact resulting four-column grid.
That command does not establish target-planner support for the unequal split.
More combined topology, vertical merge, and newly allocated span-property cases remain unverified.
Independent split and column edits, concurrent geometry-only writes, and vertical merge coverage still need verification.
Further combined move, topology, and concurrent geometry interactions remain unverified.
Do not merge or release this checkpoint as a completed history repair.

No TODO closes. D30 also retains its existing atom ownership, nested content and application integration requirements.
