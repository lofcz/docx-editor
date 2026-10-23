# D30 table birth history investigation

This work remains isolated on `codex/d30-born-cell-history` in `/tmp/sciobot-resolved-table-targets`.
Runtime `db8e74e7` follows the work-in-progress checkpoints `015d5bea`, `a9a00a00`, `2a915a72`, and `bd2727de`.
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

## Remaining work

The target planner still refuses an unequal header split into spans 2 and 1.
The stronger history test uses the actual supported native split command and verifies its exact resulting four-column grid.
That command does not establish target-planner support for the unequal split.
More combined topology, vertical merge, and newly allocated span-property cases remain unverified.
Concurrent independent geometry frames still need combined native topology verification.
Move destinations against hidden child lists remain unverified.
Do not merge or release this checkpoint as a completed history repair.

No TODO closes. D30 also retains its existing atom ownership, nested content and application integration requirements.
