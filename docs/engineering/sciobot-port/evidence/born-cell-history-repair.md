# D30 table birth history investigation

This work remains isolated on `codex/d30-born-cell-history` in `/tmp/sciobot-resolved-table-targets`.
Runtime `a9a00a00` follows the work-in-progress checkpoint `015d5bea`.
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

## Remaining failure

Run `bun test ./docs/engineering/sciobot-port/evidence/regressions/born-cell-geometry-and-journal.test.ts` explicitly.
The merged-header case still expects a native span of 3 after a peer preserves the inserted middle column.
The retained grid and body have three columns, but undo restores the original header span of 2.
The exact expected span remains unchanged.
The second case verifies that a later native row deletion preserves a hidden birth awaiting disconnected peer text.
That case passes with identity-based journal splices.

The repair needs exact before/after geometry provenance for retained cells and columns.
Numeric span offsets cannot establish this for split cells, merged cells, or combined topology edits.
Geometry projection must also support subsequent native journal writes and cold reopening.
Move destinations against hidden child lists remain unverified.
Do not merge or release this checkpoint as a completed history repair.

No TODO closes. D30 also retains its existing atom ownership, nested content and application integration requirements.
