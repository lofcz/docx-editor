# Tracked table property holders

Runtime `f7ca8c5e` repairs the two inherited native acceptance failures in `table-authoring.test.ts`. The direct native commands retain empty row and cell property holders. Tracked creation records that presence on its revision marks. Column property history also records when a source cell originally lacked a holder. Accept and reject now preserve the exact independent native tree after DOCX reopening.

The writer extends values of the existing optional OOXML property-presence annotation. Legacy and unknown values keep their existing interpretation. No shared directory, message type, or character representation changes. Namespace bindings disappear after resolution. The compatibility decision and patch changeset accompany the implementation.

## Verification

- `bun test packages/core/src/store/__tests__/inserted-table-property-presence.test.ts packages/core/src/automation/__tests__/table-authoring.test.ts`: 40 cases, 153 assertions. The original two acceptance expectations remain unchanged. The new matrix covers rows, columns, missing/empty/formatted properties, populated rows, old marker address families, and imported unknown metadata. Acceptance uses the direct native oracle; rejection uses the source oracle, including full package fingerprints and save/reopen.
- The affected native and evaluation transport check passes 410 cases across 15 files, with 2,498 assertions.
- The complete core automation suite passes all 657 cases across 42 files, with 2,570 assertions. Its original performance limits remain unchanged. Log: `/tmp/d30-table-holder-automation-wide.log`, frozen revision `f7ca8c5e`.
- Core type checks pass across 1,601 files. Private SDK type checks pass across 2,187 files.
- `bun --preload ./scripts/test/preload-dom.ts packages/sciobot/scripts/check-table-holder-reader-compatibility.ts /tmp/docx-upstream-oct23-timing`: 72 exact native snapshots. The older reader/writer at `15b9dcfe` edits existing cells while proposals remain pending. It then edits new cells after current-writer acceptance. Both older peers undo and redo their text. Both current peers match each independent native oracle. Current-writer acceptance undo and redo retain the foreign text in the new cells. Both row and column topologies pass. Log: `/tmp/d30-table-holder-reader-history.log`.

The older writer still has its original cleanup bug when it resolves new proposals. This check establishes preservation and ordinary editing across readers; it does not claim that old resolution code gains the fix.

The EigenPal Pro License transport remains test-only. Production uses Apache core and the private adapter. D30 remains open. No TODO closes at this checkpoint.
