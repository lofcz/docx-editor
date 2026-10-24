# Replacement cell coverage during repeated review waves

D30 and R19 remain open. Worktree: `codex/d30-composer-reconcile`, based on `2c47facc` and helper extraction `56584a07`.

## Cause and repair

A second AI review wave removes its earlier split and creates a replacement cell. The transaction starts and ends with the same surviving cell spans. The geometry recorder therefore omitted a new frame. Undo retained the replacement shell for possible foreign content. Without coverage evidence, the reader treated its active column as sufficient to show the empty cell. The table gained a third cell.

The recorder now captures a new cell's coverage when it replaces an occupied column in the same native row. It uses existing native row, cell, and column identities. It still omits geometry for independent row births. Undo can hide the empty replacement while retaining independent participant content. No native tree oracle or history assertion changes.

## Evidence

All eight unchanged mixed-split regressions pass, with 688 assertions. Both decisions, both client orders, both reconnect receivers, both histories, caret geometry, and Word and room reopening remain covered. Log: `/home/lofcz/.cache/sciobot-docx-native/d30-split-coverage-repair.log`.

The new format is `docx-collaboration:1.7.8.1`. The schema remains version 7; deterministic repair changes from 7 to 8. Previous-profile admission refuses without shared writes. Upgrade instructions require verified old-client export and fresh room seeding. No saved room is migrated. The transport remains evaluation-only under the EigenPal Pro License.

The broader source batch passes all 147 cases and 7,606 assertions across seven files. The compiled SDK, core, and evaluation transport pass all 58 shared table and late-image cases with 5,110 assertions. Logs: `/home/lofcz/.cache/sciobot-docx-native/d30-replacement-coverage-{batch,compiled}.log`.

The final implementation indexes occupied columns by native row. It avoids scanning all previous cells for every new cell. The rebuilt evaluation transport passes all ten focused compiled regressions and 830 assertions after this refinement. The full test bodies and their original assertions remain unchanged. The compiled harness supports an optional `--test-name-pattern=REGEX` so focused follow-ups do not repeat the complete shared table matrix.

Pro source typechecking passes 2,328 files. Scoped lint, formatting, license headers, line caps, and compatibility policy pass. Logs use the same prefix, with `linear-compiled`, `linear-types`, `linear-lint`, `format-check`, `license`, `caps`, and `policy` suffixes.

Four prior committed-view regressions remain unresolved. No TODO closes at this checkpoint.
