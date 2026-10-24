# Native inline removal checkpoint

Runtime: `72c1e85c`, branch `codex/d30-repeated-atom-anchors`.
D30 remains open. This checkpoint closes no leaf.

## Behavior

The private plan now carries `inline-remove` intentions for managed pictures, inline math, mentions, and sources.
Each intention carries the original control identity, complete payload, and native subtree fingerprint.
The executor checks these values before publication. Native protection still applies.

Planning removes addressed controls from an immutable native preview.
It adjusts surviving inline bindings and plans subsequent changes in the resulting coordinates.
Ordinary text never deletes a locked fallback.
Equal-cost alignments prefer fewer native removals.
Source property updates preserve their original controls instead of becoming deletion and insertion.
Self-closing source nodes survive removal of adjacent atoms.

Native control removal, payload retirement, and orphan image cleanup share one transaction.
Payload retirement retains an empty identity shell for shared history.
Shared execution requires complete owned review history before publication.
Known refusal codes survive the outer executor error handler when no write occurred.
No Pro runtime enters the private production graph.

## Verification

The new fixture contains 27 cases:

- Twelve local cases cover all four atom kinds in plain, linked, and code-linked paragraphs.
- Eight shared cases cover all four kinds with both client orderings.
- Three compound cases combine removals, text, births, image replacement, and retained source updates in paragraphs, cells, and registered containers.
- Two compound shared cases cover deletion and birth with independent peer typing and both histories.
- Two refusal cases check forged native sources and missing shared ownership before publication.

Assertions include complete Word fingerprints, surviving control identities, payloads, media bytes, empty replanning, undo/redo, and DOCX reopening.
Shared cases also check both actor histories and DOCX/room reopening.
The four-kind shared matrix checks caret geometry.
The alignment tests compare bounded streams against an independent edit-distance oracle, including identity tie breaks.

Recorded runs:

- `/tmp/d30-removal-integration.log`: 97 passed, two failed across seven files. C01 exposed source rebirth; the refusal test exposed error-code loss.
- `/tmp/d30-removal-final-source.log`: 45 passed, four failed across five files. Three failures exposed discarded surviving source nodes; one exceeded its unchanged five-second limit.
- `/tmp/d30-removal-preserve-empty-leaves.log`: all five focused cases passed after correction, including all three compound owners, C01 source content, and the timed-out image history case. No limits changed.
- `/tmp/d30-removal-complete-compiled.log`: all 27 original removal cases passed against built private entries, built provider, and built core; 943 assertions.
- The final change after that compiled run only narrows the surviving-children guard with `Array.isArray` for TypeScript. Accepted target AST behavior stays equivalent. No broad suite was repeated for this adjustment.
- `/tmp/d30-removal-closed-types.log`: 2,226 files passed.
- `/tmp/d30-removal-closed-build.log`: 273 private modules; native core stays external.
- Root formatting and lint passed. Final changed-source formatting, private lint, line caps, and diff checks passed.
- `/tmp/d30-removal-compatibility.log`: compatibility policy passed.

These runs provide composite evidence. They are not a claim that one final full source suite passed.
Temporary execution logs are diagnostic artifacts, not durable acceptance fixtures.
The committed fixtures retain the executable assertions.

## Remaining scope

Tracked atom execution and the remaining D30 integration checks stay open.
The complete C01/composer integration gate remains required before D30 closure.
The production coverage audit still refuses application revision drift at `073cb22447c6ac6a3699958e27bb1f3b8661bd22`.
No production source, audit baseline, or migration was changed.
