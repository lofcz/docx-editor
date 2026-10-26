# Styled peer continuation during AI replacement

The private composer now retains unchanged styled peer tails during replacement of owned atom-bearing tables. It preserves their physical paragraph, original run identities, native properties, and text nodes. Ordered native atom slots and exact source identities establish continuity. Complete text and native property spans verify the requested target. Equal adjacent style spans can cross text-node boundaries without coalescing the retained source runs.

The initial regression reached an equivalent working target but created another replacement when that target was repeated. The continuation path admitted only plain runs. The no-op check also disabled native equivalence whenever application proposals were pending.

The fix addresses both paths. Native equivalence now reads a detached package with the same proposed properties and presence as the working projection. It never installs that package in the editor. Pending canonical payloads retain their committed values, so those values alone cannot establish working-target equality. Returning a pending property to its committed value still removes its proposal. An equivalent proposed numeric spelling now returns a no-op with the original package reference and revision.

## Regression coverage

Eight new shared cases cover fully bold and partly bold peer text, both client orders, and acceptance/rejection. A peer types and formats text inside the first AI wave. The next wave changes all four native atoms while retaining that styled text. Assertions cover:

- Exact native decisions against independently constructed Word XML and a separate direct-target document.
- Original view retaining the committed styled peer text.
- Both AI creation frames and the peer's separate typing and formatting history frames.
- Repeating an equivalent target without a new package, revision, or history frame.
- Stable paragraph identity, remote caret geometry, native convergence, media bytes, relationships, and payload membership.
- Saved Word output and cold shared-room reopening.

No original plain or interior regression is removed. Test deadlines and native fingerprint assertions stay unchanged. The new style cases require the existing per-character property spans to match; this change does not reinterpret requested restyling as unchanged formatting.

## Source verification

From the fork root:

```sh
bun test --timeout 60000 --preload ./scripts/test/preload-dom.ts packages/sciobot/src/review/recompose-styled-cell-peer.test.ts packages/sciobot/src/review/recompose-new-table-atom-peer-online.test.ts packages/sciobot/src/review/recompose-new-table-atom-peer-interior.test.ts packages/sciobot/src/review/recompose-payloads.test.ts packages/sciobot/src/review/recompose-payloads-collaboration.test.ts packages/sciobot/src/review/recompose-inline-images-shared.test.ts packages/sciobot/src/review/recompose-atomic-targets.test.ts
```

All 47 cases pass with 4,675 assertions across seven files. The log is `/tmp/r19-styled-final-batch.log`. The initial failure is `/tmp/r19-styled-peer-before.log`. Private types pass across 2,429 files. Scoped formatting and lint pass after the final formatting correction. Logs use `/tmp/r19-styled-types-final.log`, `/tmp/r19-styled-format-fixed.log`, and `/tmp/r19-styled-lint.log`.

The historical native revision-deletion failure also passes unchanged on the current runtime: one case and six assertions in `/tmp/r19-revision-deletion-current.log`. This does not establish broader R19 completion.

## Compatibility and remaining scope

The change uses existing canonical operations, property/presence projections, payload schemas, and ownership protocols. Native reader semantics and format `docx-collaboration:1.9.15.1` stay unchanged. The private SDK has an empty Changeset. The compatibility decision is `.collaboration/changes/sciobot-styled-peer-continuation.json`.

The evaluated transport remains test-only under the EigenPal Pro License. Production changes import Apache core and private modules only. R19 remains open for its remaining ownership routes and application verification. This checkpoint covers styled committed tails with unchanged native formatting. It does not establish general topology-changing or interior styled recovery.

## Installed runtime and application ownership

Immutable runtime `38d620acce8aec9678de93ea5871ba4e10dc8f1b` is installed by application `01ef1c5e62574cfbba935c6e373c71d545163aa7`. The private archive is `sciobot-docx-4772cd67c019.tgz`. All five archives match `SOURCE.json` and the package manifest. App, SDK, and exporter resolve one core runtime. Frozen installation passes. React, Vue, and evaluated Pro builds pass for compiled verification.

The built styled-peer file passes all eight cases and 1,288 assertions. The original linked-image shared file also passes all eight cases. The first combined compiled run hit three unchanged 30-second styled-case timeouts while machine load exceeded 90. The complete styled file then passes unchanged under the original deadlines. Logs are `/tmp/r19-styled-compiled.log` and `/tmp/r19-styled-compiled-verified.log`. No assertion or timeout is relaxed.

Application inspection found that suggestion executors were recreated for every tool call with an empty ownership ledger. The application now keeps the exact returned ledger and content proofs in a WeakMap keyed by the current workspace session. Successful calls update that state. Failed calls retain it. Resolved groups stay available for undo. A new mount epoch receives separate state. Same-author foreign proposals never grant ownership.

Two new real-browser cases pass both client orders with 62 assertions. Four successive suggestion calls, peer input, failed replacement counts, native undo/redo, and a later branch after undo preserve exact working content. Earlier owned group identities disappear from the active review list. Unrelated same-author group identities remain. Native text revisions stay unnested, and both saved documents converge. Different text diffs legitimately create different numbers of review groups, so the tests check exact identity withdrawal and content rather than assuming one group per request.

The original sixteen direct/read/bridge browser cases pass in the combined run. The two new cases pass in their corrected identity-based file selection. The composite browser evidence covers eighteen distinct cases. Eleven executor unit cases also pass. Frontend types pass across 6,840 files; harness types pass across 2,414 files. Scoped format/lint and diff checks pass. Logs are `/tmp/r19-app-tool-waves.log`, `/tmp/r19-app-ownership-final.log`, `/tmp/r19-app-edit-units.log`, `/tmp/r19-app-types-final.log`, `/tmp/r19-app-harness-final.log`, and `/tmp/r19-app-lint-final.log`.

R19 and I04 remain open. This session-local ledger does not restore host ownership after a new document mount. Remaining native recovery routes and declared integration dependencies still need completion.
