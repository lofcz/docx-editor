# Shared composer history repairs

Runtime `827f0820` repairs proposal history publication. Runtime `8d967ee7` repairs concurrent text reconciliation during structural replacement. Follow-up `9c8a08c8` preserves exact typing origins when a text edit also changes paragraph metadata. The production adapter keeps its import boundary; the changed transport remains evaluation-only under the EigenPal Pro License.

## Causes and resulting behavior

The namespace failure occurred during history, after a correct reconnect. A restored revision attribute could become visible before a queued normalization transaction delivered its namespace-binding event. The old receiver materialized that incomplete registry state, refused `invalid-qname`, and retained the previous proposal view. Publication now runs after Yjs finishes all synchronous transaction cleanup observers. Local journals still publish in their original commit frame; the receiver does not replay their absolute coordinates later. Destroy removes the replacement listener.

The text failure had a separate cause. Fresh typing origins had been enabled for every journal containing a text splice. Structural replacement journals that rebuilt text without explicit native source recipes then lost their established fragment reconciliation. Peer text survived in Original but disappeared from Proposed. Rebuilt text without a recipe retains that reconciliation. Explicit native source recipes retain exact mapping. Text edits without unannotated text births retain exact fresh origins, including frames that also write paragraph metadata.

A stronger regression combines peer typing with native paragraph spacing in one transaction while another actor edits atom hyperlinks. The complete original Word, control, payload, deletion, both-history, geometry and DOCX/room assertions remain. The existing twelve ordinary-typing cases stay intact. Twelve additional compound cases prevent a narrow typing-only fix from reintroducing the equal-space ownership bug.

## Focused verification

The first cleanup repair passes 84 of 96 cases across the five previously failing composer files. Its twelve remaining born-table/list failures identify the separate text-origin cause. After that repair, the focused matrix passes 36 cases and 2,776 assertions: all twelve born-table/list cases plus 24 atom-link histories across ordinary and compound typing. All original expected documents, timeouts and native oracles remain intact.

The broader collaboration gate passes 235 cases across 32 files, with 2,140 assertions. It covers queued journals, remote receive and reentrant windows, native trust/security, undo groups, typing history, selection, split-source behavior, lifecycle and receive costs. Existing typing gates pass; attached insert median is 0.526 ms. Four hundred edits leave zero pending journals, zero registry-node growth, and one unchanged awareness entry. Logs: `/tmp/composer-final-provider-gate.log` and `/tmp/composer-compound-typing-fixed.log`.

The final C01/composer gate at `9c8a08c8` passes **333 tests across 24 files**, with **16,071 assertions**, in 256.73 seconds. All 64 recorded failing test names pass. Every composer fixture remains byte-identical to checkpoint `14a60030`. No fixture or oracle was recaptured. The [structured repair evidence](composer-shared-history-repairs.json) retains the original failure names, exact commands, revisions and results.

The final compiled SDK/core/exporter atom matrix passes **103 cases and 11,304 assertions**. The original compiled code matrix passes **56 cases and 1,267 assertions**. These execute original fixture bodies with imports redirected to compiled packages and one compiled core. The atom matrix includes the twelve additional compound typing cases. No timeout or performance limit changes.

The earlier reader at `/tmp/docx-upstream-oct23-timing` preserves **72 exact native snapshots** across its own edits and both histories, before and after current-writer acceptance. The completed room state keeps the same shared interpretation.

Logs: `/tmp/composer-final-source-gate.log`, `/tmp/composer-final-compiled-atoms.log`, `/tmp/composer-final-compiled-code.log`, `/tmp/composer-final-prior-reader.log`.

```sh
mapfile -t composer_tests < <(rg --files packages/sciobot/src | rg '(target-c01-matrix|recompose).*test.ts$')
bun test --preload ./scripts/test/preload-dom.ts "${composer_tests[@]}"
bun packages/sciobot/scripts/check-code-inline-controls-compiled.mjs inline-images-materialization.test.ts inline-images-collaboration.test.ts imported-image-boundary-collaboration.test.ts retained-atom-links.test.ts retained-atom-links-collaboration.test.ts retained-atom-marks.test.ts retained-atom-marks-collaboration.test.ts
bun packages/sciobot/scripts/check-code-inline-controls-compiled.mjs
```

## Build and compatibility checks

Pro React types cover 2,310 files; Vue types cover 1,861. Private types cover 2,203 files. Core, Vue, Pro and private builds pass with local sibling declarations. The private build verifies 261 source modules with native core external. Pro API checks pass without snapshot changes. No dependency version, lockfile, public document member, shared format field or primitive shape changes.

Root formatting ran for the cleanup repair. Final changed files also pass scoped formatting. Root lint and the repository compatibility policy pass; the policy checks 190 decisions and two inherited format fields. The two patch changesets record the published transport fixes. Native source recipes and ordinary typing offsets keep their existing representation; earlier readers can read the same completed room state. Their existing authoring bugs do not define a new format.

## Remaining port work

D30 still has two explicit failing retained/fresh linked-atom planning probes. Image removal/replacement, nested registered cell fragments, tracked atom ownership and installed application parity remain open. R19 retains its remaining insertion, reconciliation, resource and application requirements. Neither checkbox closes from these regression repairs.

The production audit still rejects application revision drift. All 199 pinned source hashes match; the baseline remains intact. The backlog remains 276 leaves: 179 closed and 97 open.
