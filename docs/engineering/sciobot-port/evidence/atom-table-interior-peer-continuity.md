# Interior peer text through atom-table rewrites

D30 and R19 remain open. This checkpoint follows `88eb245e` on `codex/d30-composer-reconcile`.

## Failure and repair

A disconnected participant types between native inline atoms. Another participant replaces the owned table and changes every atom. Continuing the paragraph address alone keeps the peer text in the cell, but loses its interior position.

The private coordinator now inventories ordered atom kinds and their physical text gaps. It continues text identities only when the complete ordered atom inventory matches, all source atoms belong to the selected proposals, and each gap has the same complete native text elements. Position establishes correspondence. Fingerprint equality verifies unchanged content; equal labels grant no identity. The target keeps its new runs, formatting, revision wrappers, controls, payloads, and media. Existing source identities cannot overwrite live target nodes. Unknown structures and changed gaps retain their existing behavior.

Three native behaviors also need correction:

- Diff cleanup must retain a complete unchanged native text element between replaced elements. Otherwise it deletes the spaces that anchor independent peer typing.
- A whole inline control replacement uses a boundary inside its removed content. Concurrent text before the old control stays before its replacement. Text after a multi-character control stays after it. Ordinary text replacement and edits within the same control retain their existing boundary.
- Terminal tracked append projection must distinguish concurrent content from content already observed by a participant. A right-hand native shell proves existing content. A deleted clock cannot prove an append boundary. The reader stops at every deleted clock because collection removes its content type. The reader must preserve those authored positions.

These native changes use format `docx-collaboration:1.7.12.1`. Shared schema 7 and journal kinds stay unchanged. The previous format refuses admission without writes. Deployment requires verified export with the previous compatible client and fresh room seeding. No saved room is migrated. Transport changes remain evaluation-only under the EigenPal Pro License.

## Independent checks

The interior regression covers the math/image, image/mention, and mention/source gaps. It uses both client orders and both review decisions. Every case changes all four atoms while the peer remains disconnected. Replanning the intended complete Markdown must return no operations after reconnect.

The tests retain exact Word fingerprints, separate authorship runs, independent accepted and rejected documents, both creation frames, both participants' histories, exact live convergence, media bytes, image relationships, payload membership, saved Word reopening, and cold room reopening. Remote carets use the end of the peer text, including its interior position. Original timeouts remain unchanged.

Native tests require complete native gap identities, refuse identity from equal newly created text, and distinguish causal from concurrent terminal inserts. The causal tests use live replicas and cold reopening with garbage collection enabled and disabled.

## Verification

The complete C01/composer batch passes 585 cases and retains the four previously recorded committed-view failures. It runs 589 cases across 45 files and 28,778 assertions. This batch precedes the final deleted-clock guard.

The final source batch passes all 80 cases across ten files and 4,519 assertions. It includes every interior placement, the original suffix rewrites, original plain-table peer rewrites, prior cell histories, previous-format refusal, constructor field ownership, and native typing alignment. The collected-formatting regression first fails, then passes after the guard stops at every deleted clock. Live and cold replicas now choose the same ordering.

The actual built private SDK, core, and evaluation transport pass all 28 peer/cell cases across four files and 3,692 assertions. Native history, property routing, journal order, resource limits, and receive costs pass all 34 cases across eight files and 133 assertions. Existing cost limits remain unchanged.

Private types pass across 2,295 files. Pro types pass across 2,332 files. Builds retain 289 private modules with core external and 115 evaluation transport modules. Compatibility policy passes with 222 decisions and two changed format fields. Root lint and line caps pass. No original expectation or timeout changes.

The final workspace format, lint, parity, license-header, and compatibility gates pass. Whole-workspace types fail in ten examples and benchmarks with missing dependencies, including `yjs`, Hocuspocus, OpenTelemetry, and Node types. The affected private and Pro projects pass their source type checks, including Pro's Vue program in the workspace run. No dependency installation or shared dependency change is attempted.

The API snapshot check remains incomplete. Font and PDF declaration builds pass after the first missing-artifact refusal. The next check refuses missing React declarations; Vue, Pro, and editor-api declarations are also absent in this isolated checkout. No API snapshot or forgotten-export baseline is recaptured. The native runtime and compiled private checks above do not establish a complete public API gate.

Logs use `/home/lofcz/.cache/sciobot-docx-native/d30-cell-interior-` with `composer-final`, `source-final`, `sdk-types-final2`, `pro-types-final2`, `sdk-build`, `pro-build-final`, `compiled-final`, `native-cost-final`, `lint`, and `policy` suffixes. The first collected-formatting failure uses `d30-cell-gap-collected-proof.log`.

## Remaining scope

This checkpoint closes no TODO. Topology-changing replacements, nonempty or independently formatted recovery slots, remaining registered cell fragments, and application integration remain outstanding. Four committed-view expectations remain unresolved.

A separate native revision test also fails with all three changed native modules restored to `88eb245e`. `document-replacement-times.test.ts`, case `deleting a revision while another participant edits its text converges`, expects the independently replaced text to retain its insertion review item. The current reader drops that item. This baseline failure remains open; the interior fix does not resolve it.

The production inventory audit refuses application revision drift. All 199 pinned source hashes still match. The pinned application revision is `5650bdf550b7987174d1b61fe2a925b86fb24f15`; the observed revision is `997e989f59270930c9af761547946ec14b1d0cac`. No baseline is recaptured, and the application stays unchanged by this work.
