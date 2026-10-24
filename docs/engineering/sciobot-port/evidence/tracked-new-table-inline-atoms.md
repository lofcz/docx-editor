# Tracked inline atoms in newly inserted tables

D30 and R19 remain open. Work continues on `codex/d30-composer-reconcile`, after `5f1bed62`.

## Implementation

Tracked sibling tables now pass their cell-prefixed atom bindings and admitted image resources into native cell allocation. Each cell stages complete controls, payloads, valid Word identities, paragraph resources, and image reservations. The outer sibling stage finalizes only its own paragraph atoms. Table atoms finalize exactly once within their actual cell owners.

Private presence schema 6, with native-inline profile 4, owns a complete sibling gap. A new table can share its native receipt with following caret paragraphs, adjacent new paragraphs, other new tables, and removed source paragraphs. Paragraph-only and single-table cell proofs cannot represent that scope.

The binder derives roots from persistent revision addresses. It requires the atom's insertion wrapper and paragraph mark, exact attribution, bounded unique addresses, complete coupled native dependencies, one native story, contiguous roots, and insertion markers for every row of each new table. Existing tables with only a newly inserted row do not qualify as whole-table births. The binder rejects missing row births, unrelated roots, omitted dependencies, and downgraded profiles. Serialized proofs contain no transient canonical IDs.

The previous built private SDK refuses all four tested pending and retained decision snapshots. Existing legitimate presence profiles keep their previous behavior. A direction guard also prevents native birth proofs from binding removal intentions. Two negative cases reproduced the earlier acceptance of forged inverse records before the guard. Native transport format, journal kinds, and deterministic repair stay unchanged. Test transport evaluation uses the EigenPal Pro License. Production private runtime imports no Pro implementation.

## Oracles and verification

Eight local cases cover middle and terminal tables, mixed gaps with two tables and a separate paragraph atom, and replacement of an original paragraph. Both decisions compare complete native Word trees with independent direct execution. They also verify image bytes, relationships, payloads, both history boundaries, and save/reopen.

Twelve real-session cases cover both client orders and both decisions for middle, terminal, and mixed table gaps. They check offline peer text in retained content, both participants' histories, pending and resolved cold rooms, exact canonical convergence, complete native Word output, media, and payloads.

Six repeated-wave cases change all four atom kinds and verify both decisions and both creation histories. Guard checks refuse forged or downgraded scopes without package writes. The older-reader check uses the previous built SDK's public working-view reader.

Initial focused results: eight local cases and 418 assertions; twelve shared cases and 1,632 assertions; six repeated-wave cases plus two guard cases and 423 assertions. The private build verifies 287 source modules with core external. The initial private source type check passes across 2,288 files. Repository lint and compatibility policy pass with existing warnings.

## Limits

This checkpoint closes no TODO. Four previous committed-view regressions remain unresolved. Further registered cell fragments, peer edits inside these atom-bearing new tables, and complete application verification remain outstanding. The C01/composer boundary passes 574 cases and fails the four known committed-view expectations across 45 files and 25,638 assertions. No additional failure appears. This boundary precedes the final direction-refusal guard. The final affected batch passes all 48 cases across eight files and 4,165 assertions, including both direction refusals and the interleaved-root refusal. The compiled private SDK and evaluation transport pass all 26 new-table cases and 2,398 assertions. No original output expectation or timeout changes.

Logs use `/home/lofcz/.cache/sciobot-docx-native/d30-new-table-atoms-` with `composer-boundary`, `final-source`, `compiled`, `direction-before`, `prior-reader-final`, and `final-build` suffixes.

Final checks pass: private source types across 2,289 files, repository lint with existing warnings, formatting, line caps, and compatibility policy. The final private build keeps core external. Check logs use `final-types4`, `final-lint`, `final-policy`, and `final-build` suffixes.
