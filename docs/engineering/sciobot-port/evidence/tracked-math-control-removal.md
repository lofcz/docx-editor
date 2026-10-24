# Native equations in deleted inline controls

D30 and R19 remain open. This checkpoint removes the native equation refusal for whole inline control deletion.

The canonical writer keeps each complete `m:oMath` or `m:oMathPara` subtree unchanged inside the control's deletion. Word text beside the equation still becomes deleted text. Nested revisions, locks, attribution, and bounded traversal retain their existing checks. Equation-only controls now classify as inline content. Equation layout retains its enclosing hyperlink, so projection no longer splits that link around the equation.

The schema permits these structures through `CT_RunTrackChange`, `EG_ContentRunContent`, and `CT_SdtContentRun` in `reference/ecma-376/part1/schemas/wml.xsd`.

Verification retains independent expected native outcomes:

- The source integration batch passes 98 cases across eight files with 1967 assertions. It includes C01, AI recomposition, layout, native deletion, and real shared sessions.
- The final native fixture passes 17 cases with 146 assertions. It covers equation-only and mixed controls, range and whole-control commands, hyperlinks, both decisions, unknown metadata, immutable equations, rejection identities, and XML reopening. Nested revision ownership still refuses.
- The compiled core/private SDK/evaluation transport passes all 59 unchanged fixture bodies with 1308 assertions. Shared cases cover both client orders, offline peer typing, both actors' histories, decision history, caret geometry, convergence, and DOCX/room reopening. AI cases cover locked and unlocked controls, committed and proposed views, payload retirement, exact decisions, and history.
- Four control regression files pass 214 cases with 938 assertions.
- The core build passes. Private source types pass 2254 files. Core API checking passes 15 entries with zero errors, 2489 warnings, and 391 forgotten exports. No snapshots or audit baselines change.
- Scoped lint, max-lines caps, formatting, and whitespace checks pass. Collaboration policy passes 208 decisions with two cumulative format fields changed. This checkpoint adds no shared format field or materialization rule.

Logs reside in `/home/lofcz/.cache/sciobot-docx-native/`: `d30-math-removal-batch.log`, `d30-math-goldens-final.log`, `d30-math-removal-compiled.log`, `d30-math-control-regressions.log`, `d30-math-core-build.log`, `d30-math-types-final.log`, `d30-math-api.log`, and `d30-math-policy.log`.

The implementation uses Apache core operations. Test-only transport uses the EigenPal Pro License. Production adapter modules gain no Pro import. The application remains read-only.

Remaining D30 work includes tracked images and resource ownership, new-cell atom ownership, arbitrary retained atom hyperlink placement, and complete integration gates. No TODO closes from this repair alone.
