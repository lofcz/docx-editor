# Mixed retained and fresh inline atoms

Runtime `7234d5a0` plans fresh atoms beside retained native controls. Follow-up `7ea0a3d0` fixes formatting inheritance.

The planner resolves retained controls through their original source bindings and native UTF-16 ranges. It diffs text gaps independently. Retained controls become literal spans only in the temporary planning AST. Anchored text operations cannot replace those spans. The existing `inline-atoms` codec describes only fresh controls and image resources. Execution preserves retained native controls, payloads, and media.

Named hyperlink insertion now places text beside protected child controls within the same hyperlink. Interior protection and restrictions around the hyperlink still refuse writes. Application honors the resolved sibling index and inherits the visible run properties. New runs exclude pending formatting records. A native regression first found two records where only the original peer record should exist; it now passes.

## Verification

The [structured results](mixed-inline-atoms.json) record commands, revisions, fixture hashes, and results.

- Both original [linked retained/new atom probes](regressions/linked-retained-and-new-atoms.test.ts) pass without assertion changes.
- Forty-eight local cases cover math, image, mention, and source retention. Fresh atoms appear before and after retained atoms. Plain, linked, and code-linked content includes repeated-text gap edits. Cases execute serialized plans, preserve native control fingerprints and payloads, and verify exact undo, redo, media, and reopening.
- Thirty-two shared cases cover both client orders and both inserting actors. Concurrent native peer typing supplies an independent Word oracle. Both actor histories preserve full native Word, registered payloads, original controls, and media. Remote selection geometry and DOCX/room reopening pass.
- The compiled SDK/core/exporter matrix passes 239 cases across twelve unchanged test bodies, including the eighty mixed-atom cases.
- Native control, link, wrapper, and typing checks pass 1,028 cases across 33 files.
- The final C01/composer gate passes all 333 cases with 16,071 assertions. Its 24 original fixtures remain byte-identical to `4a0435ad`.
- Core and private types pass 1,605 and 2,206 files. Both builds pass. The private build verifies 262 modules with native core external. Core API checks report fifteen entries and zero errors. Root formatting and lint pass; the final inheritance follow-up also receives scoped formatting. Collaboration compatibility remains unchanged.

## Remaining work

D30 stays open. Identical new atoms with ambiguous retained identity, image removal/replacement, nested registered cell fragments, tracked atom ownership, and installed application checks remain unproven. Existing R19 requirements also remain open. This checkpoint closes no TODO leaves.

The stock production inventory and backlog checks still reject application revision drift. All 199 pinned source hashes match. No application source or inventory baseline changes occurred.
