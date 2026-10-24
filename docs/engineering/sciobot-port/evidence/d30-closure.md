# D30 acceptance audit

This audit covers D30's written requirements. R19, application integration, and cutover remain separate open leaves.

## Requirements and evidence

| Requirement | Current evidence |
| --- | --- |
| Explicit image, math, mention, and source operations in new paragraphs | `diff/inserted-atoms.test.ts` checks all four deferred operation kinds, exact payloads, asset references, offsets, and plan serialization. |
| Explicit operations in retained paragraphs | `diff/retained-atoms.test.ts` and `diff/retained-atom-identities.test.ts` check ordinary paragraphs, headings, cells, native identities, marks, and links. Image coverage also uses `diff/inline-images-materialization.test.ts`. |
| Materialization reaches target Markdown through save and reopen | `diff/inserted-atoms-materialization.test.ts`, retained-atom fixtures, and image fixtures compare native output, target projection, history, and saved documents. Refusal tests complement successful authoring cases. |
| Media added only for new images | The image fixture verifies deferred resources, existing native image identities and bytes, newly resolved image references, exact media relationships, and nonstandard main parts. Math, mention, and source operations allocate no image resources. |
| C01 matrix and composer regressions | The final source integration passes 658 cases across 55 files. The separately run table/native batch passes 89 cases across five files. Both batches use the same frozen runtime source. |
| Prerequisites | D15, D29, X23, and X38 are closed in their workstream ledgers. |

## Final implementation boundary

Online peer tails no longer invalidate the ownership proof of an independently marked atom paragraph. The proof excludes committed trailing text runs. Paragraph properties, controls, interior content, unknown elements, and foreign revisions remain guarded.

For a same-topology replacement, the coordinator verifies the physical cell, ordered atom slots, source identities, complete unchanged tail, and exact new deletion receipts. It withdraws only that stage's redundant deletion, removes the generated duplicate, and continues the original peer paragraph and text. No label matching grants identity. A repeated equivalent target produces no history frame. Its comparison invokes no host allocation or payload services.

Shared redo now recovers ordinary terminal text through formatting clocks removed by history. The character-deletion ledger blocks traversal through removed characters. Content-control labels keep their existing conservative boundary rule. This changes the evaluation format to `docx-collaboration:1.7.13.1`. Previous rooms require verified export and fresh seeding. Transport remains evaluated under the EigenPal Pro License. The private production adapter imports no Pro runtime.

The user selected independent commitment for untracked collaborator edits inside pending AI content. Four committed-view expectations and one native replacement expectation now express that policy. These expectation changes implement the selected policy. Native Word, media, history, convergence, and reopen oracles otherwise remain exact. Original timeouts remain unchanged.

## Verification

- Source integration: 658 passing cases, 55 files, 29,132 assertions.
- Table/native integration: 89 passing cases, five files, 2,479 assertions.
- Explicit atom planning and materialization: 57 passing cases, five files, 908 assertions.
- Private types: 2,301 files pass. Pro types: 2,332 files pass.
- Core, fonts, Pro, Markdown exporter, and private SDK builds pass. The private build retains 292 source modules with core external.

Compiled verification passes all 71 cases across six files and 3,799 assertions. It uses the built private SDK, one built core, the matching built Markdown exporter, and the built evaluation transport. Its first run identified an old exporter build and two unmapped dynamic fixture imports. The exporter and font declarations are rebuilt. The harness maps those imports and rejects remaining relative runtime imports. Test bodies and assertions remain unchanged by the harness.

Logs use `/home/lofcz/.cache/sciobot-docx-native/d30-primary-` with `final-integration`, `final-table-and-native`, `acceptance-materialization`, `final-sdk-types`, `final-pro-types`, `core-build`, `pro-build`, `markdown-build-final`, `compiled-verified`, and `sdk-build` suffixes.

## Separate remaining work

R19 retains ownership reconciliation for interior and styled recovered content, and replacements that change table topology. Further registered content combinations and application integration stay in their existing leaves. D30 closure does not establish full application parity or cutover readiness.

All active development uses `codex/react-block-proof`. Committed integration history through `da3fdaf0` is pushed to that branch. No new work branch is required.
