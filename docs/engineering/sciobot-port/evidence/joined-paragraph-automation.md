# Logical joined-paragraph automation

This checkpoint repairs the inherited acceptance failure recorded in [native provenance evidence](native-text-slice-provenance.md). It does not close D30.

The canonical store retains physical paragraphs and private forward markers after resolving paragraph marks. Word export already folds those fragments. Automation now groups them for paragraph queries, body text, word ranges, reviewed text and searches. Ranges retain canonical fragment endpoints. Read-only projections keep every original content node and the surviving mark's properties. Content-control boundaries stop joins; a paragraph before a table can join its first cell's first paragraph.

Logical paragraph handles carry local coordinate metadata. Ordinary paragraph handles keep their original behavior. Start/end writes map to physical fragment boundaries. Cross-fragment replacement uses existing canonical range edits. Structural deletion removes the complete group. Before/after insertion uses logical edges. Delimiter splitting consolidates fragments with existing joins before splitting and retargets the logical proxy. A delimiter-free split leaves the document unchanged. No room field or stored DOCX field changes.

Focused checks pass the original six-case shared equation/mark file, existing story reads and searches, ten joined-paragraph runtime cases, and five immutable native grouping cases. An additional word-range case verifies tokens crossing old marks, both reviewed-text versions, and an editable token. The runtime and native grouping checks preserve exact assertions. Source type checks and root lint pass at this implementation checkpoint. Wider suites, compiled API verification and further shared-history checks remain pending.

The private production adapter still uses Apache core and its own package. The EigenPal Pro License collaboration transport and document API checks are evaluation fixtures.


The broader core automation run at `9cfd1f3d` passes 654 cases and exposes three failures. The 4,000-paragraph benchmark passes separately under its unchanged 4,000 ms limit; the prior reader/writer revision also passes that check. Two native tracked-row acceptance fingerprints fail unchanged at `15b9dcfe`; they remain separate inherited defects. The complete document API runtime run passes 730 cases and exposes six DOM-free compilation diagnostics from direct `crypto.randomUUID()` calls in existing native metadata writers. The following UUID repair keeps their existing UUID representation and removes the DOM type dependency.

Four shared cases cover both client orders and both replacement writers. A tracked split is accepted on the other peer, then a range spanning physical fragments is replaced. Both peers retain exact saved Word XML through writer undo and redo. Cold server reopening and re-saving produce the same native XML. These cases pass with 72 assertions.


The DOM-free repair passes the original neutral-project compiler assertion and three secure-UUID cases. The prior revision fails that same compiler assertion. All 80 native origin/clipboard/export cases pass unchanged with 256 assertions. Core types pass across 1,601 files. The complete runtime suite's 730 behavioral cases pass; its only compilation failure passes with this repair. The two inherited native table fingerprint failures remain open. No broad all-green claim or TODO closure is made.


Final tracking coverage passes all 369 cases across 27 files, with 3,279 assertions. It includes the unchanged acceptance case that failed before this repair. Fresh compiled core and document API packages pass the same four shared-history cases and all 72 original assertions. Run `bun packages/sciobot/scripts/check-joined-paragraphs-compiled.mjs` after building both packages. The runner changes imports only and retains the test-only evaluation transport. Fresh compiled core also passes all 56 private code-control fixtures and 1,267 assertions.

Core and document API builds and API checks pass without public signature snapshot changes. Office.js signature reporting passes; it does not establish full runtime equivalence. Root formatting and lint pass, with targeted formatting checks for follow-up files. Compatibility policy covers 183 decisions and two inherited format fields. No TODO closes. The next native table repair must preserve the empty `trPr` and `tcPr` holders that direct row creation retains; acceptance currently removes them. Keep both original fingerprint assertions.
