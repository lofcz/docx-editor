# Logical joined-paragraph automation

This checkpoint repairs the inherited acceptance failure recorded in [native provenance evidence](native-text-slice-provenance.md). It does not close D30.

The canonical store retains physical paragraphs and private forward markers after resolving paragraph marks. Word export already folds those fragments. Automation now groups them for paragraph queries, body text, word ranges, reviewed text and searches. Ranges retain canonical fragment endpoints. Read-only projections keep every original content node and the surviving mark's properties. Content-control boundaries stop joins; a paragraph before a table can join its first cell's first paragraph.

Logical paragraph handles carry local coordinate metadata. Ordinary paragraph handles keep their original behavior. Start/end writes map to physical fragment boundaries. Cross-fragment replacement uses existing canonical range edits. Structural deletion removes the complete group. Before/after insertion uses logical edges. Delimiter splitting consolidates fragments with existing joins before splitting and retargets the logical proxy. A delimiter-free split leaves the document unchanged. No room field or stored DOCX field changes.

Focused checks pass the original six-case shared equation/mark file, existing story reads and searches, ten joined-paragraph runtime cases, and five immutable native grouping cases. An additional word-range case verifies tokens crossing old marks, both reviewed-text versions, and an editable token. The runtime and native grouping checks preserve exact assertions. Source type checks and root lint pass at this implementation checkpoint. Wider suites, compiled API verification and further shared-history checks remain pending.

The private production adapter still uses Apache core and its own package. The EigenPal Pro License collaboration transport and document API checks are evaluation fixtures.
