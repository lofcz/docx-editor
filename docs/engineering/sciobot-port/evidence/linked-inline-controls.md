# D30 linked inline controls checkpoint

The isolated branch `codex/d30-linked-atoms` follows the native table history repair branch.
Runtime `38f89183` implements linked controls; `6f2cacea` adds shared-history verification.
The primary branch remains at `75930cb4`. D30 remains open.

## Native behavior

New paragraph, heading, table-cell and registered-block fragments can contain math and source controls inside native hyperlinks.
The existing native fragment writer owns the outer hyperlink; atom XML renders only the control and its marked fallback.
Retained paragraphs allocate their atoms before wrapping complete authored hyperlink ranges.
The planner excludes temporary text-only skeleton links when the atom executor owns the complete final link.
Imported links unwrap through existing native operations and regain their original identity and metadata after wrapping.
External relationships resolve inside the admitted transaction; bookmark targets remain native anchors.
No planning fetch, DOCX reconstruction, public API change or production Pro import is added.

Retained control fallback runs receive explicit target marks.
Other inherited native formatting remains intact.
This prevents the second atom from inheriting the first atom's bold mark while preserving imported colors.
No-op planning compares hyperlink targets, titles and per-token ownership before accepting existing opaque atoms.
Different link targets or ownership still require separate reconciliation.

The linked table case also exposes a native terminal paragraph absent from Markdown.
The planner keeps unrepresented empty trailing paragraphs outside the editable sibling range, as its `afterId` boundary.
It does not remove or rewrite those paragraphs.
The table test preserves the exact complete Word fingerprint when replanned and reopened.

## Focused verification

All five new-location and retained-paragraph cases first fail before implementation.
The final local matrix has six cases and 123 assertions.
It includes an imported hyperlink with `w:history`, an existing relationship ID, and distinct inherited colors.
Checks cover native hyperlink/control ownership, separate bold marks, exact imported metadata, D01 reconstruction,
empty replanning, exact Word fingerprints through undo/redo, and DOCX reopening with unchanged payload records.

Four real-session cases cover both client orders and reconnect receivers.
Bob types inside the original linked text while disconnected. Alice undoes the atom insertion before reconnect.
Both peers match an independent native edit of the original imported Word tree.
The original link identity and metadata survive, and atom payload records disappear with their author.
DOCX and shared-room reopening match that exact native tree.
Both actors' undo/redo restore complete original, peer-only, and authored Word fingerprints.
All four cases pass with 88 assertions in 6.43 seconds.

The old linked-math refusal target moves unchanged into the existing positive materialization matrix.
That matrix now verifies successful landing, empty replanning, both histories and Word/payload reopening for that exact target.
The unprepared-image refusal remains unchanged.

The private SDK type check passes across 2172 files. Root lint passes.
Root formatting passes. Compatibility policy passes with 170 decisions and two inherited format field changes.
The private SDK build verifies 256 bundled source modules and keeps native core external.
Final verification passes 1063 cases across 103 files, with 38783 assertions, in 787.53 seconds.
The command covers all private diff and AI cases plus every `review/recompose*.test.ts` fixture.
It includes the C01 matrix, both new linked-control suites, native atom and image materialization, streaming,
and review composition across native tables, lists, arbitrary React content, imported links, and peer histories.
The four original table-history regression probes also pass unchanged, with 34 assertions, in 4.96 seconds.
Shared schema, payload schema and native primitive kinds do not change in this private adapter batch.
The inherited native history repair still proposes an unreleased format upgrade and stays isolated.

The stock production coverage audit rejects its full snapshot comparison.
Its pinned application revision is `5650bdf550b7987174d1b61fe2a925b86fb24f15`; the current revision is `073cb22447c6ac6a3699958e27bb1f3b8661bd22`.
An independent SHA-256 comparison verifies all 199 pinned files unchanged.
The application checkout has unrelated curriculum batch edits; this work does not change them.
The baseline is not recaptured, and no full production audit pass is claimed for this checkpoint.

## Remaining work

Linked drawings, code-marked atom rendering, edits that change existing atom hyperlink ownership,
and further nested/registered new-cell fragments remain unfinished.
Tracked atom ownership and installed application verification also remain open.
The table history repair retains its separate unresolved geometry cases.
No D30 closure is claimed.
