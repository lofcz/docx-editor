# Peer continuity across atom-table rewrites

D30 and R19 remain open. Work continues on `codex/d30-composer-reconcile`, after `cac25cd5`.

## Failure and repair

An offline participant appended text to a paragraph containing all four atom kinds. A second AI wave replaced the owned table and changed every atom. Reconnection preserved the participant's text in a recovered paragraph outside the replacement table. Replanning the intended target required inserting that text into the new cell and deleting the recovered paragraph.

The replacement coordinator now continues the exact native paragraph address for an owned atom paragraph. It derives candidates from selected insertion markers, complete native table births, surviving story boundaries, ordered tables, and matching physical cell topology. Equal labels and equal text grant no identity. The source and target paragraph counts must match. The source slot must remain constructor-empty after the owned inverse. Nonempty or independently formatted recovery slots remain untouched. Plain-text paragraphs keep their previous writer behavior.

The coordinator uses explicit native moves. It preserves the paragraph's Word identities and its existing property holder. It moves the new owned property children and inline content into that address. Final native validation runs before publication. The old owned atom payloads remain rejected; the new controls keep their complete payloads, resource reservations, and native receipts. No continuity metadata or canonical IDs enter a serialized proposal.

Two native transport defects also surfaced. Property-holder moves into a paragraph reached shared text instead of the record pipeline. The router now forwards those moves using record-child positions after source detachment. During undo, an empty retained constructor could win singleton selection before visibility filtering. Its invisibility then hid the earlier live paragraph-mark holder too. Materialization now excludes hidden constructors before choosing property winners.

The latter changes deterministic repair to `docx-collaboration:1.7.11.1`. Shared schema 7 and journal kinds stay unchanged. Previous-format admission refuses without writes. Deployment requires verified export through the previous compatible client and fresh room seeding. No saved room is migrated. Transport changes remain evaluation-only under the EigenPal Pro License. Production private runtime imports no Pro implementation.

## Exact oracles

Eight real-session cases cover both client orders, both decisions, and peer updates arriving before or after creation undo. The undo oracle is independently authored Word XML: one surviving row and column with the original column width and the participant's text. Rejection independently recovers that paragraph between the original body paragraphs. Acceptance uses independent direct execution. Fresh control identities bind before decisions. Full Word fingerprints, exact live convergence, image bytes and relationships, both histories, payload membership, and cold rooms remain checked.

Four further cases replace all atoms during the disconnection. They require replanning the complete intended target to return no operations. They check both creation frames, both decisions, both participants' histories, complete Word output, resources, and cold reopening. Remote caret checks use the continued paragraph address after undo, redo, and decisions. Both peers must agree on the caret position.

The pending package retains private empty recovery slots for other source paragraphs. Fresh Word identity pairing uses a Word round trip that hides only those empty slots. Expected authored Word content remains independent. Exact live canonical convergence remains a separate assertion. No original acceptance expectation or timeout changes.

## Verification

- The affected source boundary passes all 101 cases across fourteen files and 4,473 assertions. It includes the unchanged four plain-table peer rewrite cases, prior cell histories, property conflicts, native journal order, format refusal, and tombstones.
- Strengthened caret verification passes all four rewrite cases and 604 assertions in source and compiled packages.
- The compiled private SDK, built core, and evaluation transport pass all sixteen peer/table cases across three files and 1,684 assertions before the extra caret assertions. The strengthened four-case rerun also passes.
- Resource and receive-cost checks plus native property moves pass 21 cases across five files and 78 assertions. Existing limits remain unchanged.
- Native property tests compare independent native moves, repeated journal application, Word reopening, and cold rooms. The undone-constructor case also checks later edits and repeated winner changes.
- Private source types pass across 2,292 files. Pro source types pass across 2,330 files. Both evaluation builds pass; the private build verifies 288 modules with core external.
- Repository lint, scoped lint, line caps, formatting, and compatibility policy pass with existing repository warnings.

Logs use `/home/lofcz/.cache/sciobot-docx-native/d30-cell-continuity-` with `final-source`, `caret-source`, `compiled`, `caret-compiled`, `cost`, `sdk-types-final3`, `pro-types-final2`, `sdk-build-final`, `pro-build-final`, `lint`, and `policy` suffixes. Earlier failure logs use `d30-new-table-peer-waves-`.

## Remaining scope

This checkpoint closes no TODO. Interior peer placements, topology-changing replacements, nonempty or independently formatted recovery slots, further registered cell fragments, and application integration remain outstanding. Four earlier committed-view expectations remain unresolved. The full C01/composer gate has not been rerun at this checkpoint.
