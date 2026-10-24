# Peer text in paired review snapshots

Branch: `codex/d30-link-source-replicas`, after native provenance checkpoint `0752ebc7`.

## Behavior

Historical and proposed hyperlink text retain explicit native source identities. Opposite revisions with matching author and date can show both copies. Equal text without source identity does not establish a pair. More than two copies and mismatched review ownership do not establish a pair.

Offline peer text inside the historical source projects into both snapshots at the same character boundary. Placement uses the exact physical counterpart item. The proposed copy receives its own revision, text-element kind, and hyperlink attributes. Source formatting remains intact.

Before a native decision adopts projected text, it persists the historical source's derived revision membership. Both physical copies then match the canonical snapshot used by preflight. Removing one copy does not globally delete a surviving source identity.

Creation undo removes paired revision keys introduced by that exact history step. This also removes later Yjs cleanup markers that would otherwise strike surviving source text. Redo follows author-local redo items and restores the observed character cohorts. This does not place foreign edits into the author's history.

## Verification

The unchanged eight-session hyperlink regression passes all 344 assertions. It covers outside and interior peer edits, both client orders, exact native decisions, payload retirement, both actor histories, decision histories, caret geometry, DOCX reopening, and cold room reopening.

The final focused source command passes 31 cases and 467 assertions. It includes the same eight sessions, source-identity negative cases, and read-only rejection of rooms using the previous format. Pro source types pass 2,323 files. Private types pass 2,260 files.

The first affected batch passes 258 cases. Its deletion-index random oracle exceeds the original five-second timeout during the batch. The unchanged case passes separately in 4.564 seconds. The withdrawal cost check passes in 64.93 milliseconds. No timeout or cost limit changes.

The wider composer batch exposes four committed-view expectations in `recompose-collaboration.test.ts`. The previous `2a244d56` checkpoint reproduces all four failures unchanged: nine cases pass and four fail. The peer text appears in the committed view after the earlier insertion-ownership repair. This work does not change those assertions or claim the composer gate passes.

The built SDK, core, and evaluated transport pass 84 cases and 1,730 assertions across the linked, retained, and born-paragraph fixtures. The private build verifies 282 modules with core external. Scoped formatting, lint, and collaboration policy pass.

Logs: `/home/lofcz/.cache/sciobot-docx-native/d30-paired-*`.

## Format and limits

The evaluated transport changes deterministic interpretation of existing identity and revision marks. Repair version increases from 5 to 6. Format is `docx-collaboration:1.6.6.1`; no shared fields are added. The migration decision and guide require coordinated export and fresh room seeding. No saved room is migrated here.

The transport remains evaluation-only under the EigenPal Pro License. The private production adapter imports no transport implementation. Core stays external in the private build. The application checkout remains unchanged.

D30 and R19 remain open. Tracked images, born-cell ownership, mixed ownership, and installed application gates remain required.
