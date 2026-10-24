# Native tracked text copy recipes

D30 and R19 remain open. This WIP stays isolated on `codex/d30-link-source-replicas` in `/home/lofcz/.cache/sciobot-docx-native/d30-link-source-replicas`.

The verified writer remains on `codex/d30-inline-births-recovery`. This branch does not establish a fix for the interior peer regression.

The native hyperlink writers now record the actual source positions for both historical and proposed copies. Each copy keeps independent native node identities. The transaction journal uses its existing `sourceText` recipe shape. No public API or shared format field changes.

Proposed link properties still come from the native preview. Copy recipes use actual selected source nodes. Suppressed preview objects cannot lend provenance to the live transaction. The positional copier supports differently partitioned leaves, repeated equal strings, partial ranges, Unicode, and optional native hyperlink styles. It refuses incomplete or changed text projections before recording a recipe. Properties and non-text nodes keep ordinary cloning behavior.

Verification:

- Native hyperlink, primitive journal, and copy-recipe checks pass 77 cases across five files with 901 assertions. Log: `/home/lofcz/.cache/sciobot-docx-native/d30-copy-recipes-final-native.log`.
- Nine copy-recipe cases assert exact source IDs and offsets, distinct repeated source leaves, source timing, exact accept/reject fingerprints, and XML reopening. Invalid source projections refuse.
- Private source types pass 2257 files. Log: `d30-copy-recipes-types.log` in the same directory.
- Scoped lint, max-lines caps, formatting, and whitespace checks pass. Compatibility policy passes 211 decisions with two cumulative format fields changed; this checkpoint adds none.
- The unchanged shared fixture still passes its four outside-range cases and fails its four interior cases. Its proposed view hides the interior edit in a deletion branch. The existing transport still treats copied text as moves or fresh content, rather than paired review snapshots. The earlier combined log is `d30-text-copy-batch.log`; its three fixture-observation failures were corrected by using an observed transaction frame. The final native checks above pass those fixtures.

The next stage must retain paired historical/proposed text references in the transport. It must preserve causal revision membership and both copies' peer edits through review and history. Source recipes alone do not satisfy the native accept/reject, shared history, or reopening requirements. Any change to deterministic reading or shared copy representation needs an explicit compatibility decision and version review.

No compiled integration, release build, migration, application change, or TODO closure is claimed for this WIP. The application remains read-only. Test-only transport uses the EigenPal Pro License.
