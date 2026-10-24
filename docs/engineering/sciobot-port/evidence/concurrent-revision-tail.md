# Concurrent tracked terminal text

D30 and R19 remain open. This checkpoint repairs the inherited offline suffix failure in owned presence recomposition.

The original case replaces `Outside` with `Rewritten` while another participant appends ` Peer` offline. One client order previously produced `Rewritte Peern`. A source-range expansion experiment hid peer text in a deletion and was discarded.

The transport now derives terminal insertion placement from stable character identities. Concurrent plain text and a tracked terminal append must name the same original boundary. The tracked append appears first. Causal text typed after observing that append retains its position. Existing shared characters, payloads, and native identities remain intact.

Undo can replace every original character with a restored copy. Its nearest live insertion origin then skips original characters that still appear as restored copies. The bounded boundary walk follows original identity clocks and recognizes visible identities before interpreting collected content. It uses one identity set for both candidate runs. Existing move/follow views remain authoritative when they already place or hide text.

The transport uses the EigenPal Pro License for evaluation. No production private adapter import changes. This changes deterministic repair, so repairVersion increases from 3 to 4. The shared schema remains 6. The compatibility decision and export-and-reseed instructions accompany the change. No migration runs in this task.

The temporary worktree and uncommitted prototype disappeared between turns. The committed checkpoint `40b95593` remains intact. Work resumed in `/home/lofcz/.cache/sciobot-docx-native/d30-inline-births-recovery`, branch `codex/d30-inline-births-recovery`.

Verification uses source aliases in `tsconfig.json`:

- `bun test --tsconfig-override tsconfig.json` with `owned-presence-shared.test.ts`, `terminal-revision-boundary.test.ts`, and `recompose-inline-births.test.ts`: all 32 runtime cases pass. The final combined log is `/home/lofcz/.cache/sciobot-docx-native/d30-recovery-repair-batch.log`. Its one initial failure came from a new compatibility assertion, corrected below. The nine original owned-presence cases retain exact native decision fingerprints, histories, geometry, and room reopening.
- `document-format.test.ts`: 18 pass, zero failures, 106 assertions. Final log: `/home/lofcz/.cache/sciobot-docx-native/d30-recovery-format-final.log`. Reading earlier metadata remains allowed; admission refuses it without writes.
- Seven affected transport files covering revision boundaries, revisions, foreign revisions, moved-text undo, following positions, following cost, and revision decoding: 164 pass, zero failures, 1613 assertions. Log: `/home/lofcz/.cache/sciobot-docx-native/d30-recovery-repair-transport.log`.
- Pro source types: 2314 files pass. Private adapter source types: 2244 files pass.
- Scoped lint, max-lines caps, collaboration policy, and whitespace checks pass. Policy reports 204 decisions.

`bun docs/engineering/sciobot-port/check.mjs` remains blocked by the frozen production coverage snapshot mismatch. The audit instructs recapture; this task preserves the original snapshot and leaves the gate failing. The application remains read-only.

Original acceptance fingerprints, history assertions, and timeouts remain unchanged. No baseline was recaptured. No new compiled package verification occurs at this checkpoint. Full C01/composer integration and remaining D30 features still require completion. No backlog leaf closes from this checkpoint alone.
