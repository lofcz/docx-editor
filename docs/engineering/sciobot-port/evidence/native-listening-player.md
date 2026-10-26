# Native listening playback and generation

X10 closes in application commit `a7d38a465`, following `9c478e9d6`, `f75968f31`, and `378c7759e`, on `codex/react-block-proof`. The application checkout is `/GitHub/sciobot-next-docx-work`.

## Behavior

The native listening widget uses the existing `ListeningPlayer`, audio cache, and dialogue rendering services. Playback and hydration remain local React state. Restored recordings hydrate without generation, document writes, or history entries. The bounded transcript scrolls inside the paginated widget.

Explicit generation records the audio reference, take seed, resolved speaker voices, and validated audio metadata through existing per-field payload registers. Each successful take has one native undo entry. Unknown payload data survives. Native transcript text and unrelated peer edits retain their exact expected values.

Pending review cannot play audio for a different script. Reject restores matching playback. Accept permits explicit generation for the accepted script. Content locks retain playback and prevent generation. Ancestor locks, data bindings, and temporary controls also prevent payload writes. Scoped cancellation and source/host identity checks reject obsolete render responses.

## Verification

Run browser commands from the application checkout with these environment values:

```bash
export DOCX_NATIVE_SNAPSHOT=/home/lofcz/.cache/sciobot-docx-native/port-current-integration
export DOCX_HARNESS_PORT=8122
export DOCX_HARNESS_URL=http://127.0.0.1:8122
bun test --timeout 600000 scripts/docx-native/harness/listening-player.test.ts
```

The sustained-playback run at `378c7759e` passes 11 cases and 99 assertions. Its log is `/tmp/x10-browser-sustained.log`. Both client orders and reconnect receivers cover actual playback/pause, hydration, long transcripts, generation, deletion during pending work, separate take history, saved convergence, and cold reopening. Bob's player remains paused during Alice's playback. Native payloads, control fingerprints, text, and history are explicit oracles. Remote caret geometry is checked during peer editing and offline deletion.

The final production source passes this focused command:

```bash
bun test --timeout 600000 \
  --test-name-pattern 'proposed listening scripts|content-locked listening|explicit regeneration|a delayed listening render' \
  scripts/docx-native/harness/listening-player.test.ts
```

Four cases and 24 assertions pass, with nine unrelated cases filtered out. The log is `/tmp/x10-review-lock-final.log`. This verifies two additional cases and repeats owner/history checks after final guards. The selected commands verify 13 distinct player cases.

The final focused unit batch passes 31 cases across these files:

- `src/lib/docx/listening-locks.test.ts`
- `src/lib/docx/listening-audio.test.ts`
- `src/lib/docx/listening.test.ts`
- `src/hooks/chat/docxAtomicFallbacks.test.ts`
- `src/components/workspace/docx/blocks/registry.test.ts`

The unit log is `/tmp/x10-units-final.log`. Frontend and harness type checks pass 6,849 and 2,424 files. Logs are `/tmp/x10-fe-lock-final.log` and `/tmp/x10-harness-lock-final.log`. Scoped lint has no errors. React Compiler skips one manual memoization optimization; the warning is recorded in `/tmp/x10-lint-lock-final.log`.

All eight existing `listening-edits.test.ts` cases also pass. The log is `/tmp/x10-browser-review-verified.log`. That combined invocation later fails while loading another file because shared server cleanup closes its port. The affected player case passes in the final command. The combined invocation is not counted as a successful full command.

## Fixture and compatibility limits

Browser fixtures inject fake cache/render services and a local one-minute WAV. They consume no generation credits. The sustained WAV makes playing state observable under load. Original active/paused assertions and test deadlines remain intact. The existing peer typing check follows up to five actual native undo frames for five characters. Every frame retains exact AI groups, payloads, and native controls. Redo uses the same frame count and restores the exact final text.

The installed runtime remains `4f0029f0b8df49f1aa966c43ce52a7a97b8443b0`. The private archive remains `vendor/docx-native/sciobot-docx-853c85888966.tgz`. Existing extensible payload attributes store audio metadata; shared format is unchanged. No package rebuild, SQL migration, or public API change is needed. Production uses existing application services and Apache core/private adapter imports. Collaboration evaluation transport remains test-only under the EigenPal Pro License. Broader application cutover and later-wave recovery retain their existing open leaves.

## Reference inventory reconciliation

The reference checkout contains concurrent, uncommitted page-break support changes. Five inventoried sources changed, and the renderer became a discovered source. The reviewed diffs add explicit page-break renderers, Markdown parsing/serialization, slash insertion, and plain-text form feeds. Other changes are formatting. The inventory records these sources and preserves pending production verification. Open I11 includes explicit page-break migration; this closure claims no page-break parity. The original checkout remains untouched. Its HEAD stays `686f2b973`.
