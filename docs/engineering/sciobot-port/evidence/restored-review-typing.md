# Restored peer typing in review branches

The linked-image regression now passes all eight original cases. The previous runtime failed both interior-peer acceptance-undo cases. Text remained present, but undo placed it outside its historical revision and omitted its proposed copy.

The raw text and formatting returned correctly. Native branch inference used the restoration site's insertion origin. Character collection also erased the original anchor's string content. Neither event changes the original character identity.

The evaluation transport now resolves restored original typing through its original identity. Its anchor walk compares exact character identities before checking retained string content. Collected items are traversed by clock, so merged deleted records retain the same interpretation.

The existing test keeps its complete native fingerprints, independent Word decisions, media bytes, relationships, payloads, histories, geometry, and reopening assertions. It now repeats decision undo/redo twice. Each undo also reopens the pending shared room in two fresh editors. No expected content or deadline changes.

## Compatibility and production boundary

The repair changes native materialization semantics. Repair version 15 fences the change as `docx-collaboration:1.9.15.1`. The format test rejects previous repair version 14 without shared writes. Private native ownership guards require the matching profile. Private ownership record encodings stay unchanged.

Export and verify DOCX with the previous compatible deployment before seeding fresh rooms. Keep original rooms and offline updates. Do not rewrite room metadata. DOCX does not transfer room undo history.

The transport remains evaluation-only under the EigenPal Pro License. This repair does not add that transport to the production SDK. R19 stays open for its remaining ownership routes and installed application checks.

## Verification

From the fork root:

```sh
bun test --timeout 60000 --preload ./scripts/test/preload-dom.ts packages/sciobot/src/review/recompose-inline-images-shared.test.ts packages/sciobot/src/review/recompose-inline-images.test.ts packages/sciobot/src/review/recompose-image-removal.test.ts packages/sciobot/src/review/recompose-image-removal-shared.test.ts packages/pro/src/collaboration/__tests__/revision-text-copies.test.ts packages/pro/src/collaboration/__tests__/document-format.test.ts packages/sciobot/src/collaboration/property-group-bindings.test.ts packages/sciobot/src/collaboration/property-birth-anchors.test.ts
```

All 79 cases pass, with 2,601 assertions across eight files. The log is `/tmp/r19-image-history-batch.log`.

The related batch covers `recompose-links`, `recompose-linked-inline-births-shared`, `recompose-mixed-collaboration`, `recompose-payloads-collaboration`, `document-typing-after-move`, `paragraph-text-typing`, and `review-sequences`. All 95 cases pass, with 3,903 assertions across seven files. The log is `/tmp/r19-related-history-batch.log`. It uses the same test command and unchanged timeout. Combined coverage is 174 cases and 6,504 assertions.

Private types pass across 2,427 files. Pro and Vue types also pass. Logs are `/tmp/r19-sdk-types.log` and `/tmp/r19-pro-types.log`.
