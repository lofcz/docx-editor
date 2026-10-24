# Managed inline image replacement checkpoint

Runtime `42eaeaf0` replaces managed images through explicit private plans and revision-bound resource preparation. It retains the original control, drawing, geometry, crop, and unmapped native metadata. Metadata-only edits fetch nothing. Byte-identical resources keep their native relationship and media part.

Shared replacements use the existing application-owned scalar history scope. Resource retirement clears existing descriptor scalars instead of deleting their shared maps. Ordinary descriptor puts remain atomic. Both reader implementations remain byte-identical to the previous runtime. They already omit empty relationship identifiers and empty binary digests. The shared room version fields do not change.

The production private SDK imports no Pro runtime. Evaluation transport fixtures use the EigenPal Pro License.

Verification at this boundary:

- The combined native batch passes 255 cases across seven files, with 30,012 assertions. It includes mixed and repeated atoms, image replacement, and preparation/codec refusals.
- The built SDK passes the same 88 image cases across two files, with 12,092 assertions. Full native Word, payload, media, both actor histories, remote geometry, awareness, and DOCX/room reopening remain required.
- Four evaluation transport image regression files pass seven cases and 54 assertions.
- Root formatting and lint pass. Core, private SDK, and Pro type checks and builds pass. The core API extraction scans 15 entries with zero errors. Compatibility policy passes against `a3dd59eb`.

D30 remains open. Combined replacement and fresh atom insertion still need reconciliation. The final C01/composer closure audit has not run for this checkpoint. The production inventory revision drift remains unresolved; no baseline was recaptured.

Machine-readable results: [inline-image-replacement.json](inline-image-replacement.json).
