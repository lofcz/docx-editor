# Explicit hyperlink text ownership

Runtime `d11c08c7` repairs the imported linked-image boundary regression. Follow-up `c37f8096` preserves existing tracked leading-link insertion behavior. The original regression keeps all native metadata, inherited color, opaque run properties, derived code style, media, undo/redo and reopen assertions. It also checks the exact insertion offset and the unchanged outside text leaf.

## Behavior

The native `insertText.inside` field can name a typed hyperlink for untracked insertion. The same landing resolver selects the destination for application and content restrictions. Locks, bindings, temporary-control transitions and document forms protection remain enforced. Invalid owners, unrelated paragraphs, outside offsets, deleted revision content and tracked hyperlink-owned insertion refuse explicitly. Legacy protected form-field writes retain their content-control-only owner validation. Ordinary typing keeps its existing behavior.

The private text planner resolves repeated-character ambiguity against complete target link spans. It changes an insertion coordinate only when the resulting native text stays exact and more original hyperlink boundaries survive. It chooses the closest equivalent boundary. UTF-16 coordinates remain intact for supplementary characters. Leading insertions continue to use existing right bias, preserving their tracked proposal support. Trailing insertions carry an explicit private `insideHyperlink` address. The direct executor verifies that this address still identifies a hyperlink. The plan codec accepts it only for pure supported text insertion. The tracked executor refuses unsupported trailing ownership rather than dropping it.

No public Office-shaped document member, shared field, message kind or stored primitive changes. The minor core changeset records the expanded native owner support. The private plan field does not enter shared rooms. The EigenPal Pro License transport remains confined to evaluation fixtures.

## Verification

- Native ownership cases: 14 pass, including exact envelope identity and protection refusals.
- Broader native checks: 1,094 pass across 45 files, with 4,629 assertions.
- Focused final private batch: 57 pass, with 1,662 assertions. This includes imported images, four independent shared histories, repeated letters, supplementary characters, adjacent links and existing link recomposition.
- Plan codec checks: 11 pass, with 89 assertions.
- Other affected private cases: 82 pass across six files, with 1,544 assertions.
- Final compiled SDK/core/exporter matrix: 91 pass, with 10,228 assertions. It retains the original fixture bodies and original timeouts.

The first concurrent compiled run had one unchanged five-second image-case timeout. The final run passes all 91 cases in 46.66 seconds. No timeout was widened.

Four shared imported-image cases preserve foreign text outside the exact source hyperlink through both actors' histories. They retain imported attributes, derived styles, opaque metadata, media, and DOCX/room reopening. The original outside text leaf carries only its own peer edit. Existing image and retained mark/link acceptance bodies also pass against compiled packages.

The earlier reader at `/tmp/docx-upstream-oct23-timing` retains 20 exact native snapshots, including its own ordinary edits, undo/redo, and current-writer undo/redo. It need not support the expanded authoring option to read its existing native splice primitives.

```sh
bun --preload ./scripts/test/preload-dom.ts packages/sciobot/scripts/check-hyperlink-text-reader-compatibility.ts /tmp/docx-upstream-oct23-timing
bun packages/sciobot/scripts/check-code-inline-controls-compiled.mjs inline-images-materialization.test.ts inline-images-collaboration.test.ts imported-image-boundary-collaboration.test.ts retained-atom-links.test.ts retained-atom-links-collaboration.test.ts retained-atom-marks.test.ts retained-atom-marks-collaboration.test.ts
```

Private types cover 2,203 files. The build verifies 261 private modules with core external. Core builds and API extraction/check pass without API snapshot changes. Root formatting, lint and collaboration policy pass. The policy checks 188 decisions and two inherited format fields.

Logs: `/tmp/d30-owned-text-native-broad.log`, `/tmp/d30-leading-owner-source.log`, `/tmp/d30-leading-owner-codec.log`, `/tmp/d30-owned-text-private-broad.log`, `/tmp/d30-owned-text-compiled-final.log`, `/tmp/d30-owned-text-reader.log`.

## Integration gate follow-up

Runtime `827f0820` / `9c8a08c8` resolves the 64 historical composer failures below. The final 333-case gate passes with every original fixture unchanged. See [shared history repairs](composer-shared-history-repairs.md). The linked retained/new atom planning probes remain open.

The complete C01/composer run at `d11c08c7` has 269 passing cases and 64 failures across 333 cases. Each failure also occurs at the earlier port checkpoint `88c58801`. All failure names, files and displayed diffs match. This comparison establishes their origin; it does not establish a passing integration gate. The [failure inventory](composer-inherited-failures-2026-10-24.json) retains all 64 names and baseline logs. No expectation changes or baseline recapture occurred. Follow-up leading-link checks pass the existing link recomposition cases.

These failures affect later proposal waves, born tables/lists and offline old-row deletion. One representative expected view contains `Alpha Q end` and `Final`; the actual view retains `Alpha XYZ end` and `New` after reconnect. Repairing these suggestion histories is a priority for the complete port.

Two additional [linked retained/new atom probes](regressions/linked-retained-and-new-atoms.test.ts) fail during planning. They preserve complete media, original-control identity/fingerprint, payload, undo/redo and reopen assertions. One inserts an image beside retained linked math. The other inserts math beside a retained linked image. The current gap planner rejects hyperlinks, and the unchanged-atom stream requires the same complete atom sequence. These cases require whole-paragraph source ownership across retained and fresh atoms.

D30 remains open. Image removal/replacement, nested registered cell fragments, tracked atom ownership and installed application parity remain open. The backlog stays at 179 closed and 97 open leaves.

The production inventory still rejects application revision drift. All 199 pinned file hashes match. The application and inventory baseline remain unchanged.
