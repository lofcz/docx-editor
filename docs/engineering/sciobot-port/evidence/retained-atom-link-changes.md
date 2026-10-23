# Retained inline atom hyperlink changes

Runtime `fc6cd155` resolves unchanged atom sequences to exact native widths before planning hyperlink structure. It supports insertion, expansion, contraction, retargeting, bookmark targets, distinct adjacent links, and removal around retained math, mentions, sources, and images. Mark planning uses the same resolved streams, including when wrappers split ordinary text tokens differently. Existing native property planners still preserve other formatting.

Native direct, unstyled link envelopes keep wholly covered content controls unchanged. Partial, styled, tracked, and ancestor-protected edits retain existing restrictions. Removing an outer native hyperlink preserves its descendant controls. Bound controls also retain their complete original binding and content. The reach classifier now distinguishes an unchanged complete control from edited content. It still identifies envelope creation as document structure for forms protection.

Drawing formatting records carry their resolved link envelope. Markdown groups text and drawings under the same native link identity. Image source spans remain inside that envelope, so replacing the image's projection retains its link.

Markdown cannot nest a mention's link syntax inside another link. Linked mentions therefore use literal `<scio-mention key="user" value="Ada" />` syntax. Parsing recovers the existing mention AST without expression evaluation. Unknown attributes, duplicate fields, expressions, and children retain opaque meaning. Ordinary mention syntax remains unchanged. Native revision wrappers do not hide an enclosing link from this serializer.

## Authored typing boundaries

The stronger shared fixture first fails six of twelve cases. A peer types ` peer` at offset 1 in the native leaf `A `. Local Word output is correct, but the journal normalizes that insertion to `peer ` at offset 2. A concurrent link edit then lets the neighboring atom's wrapper capture that text.

Runtime `b8764ccc` records the native insertion's exact source object and boundary in transaction-local metadata. Primitive lowering keeps the existing `spliceText` shape and fields. The evaluation transport retains those character origins for ordinary splice journals, as it already does for explicit text slices. No new shared field or message is introduced. Earlier writers retain their old lexical-diff behavior.

All twelve original shared assertions now pass unchanged. They cover both peer orders, both link writers, addition, retargeting, removal, disconnected same-paragraph typing, explicit peer deletion, both actors' histories, remote selection geometry, and DOCX/room reopening. Independent native oracles preserve whole Word fingerprints, registered payloads, and imported link metadata. The isolated native boundary cases also cover repeated letters and supplementary characters.

The EigenPal Pro License transport remains evaluation-only. Production still uses Apache core and the private adapter.

## Verification

- Native journal, hyperlink, and control checks: 955 cases across 31 files, 4,189 assertions. Log: `/tmp/d30-retained-atom-links-core-focus.log`.
- Private C01, atom, image, link, direct-edit, property-review, and recomposition checks: 267 cases across 22 files, 8,877 assertions. Log: `/tmp/d30-retained-atom-links-sdk-focus.log`.
- Markdown exporter checks: 233 cases across 34 files, 1,379 assertions. The unchanged 500-page resource checks pass. Log: `/tmp/d30-retained-atom-links-exporter-source.log`.
- Evaluation transport checks: 133 pass and one existing deletion-index timeout across 19 files. That unchanged file passes both cases separately under its original 5,000 ms limit. The attached keystroke checks pass their original limits. Logs: `/tmp/d30-retained-atom-links-transport-focus.log`, `/tmp/d30-retained-atom-links-deletion-index-isolated.log`.
- Fresh compiled core, exporter, and private SDK pass the same 38 retained mark/link cases and 2,162 assertions. The runner rewrites imports only. Log: `/tmp/d30-retained-atom-links-compiled.log`.
- The earlier reader at `15b9dcfe` still passes 72 exact native table/history snapshots. Log: `/tmp/d30-retained-atom-links-prior-reader.log`.

Reproduce the compiled check after building all three packages:

```sh
bun packages/sciobot/scripts/check-code-inline-controls-compiled.mjs retained-atom-links.test.ts retained-atom-links-collaboration.test.ts retained-atom-marks.test.ts retained-atom-marks-collaboration.test.ts
```

## Remaining D30 work

The [linked image birth probe](regressions/linked-image-births.test.ts) preserves four additional failing targets and their complete original acceptance bodies. Two inline contexts refuse image reservation ownership before resource resolution. Two image-only contexts enter gallery/application-block normalization and lack the required registered payload materialization. Their original target strings and native/history/media assertions remain intact. They are not part of the passing retained-link claim.

```sh
bun test --preload ./scripts/test/preload-dom.ts ./docs/engineering/sciobot-port/evidence/regressions/linked-image-births.test.ts
```

Log: `/tmp/d30-linked-image-births-open-probe.log` (four failures). Existing ordinary-image tests and their unsupported ownership expectations remain unchanged. Implement native image reservation ownership next, then reconcile image-only block targets through their actual domain codec.

Retained tooltip removal still rebuilds its link through the existing plain-text planner. Retained atom/text replacement, image removal/replacement, new nested cell fragments, tracked atom ownership, and installed application parity remain open. D30 remains unchecked. No TODO closes at this checkpoint.

Builds, core/exporter API checks, root formatting, root lint, and compatibility policy pass. Core types cover 1,604 files; exporter types cover 1,533. Evaluation transport types cover 2,309 React files and 1,860 Vue files. Private types cover 2,198 files. The private build verifies 259 source modules with core external. Compatibility policy covers 186 decisions and two inherited format fields.

The unchanged production inventory still rejects application revision drift. All 199 pinned source hashes match. The baseline remains intact; no full production audit pass is claimed. The backlog remains 276 leaves, with 179 closed and 97 open.
