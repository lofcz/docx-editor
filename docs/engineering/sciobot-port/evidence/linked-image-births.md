# Linked and marked image births

Runtime `2b4fe004` replaces only the validated image reservation text with its native drawing. The reservation control, native run and hyperlink remain. The materializer preserves inherited properties, applies target marks and code styles, stores media provenance, and locks the completed image. Drawing namespace bindings travel with the native picture.

Native target parsing keeps linked image paragraphs, including image-only paragraphs. Unlinked image-only targets still use galleries. Explicit `preserveLinkedImages: false` retains captured Plate parsing. All immutable production AST fixtures remain unchanged and pass in that compatibility mode. This extends native grammar without recapturing production baselines.

## Verification

The four original [linked image probes](regressions/linked-image-births.test.ts) pass with their complete original targets and assertions: 4 tests, 96 assertions. These cover retained paragraphs, new image-only paragraphs, table cells and registered phase containers.

The affected source batch passes 96 tests and 7,899 assertions. It includes 24 shared-history image cases across ordinary, linked and linked-code ownership. Assertions retain independent media bytes, relationship isolation, foreign text, both actor histories, geometry and DOCX/room reopening. Another 55 related source cases pass with 790 assertions, including C01 and linked/code controls.

The compiled SDK, core and exporter pass the same 86 image and retained mark/link fixtures with 9,640 assertions. Only imports change. The runner now maps the private exported `targetAstKey` comparator to the compiled SDK. The build verifies 260 private modules with core external.

```sh
bun packages/sciobot/scripts/check-code-inline-controls-compiled.mjs inline-images-materialization.test.ts inline-images-collaboration.test.ts retained-atom-links.test.ts retained-atom-links-collaboration.test.ts retained-atom-marks.test.ts retained-atom-marks-collaboration.test.ts
```

Logs: `/tmp/d30-linked-image-original-probes.log`, `/tmp/d30-linked-image-batch.log`, `/tmp/d30-linked-image-related.log`, `/tmp/d30-linked-image-compiled.log`. Private type checking covers 2,199 files. Root formatting, lint and collaboration policy pass; the policy checks 187 decisions and two inherited format fields. Published core and exporter sources do not change in this batch.

## Remaining work

The stronger [imported link boundary regression](regressions/imported-linked-image-boundary.test.ts) fails at the original hyperlink identity assertion. Its source has imported history/frame attributes, a character style, a color and opaque run metadata. Inserting an image after `label ` inside that link produces a lexical space insertion after an identical existing space outside the link. The skeleton planner removes the old wrapper. The native trailing-link typing rule also keeps automatic typing outside a link, even with left bias. Moving an equivalent space insertion to that boundary alone does not fix ownership.

Keep this full regression and its native metadata, derived code style, undo/redo and reopen assertions. Resolve explicit link ownership through planning and execution next. Do not relax ordinary native typing semantics or replace source identity with Markdown equality.

D30 also needs image removal/replacement, changed atom/text sequences, nested registered cell fragments, tracked atom ownership and installed application parity. No TODO closes here: 179 of 276 leaves are closed, and 97 remain open.

The production inventory still fails application revision drift. All 199 pinned file hashes match. No baseline was recaptured.
