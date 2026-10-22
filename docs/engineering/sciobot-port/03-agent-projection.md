# Agent document projection and targeting

Keep the existing Markdown tool protocol. Map targeted changes back to canonical OOXML without rebuilding the document.

Read the [dispatch and completion rules](../../../SCIOBOT_PORT_TODO.md) before claiming a leaf.

Each named test or artifact is a deliverable. These specifications do not claim those checks already exist.

## P01

- [x] Project native text with stable offsets

**Depends:** [C03](01-contracts.md#c03), [C07](01-contracts.md#c07)

**Primary files:** `fork:packages/sciobot/src/projection/text.ts`; `sibling text.test.ts`

**Read:** [S02](sources.md#s02), [S25](sources.md#s25)

**Change:** Project body text with block IDs, revision tokens, and UTF-16 span mappings. Preserve hard breaks and empty paragraphs.

**Close when:** Czech accents, emoji, tabs, soft breaks, and hard breaks map to exact original spans. Reading never mutates the store.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `projection/text.test.ts` passes six tests and 23 assertions for accents, surrogate pairs, tabs, native line/page breaks, empty paragraphs, repeated text, and nested SDT/table reading order. Source IDs and UTF-16 ranges remain exact. Reading preserves the canonical fingerprint and returns JSON-safe data with a copied revision token. Package typecheck, scoped Prettier, and oxlint pass. Review view derivation remains assigned to R01/R02.

## P02

- [x] Project native Markdown structure and marks

**Depends:** [P01](03-agent-projection.md#p01), [P14](03-agent-projection.md#p14), [P15](03-agent-projection.md#p15)

**Primary files:** `fork:packages/sciobot/src/projection/markdown.ts`; `sibling markdown.test.ts`

**Read:** [S03](sources.md#s03), [S11](sources.md#s11)

**Change:** Serialize headings, paragraphs, quotes, lists, links, and marks. Map generated syntax separately from native text.

**Close when:** C01 formatting fixtures produce copyable Markdown. A bold delimiter never becomes a native text offset.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `projection/markdown.test.ts` passes five cases and 55 assertions over real imported DOCX layouts. C01 before/after formatting trees yield explicit copyable Markdown for bold, italic, ordered lists, and bullets. Additional cases cover headings, quotes, links, escaping, emoji, generated delimiters, and original/proposed native source offsets. Hidden review boundaries require explicit endpoint affinity; escape interiors and surrogate splits fail. Layout fingerprints stay unchanged and projection data is JSON-safe. The private package suite passes 72 tests and 524 assertions. Package typecheck, scoped format/lint, and the external-runtime build pass. Native Markdown escaping/list spacing differs from Plate; echoes use landed native output. P10 still owns structural target validation, and R01/R02 own review view derivation.

## P03

- [x] Project table cells with boundary mappings

**Depends:** [P02](03-agent-projection.md#p02)

**Primary files:** `fork:packages/sciobot/src/projection/tables.ts`; `sibling tables.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Serialize table rows and cells with escaped pipes. Retain cell IDs and explicit structural boundaries.

**Close when:** A repeated cell value resolves to the selected table cell. Embedded pipes and line breaks survive projection without shifting targets.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `projection/tables.test.ts` passes five imported-DOCX cases and 38 assertions. Native table, row, cell, and paragraph IDs resolve repeated values to the selected cell. Pipes, native breaks, empty cells, nested tables, grid spans, vertical continuations, repeated headers, and real page continuations retain their source boundaries. Generated separators and escape interiors never select native cell text. Empty/hidden cells keep identities without invented offsets. Layout records remain unchanged and the result is JSON-safe. The full private suite passes 92 tests and 594 assertions. Package typecheck, external-runtime build, scoped format/lint, and diff checks pass. This is a pure read projection; P10 still owns structural edit validation.

## P04

- [x] Project registered block attributes and children

**Depends:** [B01](02-react-blocks.md#b01), [B03](02-react-blocks.md#b03), [P02](03-agent-projection.md#p02), [P16](03-agent-projection.md#p16), [P17](03-agent-projection.md#p17)

**Primary files:** `fork:packages/sciobot/src/projection/blocks.ts`; `sibling blocks.test.ts`

**Read:** [S09](sources.md#s09), [S10](sources.md#s10), [S11](sources.md#s11)

**Change:** Serialize domain tags using existing Markdown rules. Map attribute spans to payload properties and child spans to native nodes.

**Close when:** Changing duration targets its property. Changing a child sentence targets that sentence without replacing its phase.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `projection/blocks.test.ts` passes eight imported-DOCX cases and 77 assertions. Nested title, phase, activity, source, recap, visual, and excerpt tags use application property remapping. Duration maps to its persistent payload property; child sentences retain native paragraph IDs and UTF-16 offsets. Explicit scio:label SDTs identify generated labels; unmarked first paragraphs remain content. Table constructs, inline source chips, escaping, empty containers, custom tag registration, and native Original/Proposed coordinates remain correct. Atomic fallback text stays outside native targets. Reads preserve all XML fingerprints, package revisions, and history. The private suite passes 105 tests and 722 assertions. Typecheck, external-runtime build, scoped Prettier, and oxlint pass. P05 retains complete opaque MDX AST handling; P10 retains structural edit validation.

## P05

- [x] Project opaque MDX without losing its AST

**Depends:** [B01](02-react-blocks.md#b01), [P04](03-agent-projection.md#p04), [P18](03-agent-projection.md#p18)

**Primary files:** `fork:packages/sciobot/src/projection/mdx.ts`; `sibling mdx.test.ts`

**Read:** [S11](sources.md#s11)

**Change:** Serialize complete stored MDX ASTs. Keep inline and block identities, nested children, and data-only payloads.

**Close when:** Every CUSTOM_ALLOWED_TAGS entry roundtrips its allowed attributes. Nested listening and suggestion payloads remain intact.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `capture-mdx-reference.mjs` executes the application serializer and records source hashes. All 21 CUSTOM_ALLOWED_TAGS entries preserve every allowed attribute in inline and block forms. Nested listening and suggestions retain their complete stored ASTs. `projection/mdx.test.ts` passes 52 tests and 869 assertions. Valid Unicode tag names and fragments also remain opaque. Imported DOCX projections retain block/control identities, JSON property paths, and surrounding native UTF-16 positions. Rich cell payloads use HTML table boundaries. Save/reopen retains every AST. Unknown metadata and prototype-shaped keys remain inert data. Expressions and spreads never execute. Invalid trees preserve XML fingerprints, package revision, selection, and history. The full private suite passes 168 tests and 1,685 assertions across 22 files. Package typecheck, external-runtime build, scoped Prettier, repository lint, and diff checks pass. This pure read change needs no collaboration mutation test. Renderers and legacy migration remain X09/X34-X36/I12.

## P06

- [x] Parse target Markdown without editing the document

**Depends:** [P02](03-agent-projection.md#p02), [P03](03-agent-projection.md#p03), [P04](03-agent-projection.md#p04), [P05](03-agent-projection.md#p05)

**Primary files:** `fork:packages/sciobot/src/projection/parse-target.ts`; `sibling parse-target.test.ts`

**Read:** [S03](sources.md#s03), [S11](sources.md#s11), [S14](sources.md#s14)

**Change:** Build a normalized target projection using the established Markdown and MDX rules. Keep normalization separate from canonical document writes.

**Close when:** C01 targets parse deterministically. Parse failure leaves document bytes, revision, selection, and history unchanged.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `capture-target-reference.mjs` executes the production application parser plugins and records source hashes. Twenty-four target parses match captured normalized ASTs, including both sides of all seven C01 cases. Three malformed MDX inputs retain captured syntax diagnostics and available locations. `projection/parse-target.test.ts` passes 33 tests and 226 assertions. Math/GFM/MDX syntax, line-break and ordered-list normalization, mentions, gallery grouping, inline images, code/math fences, nested opaque payloads, and rich Word table cells retain explicit expected semantics. AST positions address normalized Markdown. Expressions remain inert source, including BigInt/RegExp syntax. Immutable JSON targets exclude compiler ESTree caches and native document references. Input/node/depth limits produce no partial target. Success and failures preserve live DOCX bytes, XML fingerprints, package revision, selection, and existing undo/redo history. The full private suite passes 201 tests and 1,911 assertions. Typecheck, external-runtime build, scoped formatting, repository lint, and diff checks pass. Typed mention cleanup removes its added lint warnings. The MIT mention notice travels in the private package. Domain planners interpret preserved tag attributes/children; P11 retains lesson shape validation and tool messages. This pure parser performs no editor mutation.

## P07

- [x] Implement paginated read_document results

**Depends:** [C02](01-contracts.md#c02), [P02](03-agent-projection.md#p02)

**Primary files:** `fork:packages/sciobot/src/tools/read.ts`; `sibling read.test.ts`

**Read:** [S02](sources.md#s02), [S05](sources.md#s05)

**Change:** Implement line offsets, limits, truncation, defaults, and plain copyable content using the projection.

**Close when:** C02 read examples match field-for-field. Reading a page does not change the full projection.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `bun test packages/sciobot/src/tools/read.test.ts` passes seven tests and 16 assertions. All four C02 read examples match field-for-field, including the effective 1,000-line client default and past-end behavior. Cases also check invalid/fractional inputs, empty content, trailing blank lines, and a real native projection page. Returned Markdown copies directly into the captured replacement matcher. Reading and caller result changes preserve the full projection. Package typecheck, build boundary, scoped Prettier/oxlint, and diff checks pass. This pure reader adds no document mutation; mounted FileExecutor routing remains I02/I06.

## P08

- [x] Implement search_document results

**Depends:** [C02](01-contracts.md#c02), [P02](03-agent-projection.md#p02)

**Primary files:** `fork:packages/sciobot/src/tools/search.ts`; `sibling search.test.ts`

**Read:** [S02](sources.md#s02), [S05](sources.md#s05)

**Change:** Implement literal and regex search, line locations, context, limits, and stable error results.

**Close when:** C02 search examples match. Invalid regex and empty patterns return errors without altering the document.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `tools/search.test.ts` passes eight tests and 16 assertions. All four C02 search examples match field-for-field, including invalid-regex and empty-pattern errors. Additional cases verify literal punctuation, case-insensitive Czech text, one result per matching line, limits and full counts, line context, global-regex resets, zero-width termination, and native projection immutability. The full private suite passes 87 tests and 556 assertions across 16 files. Package typecheck, external-runtime build, scoped formatting, and repository lint pass. A fixture expectation needed unknown typing because imported JSON widens numeric literals; runtime expectations remain unchanged. Wire normalization and mounted routing remain I02/I06.

## P09

- [x] Invalidate stale agent projections

**Depends:** [C07](01-contracts.md#c07), [P01](03-agent-projection.md#p01)

**Primary files:** `fork:packages/sciobot/src/projection/revision-cache.ts`; `sibling revision-cache.test.ts`

**Read:** [S02](sources.md#s02), [S03](sources.md#s03)

**Change:** Cache by document identity and canonical revision. Invalidate after typing, review decisions, undo, redo, and remote edits.

**Close when:** A manual edit between tool calls appears in the next read. A cache from another document never passes a revision check.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `revision-cache.test.ts` passes four cases and 36 assertions. Keys include document, path, kind, load epoch, package revision, and view. Buffered typing, undo/redo, native review acceptance/rejection, remote edits, deletion, reconnect, and saved convergence invalidate real editor reads. Failed builders never return old projections. The complete private package suite passes 67 tests and 469 assertions. Package typecheck, scoped format, and repository lint pass; lint retains unrelated baseline warnings. The runtime build contains ten private modules with native packages external.

## P10

- [x] Resolve matched patch spans to canonical targets

**Depends:** [C04](01-contracts.md#c04), [P02](03-agent-projection.md#p02), [P03](03-agent-projection.md#p03), [P04](03-agent-projection.md#p04), [P05](03-agent-projection.md#p05)

**Primary files:** `fork:packages/sciobot/src/projection/resolve-patch.ts`; `sibling resolve-patch.test.ts`

**Read:** [S02](sources.md#s02), [S06](sources.md#s06)

**Change:** Translate a matched Markdown span into mapped native ranges or attribute targets. Refuse ambiguous or unmapped boundaries before writes.

**Close when:** A unique match selects exact node IDs and offsets. A mixed protected boundary produces a typed failure with zero operations.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `projection/resolve-patch.test.ts` passes ten imported-DOCX cases and 90 assertions. Located matcher spans resolve immutable native paragraph IDs and UTF-16 offsets, persistent payload property paths, complete heading/empty-paragraph/table constructs, and container identities. Complete opaque MDX targets its stored AST property. P03 cell identities permit multi-paragraph ranges within one cell; cross-cell changes require structural plans. Stale document/epoch/revision tokens, mixed property boundaries, protected content and zero-width markers, invalid protection metadata, escape interiors, surrogate splits, and ambiguous identities fail without targets or operations. Every case preserves DOCX bytes, XML fingerprints, revision, selection, and history. The full private suite passes 168 tests and 1,685 assertions across 22 files. Package typecheck, external-runtime build, scoped Prettier/oxlint, and diff checks pass. This pure resolver performs no editor mutation. C04 retains matcher occurrence counts. P12 still owns discovery of native protected content; canonical application validation remains required.

## P11

- [x] Validate edited syntax and lesson schema

**Depends:** [C07](01-contracts.md#c07), [P06](03-agent-projection.md#p06)

**Primary files:** `fork:packages/sciobot/src/tools/validate-target.ts`; `sibling validate-target.test.ts`

**Read:** [S02](sources.md#s02), [S14](sources.md#s14)

**Change:** Reuse syntax diagnostics, tag-balance hints, and lesson schema rules. Keep incremental shape warnings distinct from full-set rejection.

**Close when:** Broken MDX fails atomically. An incremental shape warning preserves its edit. Invalid full lesson replacement fails.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `validateTargetMarkdown` uses the existing bounded parser and application-owned schema/feedback. Syntax failures return no target; incremental lesson schema failures retain the target with warnings; full lesson replacements reject invalid shape. Generic targets require syntax only. The source-executed oracle at `fixtures/target-validation.json` records application `da9081ad`, seven source hashes, 32 syntax/schema cases, and nine tag-balance cases. Validation/parser coverage comprises 80 cases, including exact C02 feedback, primary Acorn point locations, normalized rejected context, inert expressions, native no-mutation checks, retained warning targets, accessors, resource limits, and bounded hint scanning. The native oracle applies the retained warning target to obtain the requested replacement. Full private suite: 389 pass and the same two R29 failures across 391 tests, 40 files, and 5800 assertions. Private typecheck/build, scoped Prettier/oxlint, root lint, diff checks, and compatibility policy pass. The build verifies 53 private source modules and a React-free document entry. No-impact decision: `.collaboration/changes/sciobot-target-validation.json`. This pure preflight requires no shared mutation test. I03-I05 retain live executor wiring. See [implementation notes](evidence/implementation-notes.md).

## P12

- [x] Preserve unprojected DOCX content around patches

**Depends:** [P01](03-agent-projection.md#p01), [P10](03-agent-projection.md#p10)

**Primary files:** `fork:packages/sciobot/src/projection/protected-content.ts`; `sibling protected-content.test.ts`

**Read:** [S25](sources.md#s25)

**Change:** Keep opaque OOXML, fields, drawings, headers, footers, and section properties outside the writable projection. Retain their canonical references.

**Close when:** A body text patch preserves fingerprints of untouched parts and nodes. Unsupported edits identify their protected boundary.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `collectProtectedContent` discovers protected canonical nodes, immutable Markdown boundaries, other XML/binary parts, and OPC metadata. `checkProtectedContent` validates an immutable body-text preview before any live write. It preserves fingerprints, namespace bindings, parent ownership, and surviving sibling order. The rich DOCX fixture covers simple/complex fields, an image, unknown inline/block XML, foreign attributes, unknown properties, sections, background, headers, footers, notes, and binary parts. Six real native patches preserve protected content through undo/redo and save/reopen. Refusals identify protected nodes or parts; hidden fields and omitted atoms retain zero-width boundaries with editable adjacent text. Final focused protection/resolver command passes 20 cases and 274 assertions. Full private suite: 414 pass and the same two R29 failures across 416 tests, 42 files, and 6157 assertions; final focused checks add part-binding and invalid-bound cases. Private typecheck/build, scoped Prettier/oxlint, root lint, diff checks, and compatibility policy pass. The build verifies 57 private modules and a React-free document entry. No-impact decision: `.collaboration/changes/sciobot-protected-content.json`. This pure preflight has no commit callback and needs no new shared mutation test. A01/A02 and I03-I05 retain live transaction wiring. See [implementation notes](evidence/implementation-notes.md).

## P13

- [x] Return echoes from the landed projection

**Depends:** [P06](03-agent-projection.md#p06), [P10](03-agent-projection.md#p10), [P11](03-agent-projection.md#p11)

**Primary files:** `fork:packages/sciobot/src/tools/edit-echo.ts`; `sibling edit-echo.test.ts`

**Read:** [S02](sources.md#s02), [S05](sources.md#s05)

**Change:** Build result excerpts from the normalized document state after application. Do not echo an unpersisted model string.

**Close when:** Copying the first successful echo into a second patch succeeds after list and whitespace normalization.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `buildEditEcho` reads before/landed working projections and a current revision token. It rejects stale, foreign, future, inconsistent, and non-working snapshots. Its detached excerpt contains only exact landed Markdown; metadata records line bounds, revision, and truncation. The source-executed oracle records 13 application echo cases, application `da9081ad`, and three source hashes. Focused echo tests pass 15 cases and 180 assertions. A real native regression validates a requested target with repeated list markers and extra whitespace, applies the first paragraph plan, reads the normalized working projection, copies its landed `2. Second` line, matches/resolves it through C04/P10, and applies a second native plan to obtain `2. Final`. Parsed semantic targets match after normalization. Native paragraph IDs, read fingerprints, and history stay intact. A later revision rejects the stale first snapshot. Full private suite: 404 pass and the same two R29 failures across 406 tests, 41 files, and 5980 assertions. Private typecheck/build, scoped Prettier/oxlint, root lint, diff checks, and compatibility policy pass. Build checks verify 54 private source modules and separate React/document entries. No-impact decision: `.collaboration/changes/sciobot-edit-echo.json`. This pure formatter requires no shared mutation scenario. I03-I05 retain live bridge wiring. See [implementation notes](evidence/implementation-notes.md).

## P14

- [x] Expose complete native Markdown source maps

**Depends:** [P01](03-agent-projection.md#p01)

**Primary files:** `fork:packages/docx-to-markdown/src/{markdown,markdown-inline,markdown-source-map,markdown-types,index}.ts`; public API snapshot and focused tests

**Read:** [S26](sources.md#s26), [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Add opt-in source maps for ordinary logical Markdown. Retain sparse exact boundaries through escaping and generated syntax.

**Close when:** Unreviewed paragraphs, formatted text, escaped punctuation, emoji, and table cells retain original UTF-16 addresses. Plain text uses compact mappings. Existing output stays unchanged without the option.

**Check:** Run native converter regressions, typecheck, build, public API checks, and scoped formatting/lint. This is a read-only conversion change.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `bun test packages/docx-to-markdown/test` passes 204 tests and 1,164 assertions across 29 files. The six logical source-map cases and the transform case verify unreviewed text, escaping, UTF-16 addresses, marks, hyperlinks, table cells, empty paragraphs, generated syntax, immutable JSON output, and compact plain text. All three export entry points support the option. Default output remains equal. Converter typecheck, build, API extraction/check, scoped Prettier, repository lint, and diff checks pass. API extraction retains preexisting package warnings. This read-only conversion change requires no live collaboration mutation test. Public API snapshots, documentation, the capability ledger, and the minor changeset are updated.

## P15

- [x] Preserve native quote styles in mapped Markdown

**Depends:** [P14](03-agent-projection.md#p14)

**Primary files:** `fork:packages/docx-to-markdown/src/markdown.ts`; focused converter tests and documentation

**Read:** [S26](sources.md#s26), [S13](sources.md#s13)

**Change:** Render the built-in Word Quote and IntenseQuote paragraph styles as Markdown blockquotes. Preserve complete source maps through generated quote prefixes.

**Close when:** Imported quote paragraphs preserve their text, marks, and native breaks. Quote prefixes have no native text offsets. Other paragraph styles keep their existing output.

**Check:** Run focused native conversion tests, package typecheck, build, formatting, and lint. This is a read-only projection change.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Built-in Quote and IntenseQuote styles emit mapped blockquotes outside GFM table cells. Custom styles keep their previous output. `bun test` over markdown, inline, records, source-map, logical-source-map, and review-binding integration files passes 70 tests and 312 assertions. The imported quote case checks bold text, native hard breaks, empty quotes, both built-in styles, exact source positions after generated prefixes, and an unchanged custom style. Converter typecheck, build, scoped Prettier, and oxlint pass. Public docs, capability notes, and changeset describe the table-cell and custom-style limits. This closes the pure converter prerequisite, not full application quote support.

## P16

- [x] Map complete Markdown paragraph constructs

**Depends:** [P14](03-agent-projection.md#p14)

**Primary files:** `fork:packages/docx-to-markdown/src/{markdown,markdown-source-map,markdown-source-types}.ts`; focused source-map tests and API snapshot

**Read:** [S26](sources.md#s26), [S25](sources.md#s25)

**Change:** Include generated heading, list, quote, and mark syntax in paragraph construct extents. Shift these extents through native table and note rendering.

**Close when:** A native paragraph has exact logical Markdown construct bounds, including empty paragraphs. Wrapping a bold paragraph includes both delimiters. Default Markdown output stays unchanged.

**Check:** Run converter regressions, typecheck, build, API checks, formatting, and lint. This is a pure read projection.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Complete paragraph construct bounds include headings, list markers, emphasis delimiters, quote prefixes, and empty paragraphs. Bounds survive concatenation, nested table HTML, cell escaping, and note definition indentation. `bun test packages/docx-to-markdown/test` passes 207 tests and 1,186 assertions across 29 files. The private suite passes 92 tests and 596 assertions, including both captured C01 list variants with explicit construct bounds. Default Markdown output stays unchanged. Both package typechecks, native/private builds, API extraction/check, scoped formatting, repository lint, and diff checks pass. Public docs, API snapshot, and the existing minor changeset cover this pure source-map extension.

## P17

- [x] Map complete Markdown table constructs

**Depends:** [P16](03-agent-projection.md#p16)

**Primary files:** `fork:packages/docx-to-markdown/src/{markdown-logical,markdown-source-map,markdown-source-types,index}.ts`; source-map tests and public API snapshot

**Read:** [S26](sources.md#s26), [S25](sources.md#s25)

**Change:** Include canonical table IDs and complete construct extents in opt-in source maps. Preserve them through cell escaping and note indentation.

**Close when:** Bounds include GFM bars and divider rows, nested HTML table tags, and tables containing empty cells. Repeated page fragments map to one logical table. Default output stays unchanged.

**Check:** Run converter regressions, typecheck, build, API checks, formatting, and lint. This is a pure read projection.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Opt-in source maps expose canonical table IDs with full GFM or nested HTML extents. Empty-cell tables retain bars and divider rows; nested tables retain both HTML tags. Bounds shift through note indentation and preserve story scopes. A real multi-page private fixture produces one logical entry despite repeated headers and page fragments. The full converter suite passes 208 tests and 1,199 assertions across 29 files. The private suite passes 97 tests and 645 assertions. Both package typechecks, native/private builds, API extraction/check, scoped formatting, repository lint, and diff checks pass. Default output remains equal. Public docs, capability notes, API snapshots, and the source-map changeset cover this pure projection change.

## P18

- [x] Expose HTML table projection for rich cell content

**Depends:** [P17](03-agent-projection.md#p17)

**Primary files:** `fork:packages/docx-to-markdown/src/{markdown,markdown-logical,markdown-types,index,media-types}.ts`; `fork:packages/sciobot/src/projection/markdown.ts`; converter tests and API snapshot

**Read:** [S26](sources.md#s26), [S11](sources.md#s11)

**Change:** Add an explicit HTML table rendering option using the existing nested-table renderer. Thread it through all Markdown entry points and preserve source maps.

**Close when:** HTML mode retains row order, cell values, marks, and canonical mappings. Multiline MDX can occupy a cell without corrupting the table grammar. Default GFM output stays unchanged.

**Check:** Run converter regressions, parser examples, package typechecks/builds, API checks, formatting, and lint. This is a pure read projection.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. All three Markdown entry points accept tableSyntax: html. The existing table renderer adds MDX block boundaries and preserves row order, header cells, marks, native breaks, paragraphs, and canonical source maps. Default and explicit GFM output remain equal. The full converter suite passes 210 tests and 1,211 assertions across 30 files. The private suite passes 106 tests and 726 assertions, including an actual MDX parse of multiline listening content in a cell. Both package typechecks, native/private builds, API extraction/check, scoped Prettier, repository lint, and diff checks pass. Public docs explain that cell content remains Markdown and requires an MDX-aware consumer. The additive API has a minor changeset. No document mutation or collaboration format changes occur.

## P19

- [x] Expose canonical tracked replacement landing

**Depends:** [C03](01-contracts.md#c03)

**Primary files:** `fork:packages/core/src/store/store/tracked-replacement-landing.ts`; native store export, editor delegation, tests, API snapshot, and changeset

**Read:** [S25](sources.md#s25)

**Change:** Expose the existing pure replacement landing helper through the canonical store entry point. Keep one implementation for editor and application callers.

**Close when:** Plain, own-insertion, foreign-insertion, and adjacent-deletion cases retain native model offsets. Reads leave the canonical tree unchanged. Existing tracked replacement regressions still pass.

**Check:** Run focused native regressions, typecheck, build, API extraction/check, formatting, lint, and compatibility policy.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. The canonical store now exports the unchanged trackedReplacementLanding implementation. The editor delegates to that same function. Six public-entry cases verify plain text, own/foreign insertions, adjacent current/old deletions, and insertion inside a strike without changing canonical fingerprints. The focused tracked store/editor suite passes 86 tests and 222 assertions; lane and ProseMirror isolation checks pass 48 tests and 169 assertions. Core typecheck/build, API extraction/check, scoped formatting, repository lint, compatibility policy, and diff checks pass. The store API snapshot adds only the helper signature. Architecture docs and the capability note identify this pure query. The additive API has a minor core changeset and a no-impact compatibility decision. This exposes existing read logic and changes no mutation behavior or shared format.

## P20

- [x] Resolve native selections to chat quote ranges

**Depends:** [P02](03-agent-projection.md#p02), [P03](03-agent-projection.md#p03), [P04](03-agent-projection.md#p04), [R02](05-suggestions.md#r02)

**Primary files:** `fork:packages/sciobot/src/projection/quote-selection.ts`; `sibling quote-selection.test.ts`

**Read:** [S27](sources.md#s27), [S02](sources.md#s02)

**Change:** Map native selection endpoints to the working projection. Return quote text, inclusive line numbers, containsMdx, and resolution status. Preserve the existing fallback for removed content.

**Close when:** Selected pending insertion text matches read_document lines. Nested activity and cell selections return correct text and lines. Unmapped ranges remain unresolved.

**Check:** Use explicit quote results from the source resolver and native fixtures. Cover reversed endpoints, escapes, removed content, Unicode, and stale revisions. Reads must preserve history.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, application reference `da9081a`, and local working tree. `bun scripts/sciobot-port/capture-quote-reference.mjs` captures 15 actual source quote outcomes with source and input hashes in `packages/sciobot/src/__tests__/fixtures/quote-reference.json`. The pure resolver supports native text and block endpoints, inclusive working-view lines, nested activity/cell selections, atomic MDX and labels, reversed ranges, Unicode, escapes, manual breaks, pending insertions, explicit live deletion fallback, and unresolved targets. Fourteen focused quote cases pass. The related projection/review subset passes 82 tests across five files. Existing undo/redo, selection, native fingerprints, revisions, pending state, and saved/reopened reads remain intact. The full private suite reports 332 pass and the same two R29 concurrency failures across 334 tests and 35 files. Private typecheck, 42-module external-runtime build, scoped Prettier, root lint, diff checks, and compatibility policy pass. Decision: `.collaboration/changes/sciobot-quote-selection.json`; no shared format changes. Native quote text follows landed serializer output, including deletion markup in the live fallback. No editor writes or production transport dependencies enter this reader. U15/U16 retain actual chat actions. See [implementation notes](evidence/implementation-notes.md).

## P21

- [x] Project native checklist state as mapped task syntax

**Depends:** [P02](03-agent-projection.md#p02), [P06](03-agent-projection.md#p06), [D06](04-edit-planner.md#d06)

**Primary files:** `fork:packages/sciobot/src/projection/lists.ts`; `sibling lists.test.ts`

**Read:** [S11](sources.md#s11), [S13](sources.md#s13), [S25](sources.md#s25)

**Change:** Emit `- [ ]` and `- [x]` for application-owned native checklist markers. Map their control identity, state, nesting, and text offsets. Keep ordinary imported form controls unchanged.

**Close when:** Native checked and unchecked items parse as GFM task items. Editing their text retains the marker. A state-only change targets its control. Original and Proposed views preserve exact source addresses.

**Check:** Add focused native projection tests with explicit expected Markdown, ASTs, and UTF-16 boundaries. Cover nesting, manual breaks, inline symbols, pending wrappers, cells, and save/reopen. Reads must preserve history.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `projectLists` derives GFM tasks from published native control visibility/ranges and detached semantic records. It preserves control/paragraph identity, checked state, nesting, and original UTF-16 text offsets. `resolvePatchSpan` returns typed checklist-state addresses; explicit read-only regions, native locks/binding, and stale tokens refuse. Domain composition remaps tasks and removes atomic fallback metadata. Working/Committed readers include the layer. Cells use HTML boundaries. Empty tasks receive a generated NBSP and a native insertion anchor; D15 must exclude that placeholder from authored content. Thirteen task cases cover parsing, marks, manual breaks, escapes, symbols, empty edges, native text and state-only oracles, undo/redo, save/reopen, Original/Proposed controls, All Markup, bound/malformed ownership, and unchanged bytes/revision/selection/history. Ordinary form output and source maps stay equal. Final task/domain/protection/resolver/review subset passes 46 cases and 494 assertions. Full private suite: 435 pass and the same two R29 failures across 437 tests, 44 files, and 6347 assertions; the final subset adds a no-task fast-path equivalence case. Private typecheck/build, scoped Prettier/oxlint, root lint, diff checks, and compatibility policy pass. The build verifies 61 private modules with one external core. No-impact decision: `.collaboration/changes/sciobot-checklist-projection.json`. This pure reader/target resolver changes no native or shared operations and needs no new collaboration mutation scenario. A01/R30 retain guarded direct and tracked execution. See [implementation notes](evidence/implementation-notes.md).

## P22

- [x] Preserve native bookmark links and tooltips in Markdown

**Depends:** [P02](03-agent-projection.md#p02), [P14](03-agent-projection.md#p14)

**Primary files:** `fork:packages/docx-to-markdown/src/markdown.ts`; `sibling markdown-semantic-policy.ts`; `fork:packages/docx-to-markdown/test/markdown-hyperlinks.test.ts`

**Read:** [S03](sources.md#s03), [S25](sources.md#s25)

**Change:** Emit sanitized internal bookmark destinations and authored hyperlink tooltips. Preserve mapped label offsets. Keep unsafe and unresolved links inert.

**Close when:** Internal and external links retain exact destinations, titles, and label marks through logical export, page export, and save/reopen. Missing bookmarks retain their authored reference. Hostile title characters cannot create extra Markdown elements.

**Check:** Add independent Markdown and parsed HTML assertions. Verify UTF-16 source maps, title escaping, unchanged source fingerprints, and R05's new-paragraph link regression. This read-only projection needs no new collaboration mutation scenario.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78` plus local changes. The record-only writer preserves sanitized bookmark destinations and escaped tooltips in logical/page Markdown. Three new cases use independent Markdown/HTML assertions for present/missing targets, marked Unicode labels, exact source bounds, hostile/empty/absent titles, unsafe/unresolved relationships, immutable source fingerprints, and save/reopen. `bun test packages/docx-to-markdown/test` passes 213 tests and 1241 assertions across 31 files. R05's 29 native editor cases and 222 assertions pass, including attributed new links and saved review decisions. The comprehensive fixture gains exactly three bookmark wrappers; removing only those wrappers restores the previous length and hash. Pagination stays unchanged. Converter/private typecheck pass 1388/1650 files; both builds pass. The private build verifies 104 modules and one external core peer. Converter API check scans two entries without errors or drift; its warning baseline stays unchanged. Scoped formatting/lint, repository lint, compatibility policy, source coverage, and backlog checks pass. Changeset: `.changeset/markdown-bookmark-links-and-titles.md`. Decision: `.collaboration/changes/markdown-bookmark-links-and-titles.json`. No shared fields or native writes change. This read-only feature requires no new collaboration mutation test. R05/R37 retain the separate failing offline forward-join cases; R38/N18 retain inline code. See [implementation notes](evidence/implementation-notes.md).

## P23

- [x] Project consecutive native quote paragraphs as one Markdown blockquote

**Depends:** [P15](03-agent-projection.md#p15)

**Primary files:** `fork:packages/docx-to-markdown/src/markdown.ts`; `focused converter quote test`

**Read:** [S25](sources.md#s25)

**Change:** Join consecutive Quote or IntenseQuote paragraphs of one style into one blockquote with `>` separator lines, shifting source maps together. Keep separate quotes when the style changes or another block intervenes.

**Close when:** Two consecutive Quote paragraphs project as one two-paragraph blockquote with exact source maps, and a multi-paragraph inserted quote then lands its target through the D28 materializer.

**Check:** Add focused converter tests and rerun the D28 fragment tests. Update API snapshots and add a changeset if the public converter output changes.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `renderLogicalBlocks` in `packages/docx-to-markdown/src/markdown.ts` now groups consecutive top-level paragraphs of one quote style (Quote or IntenseQuote, outside table cells) and joins them with a generated `

> `separator inside one blockquote. A style change or any intervening block keeps separate quotes. Existing per-paragraph quote prefixes and`concatMarkdown`shift sources, paragraph extents, and tables together. New converter case in`test/markdown-logical-source-map.test.ts`expects the exact output`> One
>
> Two

> Loud

Between

> Three
>
> Four`(an empty Quote paragraph stays a bare`>`line), exact paragraph extents, exact UTF-16 offsets for six positions, and no source slice over the generated separator. The converter suite passes all 208 cases apart from six environment cases: four browser parity/image cases pass with the repository DOM preload (12 pass), and the two 500-page heap cases time out under the shared machine load and are unrelated to paragraph joining. The planner's one-paragraph quote refusal is lifted.`fragment-seeds.test.ts`adds a two-paragraph inserted quote with marks and a link that plans one sibling insertion with the document Quote style, lands through the live direct executor as two consecutive Quote paragraphs, replans empty, undoes in one step, and reopens to the target AST. Fragment, composer, C01 matrix, and projection regressions pass 208 cases. Package typecheck (1393 files), API check (0 errors, unchanged warning counts), scoped Prettier/oxlint, and compatibility (75 decisions) pass. Changeset`.changeset/markdown-joined-quote-paragraphs.md`(patch); decision`.collaboration/changes/markdown-joined-quote-paragraphs.json`. Found while testing: editing text inside an existing quote refuses as a composite owner even for a single quote paragraph; D39 owns that gap.
