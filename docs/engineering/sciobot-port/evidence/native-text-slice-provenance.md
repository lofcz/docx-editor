# Native text-slice provenance

This is an integration checkpoint, not D30 closure. Runtime commits are `207fd2c4`, `ffadfbef`, and `35fe2455`. The branch remains isolated from the primary fork branch.

Native text descriptors carry optional ordered `sourceText` pieces. A source window names its observed leaf, UTF-16 offset and length, and the journal boundary before the split. A fresh piece names only its UTF-16 length. Coalescing and later slicing compose these pieces without guessing character offsets. Empty imported text contributes no nonexistent source window. Boundary spaces keep the native serializer's existing preservation rule.

The capture frame owns these annotations. Other transactions and failed previews cannot lend provenance to a later write. Nested pieces freeze with the descriptor and survive `structuredClone`. Retained tails preserve internal source windows. Cross-paragraph splits mint separate native identities; within-paragraph controls retain their trailing addresses. Tracked range slices supply their actual source offsets. Collaboration translates nested source addresses before minting the target.

The evaluation transport resolves source windows before replay changes their leaves. Its native view supplies actual shared character identities, including following text from joins. Scratch edits retain those identities through every fragment. Fresh characters receive no old identity. Alignment uses exact identities and whole code points. Move matching follows those identities; equal new labels cannot borrow old characters. Repeated split journals retain subsequent peer text and formatting. Invalid windows and malformed recipes refuse before shared writes.

This metadata is provider infrastructure. It adds no Office.js document member and no stored room field. The private production adapter imports no EigenPal Pro License runtime. The transport changes remain test-only for that adapter. A minor core changeset covers the additive descriptor. The compatible-writer decision retains the existing shared representation.

## Verification

At `207fd2c4`, the wider native check passes 1,054 cases across 43 files, with 10,784 assertions. The private diff, AI and composer check passes 837 cases; one case exceeds the unchanged five-second limit. Its complete five-case file passes separately with 1,613 assertions. The wider transport check passes 239 cases; three performance cases time out. The complete follow-cost file passes separately, and the deletion-index file passes alone in 4.68 seconds. No assertion or timeout changes.

The 26 additional tracked/format files pass 364 cases and expose one acceptance failure. It expects `yworldNext` after accepting a split beside another author's deleted paragraph mark, but observes separate `yworld` and `Next` paragraphs. The unchanged case also fails at prior revision `15b9dcfe`. This inherited suggestion defect remains open; the wider checks are not an unconditional all-green result.

The exact middle-fragment matrix passes 32 cases and 768 assertions. It covers client orders, reconnect receivers, prefixes, equal labels, repeated words, supplementary characters, actor and peer histories, native Word fingerprints, and DOCX/room reopening. At `ffadfbef`, all 56 original code-control fixtures pass against fresh compiled core, SDK and exporter artifacts, with 1,267 assertions. Run `bun packages/sciobot/scripts/check-code-inline-controls-compiled.mjs` after building core, fonts, exporter and private SDK. The runner changes imports only.

The follow-up malformed-recipe check passes 19 cases. Primitive completeness and range rewriting pass 146 cases. At `35fe2455`, provenance and primitive journal checks pass 165 cases with 742 assertions, including empty imported text and boundary-space preservation. The original seeded typing checks pass unchanged, using the shortcut for 164/167, 153/160 and 167/178 edits. The historical peer-born-cell probe also passes both row and column cases.

A separate process runs the prior source reader/writer at `15b9dcfe`, without the descriptor implementation. It opens the corrected room, edits retained peer text, and performs undo and redo. Both current editors consume its updates. The original writer redoes its atom birth while preserving those edits. All 16 native snapshots match independent native-operation oracles. No `sourceText` field enters shared state. Run `bun --preload ./scripts/test/preload-dom.ts packages/sciobot/scripts/check-text-slice-reader-compatibility.ts <prior-checkout>`.

Core types pass across 1,596 files. Transport types pass across 2,301 ordinary files and 1,853 Vue files. Private types pass across 2,183 files. Core, fonts, exporter and private SDK builds pass at `ffadfbef`; core stays external across 256 private modules. Core API extraction and checking pass with the additive descriptor snapshot. Root formatting, lint and compatibility policy pass; policy covers 181 decisions and two inherited format fields. The boundary-space follow-up is source-tested separately from the compiled `ffadfbef` artifacts.


## Joined-paragraph read defect and repair

The unchanged failing case also fails at `15b9dcfe`. Its primitive journal accepts the deleted mark by adding `urn:docx-editor.dev:review:retired-paragraph:1` / `joined` recovery metadata. It preserves both physical paragraphs and their identities. Word export correctly folds those records into `yworldNext`. Automation `storyReadsOver` in `packages/core/src/automation/reads.ts` still enumerates them separately, and the live paragraph query returns `yworld` and `Next`.

The next repair must expose logical joined paragraphs while preserving canonical physical addresses for editing. Paragraph ranges, search and subsequent writes need mappings through every physical fragment. A text-only concatenation would give those writes incorrect offsets. Keep the existing acceptance oracle and native recovery metadata. No diagnostic instrumentation remains in runtime or test sources.


Runtime `9cfd1f3d` repairs these logical reads and fragment-mapped writes. Shared history checks pass in source and compiled packages. Runtime `14d9787f` also repairs the inherited DOM-free identity compilation failure. See [joined paragraph verification](joined-paragraph-automation.md). Two inherited tracked-row fingerprint failures remain open.
