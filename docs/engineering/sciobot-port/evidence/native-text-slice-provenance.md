# Native text-slice provenance

This is an integration checkpoint, not D30 closure.

Native text descriptors can carry an optional ordered `sourceText` pieces. Each source window names the observed source leaf, UTF-16 offset and length, and the journal effect boundary before the split. The capture frame owns these annotations. Unobserved previews and other transactions cannot lend them to a later write. Nested provenance freezes with the descriptor and survives `structuredClone`.

The evaluation transport resolves each source window before replay changes that source. Its native view supplies actual shared character identities, including following text from joins. Scratch text edits retain those identities through every source fragment. New characters receive no old identity. Alignment uses exact shared identities and whole code points. Cross-paragraph move matching also uses exact source identities; equal new labels cannot borrow deleted characters.

The metadata is local provider infrastructure. It does not add an Office.js document method or a stored room field. The private production adapter still imports no EigenPal Pro License runtime. The evaluation transport changes stay test-only for that adapter.

Focused evidence: the expanded middle-fragment matrix passes 32 cases and 768 assertions. It covers both client orders, reconnect receivers, optional prefixes, equal atom labels, repeated words, supplementary characters, actor and peer histories, exact native Word fingerprints, and DOCX/room reopening. The original eight expectations remain in the default test suite. Source alignment, move matching, and existing split/source cases pass 69 checks and 1,032 assertions. Native journal capture passes 18 cases and 90 assertions, including nested freezing, cloning and rollback.

A separate process runs the prior source reader/writer at `15b9dcfe`, without the new descriptor or source implementation. It opens the corrected room, types inside retained peer text, and performs its own undo and redo. Both current editors consume its updates. The original writer can redo its atom birth while preserving those edits. All 16 native snapshots match independent native-operation oracles. No `sourceText` field enters shared state. Run `bun --preload ./scripts/test/preload-dom.ts packages/sciobot/scripts/check-text-slice-reader-compatibility.ts <prior-checkout>`.

Core types pass across 1,595 files. Both evaluation transport programs pass: 2,300 ordinary files and 1,852 Vue files. Broader integration, fresh compiled fixtures, API extraction, root formatting, lint and final compatibility checks remain pending. The new optional provider descriptor requires a minor package changeset. Its corrected writes use the existing shared representation; compatibility uses the compatible-writer decision.


Integration repair: tracked range slices now supply their actual offsets. Coalescing retains ordered source windows and fresh-text lengths, including subsequent slices across their boundaries. Retained tails preserve their pre-edit windows internally. Paragraph splits mint separate leaf identities; within-paragraph controls retain their tail addresses. Local-to-shared identity translation also translates nested source windows before minting the target.

The repaired checkpoint passes 146 range-rewrite and primitive-completeness cases (676 assertions), 91 exact history cases (3,058 assertions), three seeded typing cases (720 assertions), and six provenance/identity cases (20 assertions). Typing uses the shortcut for 164/167, 153/160 and 167/178 edits, within unchanged limits. Root lint passes with existing warnings. Core types pass across 1,595 files. Final wider suites and compiled verification still remain pending.
