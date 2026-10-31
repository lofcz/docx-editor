# Native code blocks

X16 and D34 evidence. Fork `62c7e8ef`, application branch `claude/docx-native-cutover`.

## Representation

A `code_block` is an atomic block. Its payload holds `language` and the exact `value`. The block control holds one Word fallback paragraph for each source line. Each paragraph uses the inline code character style, and tabs are `w:tab` elements. The Markdown projection reads the payload, so tabs, blank lines, and repeated spaces stay exact.

## Writes

- `updateAtomicBlock` (fork `blocks/update-atomic.ts`) replaces the payload and regenerates the Word fallback in one native undo unit. The block keeps its ID and control. Suggest mode and view mode refuse with `unsupported-operation`.
- Direct agent edits remove application blocks. `siblingRemoval` (fork `ai/direct-block-removal.ts`) accepts block controls, and the executor retires their payload records and the records of nested blocks. An agent code change plans as one removal and one insertion with a new block ID.

## Application

`src/components/workspace/docx/blocks/code.tsx` registers the `code_block` view. The native fallback paints the code on the page. The view adds a language selector, an edit dialog, and a copy button. The dialog textarea inserts a tab for Tab and keeps Enter and paste as plain text. Block views do not observe editing-mode changes. Therefore, the dialog always opens, and a refused save shows the engine reason.

Difference from Plate: pages show the code in the document code style without syntax colors. Word export and print show the same text.

## Checks

- Fork `src/blocks/update-atomic.test.ts`: 3 cases. Exact whitespace, Word lines, language change, save and reopen, payload store, projection, undo, redo, suggest and view refusal, and invalid input. Two real collaboration sessions in both receiver orders: peer typing beside the update, undo and redo with the peer text kept, concurrent deletion of the block, and equal saved digests.
- Fork `src/ai/direct-block-removal.test.ts`: 3 cases. Direct removal of a code block, a `lesson_visual`, and a `callout`, with payload records retired and save and reopen.
- Fork scoped suites `ai`, `diff`, `blocks`: 1396 pass and 4 fail. Three failures are on the committed baseline (`diff/inserted-atoms.test.ts` 2, `ai/list-seed-identities.test.ts` 1). `diff/inserted-atoms-collaboration.test.ts` timed out under load and passes alone (5 pass).
- Application `scripts/docx-native/harness/code-blocks.test.ts`: 3 browser cases. Dialog input with Enter, Tab, and clipboard paste; language change; Markdown export; saved `w:tab` and repeated spaces; native typing after the dialog; undo and redo; suggest-mode refusal; two sessions in both client orders with peer typing, undo, redo, equal saved digests, and equal Markdown.
