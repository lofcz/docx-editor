# Native solution reveal controls

Application: `fe157dbd7beb83d80769eff8ccca74c520367909`. Runtime remains `f5bb40c9c703e61abef8cec05021d02eda715f78`. Branch: `codex/react-block-proof`.

## Behavior

The native workspace registers `solution` with native content retained. Localized reveal controls use the existing solution label, key icon, and emerald accent. Each page fragment has a control. Answers hide on screen inside their existing Word pages; hiding preserves native pagination and caret coordinates. Body nodes stay mounted. No saved answer content becomes React state.

An editor-scoped store shares one reveal value across a solution's fragments. Separate editors retain separate preferences. Subscriptions and native focus/input listeners are released when their roots unmount. Reopening creates a new editor with answers hidden. There are no payload or shared-state writes for reveal preferences.

The existing `scrollbot:reveal-target` event opens collapsed ancestors synchronously. Native selection opens the answers it reaches. Input capture commits visibility before the native input handler, including a focused answer-only document with an unchanged initial caret. The real browser probe observes visible answer geometry during `beforeinput`, before canonical text editing.

Screen-only scoped styles hide answers. Browser print retains every answer, including nested solutions. The initial utility-class approach failed because the engine's scoped `hidden` rule overrode the print utility. The final component uses dedicated screen-only data selectors. Markdown and DOCX export always read the complete canonical content.

The original `SolutionElementStatic` and `SolutionElementDocx` render their full children. The fixture retains that same content: styled answer text on the first page, another answer paragraph after a page break, a nested solution, and an external reference. Hidden and revealed views retain identical native fingerprints and saved semantic digests. The Markdown assertions require both solution containers, all answer text, metadata, and the reference URL.

## Verification

From the isolated application worktree:

```sh
DOCX_NATIVE_SNAPSHOT=/home/lofcz/.cache/sciobot-docx-native/port-current-integration bun test --timeout 600000 scripts/docx-native/harness/solutions.test.ts scripts/docx-native/harness/containers.test.ts
bun check -p tsconfig.json
bun check -p scripts/docx-native/tsconfig.json
```

All eight browser cases pass with 137 assertions. Five solution cases cover repeated reveal/hide from continuation pages, unchanged canonical state/revision/history, print, nested synchronous AI reveal, initial-caret input, keyboard navigation, both collaboration client orders, disconnected edits, independent peer visibility, deletion during peer editing, both actors' histories, exact caret geometry, saved convergence, and hidden reopening. Three callout/story cases retain the existing container behavior.

Logs: `/tmp/x15-final-verified.log`, `/tmp/x15-types-verified.log`, `/tmp/x15-harness-verified.log`, and `/tmp/x15-lint-verified.log`. Frontend types pass across 6,840 files. Harness types pass across 2,414 files. Scoped formatting, lint, and diff checks pass. Original test deadlines stay unchanged. The final input test retains its visibility-before-write assertion; the earlier focus-only wait could not detect an already focused native page.

## Compatibility

This is application presentation over existing native slots and editor events. It changes no shared records, canonical operations, runtime dependencies, public API, or collaboration format. The installed immutable runtime and its five archives remain unchanged. No published package Changeset or room migration is required for this leaf.

The evaluated collaboration transport remains test-only under the EigenPal Pro License. Production solution components import the existing core/private SDK and application helpers. They import no Pro or Plate runtime.
