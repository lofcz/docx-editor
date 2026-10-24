# Captured checklist indentation

N08 remains open. This checkpoint closes no leaf.

## Source contract

The frozen `list.todo.on` capture calls `indentTodo` on a plain paragraph. Its result has `indent: 1`, `listStyleType: 'disc'`, and `checked: false`. The frozen `list.todo.off` capture starts with `indent: 1` and `listStyleType: 'todo'`. Its result increases indentation to 2, changes the style to `disc`, and retains `checked: false`.

The source transform sets checked state and disc style while increasing indentation. The production `BlockList` renderer selects its checkbox renderer only for `listStyleType: 'todo'`. The checked state is dormant on the resulting disc item. No reference or application source changes.

## Implementation

`indentManualChecklist` implements this action separately from `setManualList(editor, 'todo', ...)`. The existing action still creates an actual Word checkbox. The captured action uses native disc numbering and the appropriate list level. Existing paragraph indentation contributes to its starting level. Native depth limits stay unchanged.

A `urn:sciobot:docx:lists:1` property node retains dormant checked state. The node carries `mc:Ignorable` and `mc:PreserveElements` hints. It does not paint a checkbox or alter document text. Ordinary list removal and real checklist conversion clear active dormant state. Imported metadata without membership stays lossless but inactive. It does not create an unreviewable metadata-only suggestion. Native property-change snapshots retain the old state for review. Malformed, duplicate, or future owned fields refuse before canonical changes.

The source reader includes dormant metadata in its fingerprint only when present. Absent metadata retains its previous fingerprint. Optional private plan data permits dormant state with a new counter instance. The source action always creates that instance, so its native property change owns the edit. Metadata-only writes to an existing instance do not claim unsupported suggestion behavior.

This evidence covers fork save/reopen. It does not prove preservation after an external Word editor changes the document. MCE preservation hints do not require an Office application to retain the extension. See [Microsoft's OOXML extensibility constraints, section 10.1.1](https://download.microsoft.com/download/e/1/4/e14fb96f-83b8-4a2a-84db-7fa8acbe061a/Office%20Open%20XML%20Part%201%20-%20Fundamentals.pdf).

## Verification

The local cases compare both frozen source captures with native list levels, dormant checked state, disc markers, selection, text, and static complete Word output. Both review decisions preserve exact native histories and saved reopening. When an original paragraph has no property holder, Word's rejection snapshot restores an empty `pPr`. The rejection oracle includes that native holder; it does not normalize it away.

Twelve real-session cases cover both client orders, both source states, concurrent text in the affected paragraph, direct edits, acceptance, rejection, and independent histories. They compare complete main-part fingerprints with static Word oracles. Both sessions save and reopen matching packages.

The affected batch passes 151 cases and 2226 assertions across ten files. Its log is `/tmp/n08-checklist-indent-final.log`. Two Chromium cases pass in `/tmp/n08-checklist-indent-browser.log`. They use real toolbar clicks, verify pointer selection, one-step undo/redo, continued typing, native marker state, and shared saved convergence.

Private types pass 2405 files. The private build verifies 325 modules with core external. Browser types pass 3666 files across two projects with the Vite CSS declarations. Compatibility decision: `.collaboration/changes/sciobot-captured-checklist-indentation.json`. Collaboration transport stays test-only under the EigenPal Pro License. No application, migration, baseline, or reference changes.

Remaining N08 work includes production property ownership, generic property births, portable pending DOCX ownership, source sequence resets, sibling conversion, polite restarts, complete capture parity, and final browser/shared gates. The external-editing metadata limitation remains explicit. Backlog status stays 193 closed and 83 open across 276 leaves.
