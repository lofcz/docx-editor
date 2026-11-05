# Sciobot editor UI fidelity

The native DOCX editor must match or improve the production Plate editor's user experience. The earlier migration ledger proves implementation coverage. It does not close this fidelity review.

## Working agreement

- Keep three feature agents active. Use isolated feature worktrees and commit coherent changes.
- Compare the actual Plate controls and native controls before implementation.
- Reuse established application components, including the `@lofcz/mathlive` fork.
- Test pointer and keyboard input, focus, cancellation, validation, history, persistence, and applicable shared editing.
- Inspect desktop and narrow layouts, open controls, zoom, scrolling, and read-only states.
- The coordinator reviews changes and exercises the resulting interface before accepting a feature.
- Record remaining defects explicitly. Partial implementation does not close a feature.
- Desktop Microsoft Word and operating-system input-method checks remain outside the agreed scope.

## Sources

- Production Plate application: `/run/media/lofcz/ssd_external/GitHub/sciobot-next`, reference commit `17e785deb`.
- Plate fork: `/run/media/lofcz/ssd_external/GitHub/plate`.
- MathLive fork: `/run/media/lofcz/ssd_external/GitHub/mathlive`.
- Native integration: `/run/media/lofcz/ssd_external/GitHub/sciobot-next-docx-work`, starting commit `3326720a2`.
- Native engine: this checkout, starting commit `47b456e0d`.

Other sessions use the production application checkout. Feature work stays in isolated worktrees.

## Feature review

| ID | Feature | Required comparison | Owner and state |
| --- | --- | --- | --- |
| UX01 | Tables | Insertion grid, contextual controls, cell and row selection, resizing, structure, appearance, keyboard navigation | Accepted; 13 integrated browser cases pass |
| UX02 | Math | Shared MathLive dialog, virtual keyboard, selected-text insertion, inline/display editing, draft ownership | `ux_math_review`: ownership integrated; concurrent replacement undo and shared preparation guards remain open |
| UX03 | SMILES | Shared chemistry editor, AI and manual entry, preview, validation, activation, block actions, failed-save recovery | Accepted; eight integrated browser cases pass |
| UX04 | Links | Insert/edit/remove dialogs, selected text, URL validation, opening, keyboard shortcuts | Accepted; ten integrated browser cases and direct keyboard review pass |
| UX05 | Images and galleries | Upload, resize, captions, alignment, replacement, gallery actions, loading and errors | `ux_gallery`: image ownership and complete controls in progress |
| UX06 | Files, audio, video, and embeds | Insert/edit controls, previews, loading, playback, removal, accessible fallback | Pending |
| UX07 | Lists | Bullets, numbering, checklists, nesting, indentation, continuation and exit shortcuts | `ux_lists`: conversion and indentation integrated; completion styling and document reload remain open |
| UX08 | Text and paragraph formatting | Marks, font size, colors, headings, alignment, spacing, mixed selection state | Pending |
| UX09 | Fixed and floating toolbars | Control grouping, overflow, responsive layout, tooltips, selection retention, view modes | Pending |
| UX10 | Slash and insertion menus | Complete catalog, feature gates, filtering, keyboard navigation, insertion placement | Pending |
| UX11 | Block selection and movement | Handles, menus, keyboard selection, drag placement, duplication, deletion, nested blocks | Pending |
| UX12 | Columns | Count changes, resizing, movement, empty entry, discoverability and nested content | Pending; migration checks alone do not establish fidelity |
| UX13 | Code blocks | Language picker, editing, copy, syntax display, sizing, shortcuts and exit behavior | Pending |
| UX14 | Callouts and toggles | Variants, icons, collapse controls, body editing, focus and layout stability | Pending |
| UX15 | Stories and solutions | Editing controls, reveal states, content layout, nested editing and persistence | Pending |
| UX16 | Mentions, dates, and emoji | Pickers, search, keyboard navigation, insertion, editing and inline selection | Pending |
| UX17 | Contents and navigation | Heading controls, table of contents, document navigation, updates and focus | Pending |
| UX18 | Lesson titles and phases | Title fields, phase controls, placement, compact layout and editing state | Pending |
| UX19 | Activities and recap | Metadata, duration, goals, materials, adaptations, nested content and field controls | Pending |
| UX20 | Curriculum and sources | Search and selection, source chips, excerpts, actions and contextual detail | Pending |
| UX21 | Listening and lesson visuals | Established media controls, transcript fields, generation states, editing and sizing | Pending |
| UX22 | Shared MDX and custom blocks | Actual registered renderers, inline/block geometry, controls, opaque content and fallbacks | Widget reload fix integrated; browser review and remaining comparison open |
| UX23 | AI and material tools | Full material tools menu, selection quotes, prompts, progress, cancellation and result review | Pending; native toolbar opens a reduced menu |
| UX24 | Suggestions and review | Discovery, navigation, attribution, accept/reject controls, grouping and conflict feedback | Pending |
| UX25 | Core editing interactions | Caret, range selection, clipboard, shortcuts, undo/redo, text entry and focus boundaries | Pending |
| UX26 | Page view and search | Zoom, scrolling, find/replace, page navigation, resizing and selection visibility | Pending |
| UX27 | Document lifecycle | Loading, saving, recovery, import/export, document switches and error feedback | Widget reload fix integrated; browser review and remaining comparison open |
| UX28 | Final comparison | Every production kit mapped; cross-feature flows, accessibility, responsive and visual review | Coordinator: pending |

## Acceptance records

UX01, UX03, and UX04 passed coordinator review. Other features remain open.

- Tables: engine through `190a5cd2f`; app `9e05adfe5` and runtime `e7a050d91`. All 13 integrated browser cases pass with 111 assertions. Coordinator code and visual review pass. Coverage includes insertion, selection, formatting, resize, nested tables, repeated headers, tracked edits, concurrency, history, and reopen. The engine boundary run passes 62 tests; one resource-limit case times out during a parallel build and passes unchanged in isolation within its original five-second limit. Imported legacy horizontal merges and merged RTL resize retain documented limits. The mounted Plate workspace exposes neither RTL table controls nor header-cell formatting controls.
- Links: app `cbe3e8884` and `61068174a` provide the shared app popover over canonical hyperlink commands. Ten integrated browser cases pass with 60 assertions. Coordinator keyboard and visual review also pass. Coverage includes stale drafts, document replacement, concurrent edits, history, save/reopen, view/suggest modes, narrow layout, and hover dismissal. Native links accept real pasted app URLs and bookmarks. The Plate reference contains a mock page catalog; this integration does not copy that mock. Field hyperlinks retain the engine's preview-only support.
- Math: app `885d05104` reuses the shared MathLive surface. Engine `078ed3e31` retains atomic equation ownership across display conversion. Further review finds a concurrent replacement undo defect: one client order restores peer TeX but empties its Word fallback. Shared preparation guards also need correction. These defects block acceptance. Schema 10 requires export and reseeding into fresh rooms; existing room histories cannot migrate in place.
- SMILES: app `5daf553f7`, `bb5bdf946`, and `60e0c27e3` share the Plate chemistry editor and protect draft ownership. Eight integrated browser cases pass. Coordinator review covers desktop and narrow layouts, keyboard input, cancellation, validation, replacement, history, shared editing, duplication, deletion, and save/reopen. AI request tests use controlled responses; they verify interaction and cancellation, not model output quality. Plate provides AI and manual entry with a preview, without a freehand drawing canvas.
- Shared block ownership: engine `eac568851` fixes deletion of the wrong custom XML payload after duplication. The fix passes 59 focused tests. The integrated browser duplication and deletion check also passes.
- Selected-text insertion: engine `b378d94d1` replaces the selected range in one transaction. Six tests cover concurrent edits, history, and save/reopen. The app integration passes selected-text replacement, one-step history, shared editing, and reopen.
- Widget reload: engine `25c609fcd` retains block and inline registrations across `editor.load()`. Each document receives fresh widget state and measurements. Repeated document switches release old roots. Focused shared editing, history, deletion, and reopen checks pass. Integrated host and list checks pass 23 tests with 440 assertions. Browser review remains open.
- Lists: app `df1777fd9` and `f27aebba3` restore list conversion and list-level toolbar indentation. Engine `d1a4e7779` preserves nesting when task prefixes replace bullets or numbering. Completed-task styling, shared preparation guards, and listener reload remain open.
- Media: app `9ee3b3002` and `ec250fa4f` require explicit document-owned loading for gallery images, audio, video, and embeds. Imported sources remain inert until approval. Full media and gallery control comparison remains open.

The experimental app runtime uses engine `190a5cd2f`. Archive hashes match, and the app resolves one physical core package.
