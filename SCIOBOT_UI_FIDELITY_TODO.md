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
| UX02 | Math | Shared MathLive dialog, virtual keyboard, selected-text insertion, inline/display editing, draft ownership | Accepted; 19 integrated browser cases pass |
| UX03 | SMILES | Shared chemistry editor, AI and manual entry, preview, validation, activation, block actions, failed-save recovery | Accepted; eight integrated browser cases pass |
| UX04 | Links | Insert/edit/remove dialogs, selected text, URL validation, opening, keyboard shortcuts | Accepted; ten integrated browser cases and direct keyboard review pass |
| UX05 | Images and galleries | Upload, resize, captions, alignment, replacement, gallery actions, loading and errors | `ux_gallery`: upload and SVG conversion integrated; replacement and single-image controls remain open |
| UX06 | Files, audio, video, and embeds | Insert/edit controls, previews, loading, playback, removal, accessible fallback | Pending |
| UX07 | Lists | Bullets, numbering, checklists, nesting, indentation, continuation and exit shortcuts | Accepted; seven integrated list cases and two export cases pass |
| UX08 | Text and paragraph formatting | Marks, font size, colors, headings, alignment, spacing, mixed selection state | Accepted; nine integrated browser cases pass |
| UX09 | Fixed and floating toolbars | Control grouping, overflow, responsive layout, tooltips, selection retention, view modes | Accepted; 27 integrated browser cases pass |
| UX10 | Slash and insertion menus | Complete catalog, feature gates, filtering, keyboard navigation, insertion placement | Accepted; 42 integrated browser cases and coordinator query review pass |
| UX11 | Block selection and movement | Handles, menus, keyboard selection, drag placement, duplication, deletion, nested blocks | `ux_lists`: keyboard ownership integrated; menus, movement, and ordinary paragraphs remain open |
| UX12 | Columns | Count changes, resizing, movement, empty entry, discoverability and nested content | Pending; migration checks alone do not establish fidelity |
| UX13 | Code blocks | Language picker, editing, copy, syntax display, sizing, shortcuts and exit behavior | `ux_math_review`: source ownership and shared Plate authoring controls |
| UX14 | Callouts and toggles | Variants, icons, collapse controls, body editing, focus and layout stability | Pending |
| UX15 | Stories and solutions | Editing controls, reveal states, content layout, nested editing and persistence | Pending |
| UX16 | Mentions, dates, and emoji | Pickers, search, keyboard navigation, insertion, editing and inline selection | Pending |
| UX17 | Contents and navigation | Heading controls, table of contents, document navigation, updates and focus | Accepted; seven integrated browser cases and coordinator keyboard review pass |
| UX18 | Lesson titles and phases | Title fields, phase controls, placement, compact layout and editing state | Pending |
| UX19 | Activities and recap | Metadata, duration, goals, materials, adaptations, nested content and field controls | Pending |
| UX20 | Curriculum and sources | Search and selection, source chips, excerpts, actions and contextual detail | Pending |
| UX21 | Listening and lesson visuals | Established media controls, transcript fields, generation states, editing and sizing | Pending |
| UX22 | Shared MDX and custom blocks | Actual registered renderers, inline/block geometry, controls, opaque content and fallbacks | Widget reload fix verified; remaining comparison open |
| UX23 | AI and material tools | Full material tools menu, selection quotes, prompts, progress, cancellation and result review | Pending; native toolbar opens a reduced menu |
| UX24 | Suggestions and review | Discovery, navigation, attribution, accept/reject controls, grouping and conflict feedback | Pending |
| UX25 | Core editing interactions | Caret, range selection, clipboard, shortcuts, undo/redo, text entry and focus boundaries | Pending |
| UX26 | Page view and search | Zoom, scrolling, find/replace, page navigation, resizing and selection visibility | Pending |
| UX27 | Document lifecycle | Loading, saving, recovery, import/export, document switches and error feedback | Widget reload verified; font embedding, app TOC export, and remaining comparison open |
| UX28 | Final comparison | Every production kit mapped; cross-feature flows, accessibility, responsive and visual review | Coordinator: pending |

## Acceptance records

UX01, UX02, UX03, UX04, UX07, UX08, UX09, UX10, and UX17 passed coordinator review. Other features remain open.

- Tables: engine through `190a5cd2f`; app `9e05adfe5` and runtime `e7a050d91`. All 13 integrated browser cases pass with 111 assertions. Coordinator code and visual review pass. Coverage includes insertion, selection, formatting, resize, nested tables, repeated headers, tracked edits, concurrency, history, and reopen. The engine boundary run passes 62 tests; one resource-limit case times out during a parallel build and passes unchanged in isolation within its original five-second limit. Imported legacy horizontal merges and merged RTL resize retain documented limits. The mounted Plate workspace exposes neither RTL table controls nor header-cell formatting controls.
- Links: app `cbe3e8884` and `61068174a` provide the shared app popover over canonical hyperlink commands. Ten integrated browser cases pass with 60 assertions. Coordinator keyboard and visual review also pass. Coverage includes stale drafts, document replacement, concurrent edits, history, save/reopen, view/suggest modes, narrow layout, and hover dismissal. Native links accept real pasted app URLs and bookmarks. The Plate reference contains a mock page catalog; this integration does not copy that mock. Field hyperlinks retain the engine's preview-only support.
- Math: app through `46e949bb4` uses Plate's shared `@lofcz/mathlive` controls. Nineteen integrated browser cases pass on runtime F. Coverage includes keyboard input, narrow layouts, stale drafts, reload, both client orders, conversion, exact multiline TeX, history, cursors, and reopen. Coordinator review finds and fixes tall-formula clipping, short-formula bounds, and early keyboard focus during MathLive startup. Failed downloads retain modal focus and preserve the document. The ownership regression checks retain their original assertions. Schema 10 requires export and reseeding into fresh rooms; existing room histories cannot migrate in place.
- SMILES: app `5daf553f7`, `bb5bdf946`, and `60e0c27e3` share the Plate chemistry editor and protect draft ownership. Eight integrated browser cases pass. Coordinator review covers desktop and narrow layouts, keyboard input, cancellation, validation, replacement, history, shared editing, duplication, deletion, and save/reopen. AI request tests use controlled responses; they verify interaction and cancellation, not model output quality. Plate provides AI and manual entry with a preview, without a freehand drawing canvas.
- Shared block ownership: engine `eac568851` fixes deletion of the wrong custom XML payload after duplication. The fix passes 59 focused tests. The integrated browser duplication and deletion check also passes.
- Selected-text insertion: engine `b378d94d1` replaces the selected range in one transaction. Six tests cover concurrent edits, history, and save/reopen. The app integration passes selected-text replacement, one-step history, shared editing, and reopen.
- Widget reload: engine `25c609fcd` retains block and inline registrations across `editor.load()`. Each document receives fresh widget state and measurements. Repeated document switches release old roots. Focused shared editing, history, deletion, and reopen checks pass. Integrated host and list checks pass 23 tests with 440 assertions. Two browser cases pass with 20 assertions against isolated host build `7b9cce3db` and runtime B. App `23a78c80b` requires visible controls after reload and fresh explicit media approval. Runtime C includes this fix. The integrated math checks also verify document-owner reload.
- Lists: engine `998948988` derives completion appearance from checkbox state without changing paragraph style ownership. Original review assertions remain unchanged. Detached Word downloads materialize settled appearance through reversible character styles. Pending revisions remain exact; their appearance is not materialized. App `29a6471d3` adds accessible controls, and `1df3651c6` wires native export. Seven integrated list cases and two export cases pass. Coordinator visual review passes. Combined engine checks pass 107 tests across list appearance, export, and widget hosts.
- Formatting: app `6554c7abd` matches Plate's Inter body and heading sizes with six packaged static faces. Nine integrated browser cases pass with 114 assertions. Checks cover font drafts, validation, colors, mixed selections, imported styles, shared history, and reopen. Coordinator review verifies 30-pixel headings and the complete narrow-screen palette. The assets match Plate's actual Latin and Latin Extended sources. Native save/reopen retains font names; exported DOCX files do not embed host fonts. Font embedding remains open under UX27.
- Toolbars: app `d77812126` anchors floating actions at the selection focus and keeps them inside the visible viewport. The final integrated run passes 27 cases with 144 assertions and exits successfully. Coordinator review covers backward selections, narrow layouts, zoom, tooltips, keyboard focus, owner changes, and pending quote cancellation. App `72c3e4cd5` and `e12fe7c32` wait for actual tooltip and keyboard-focus transitions.
- Managed images: engine `ced3b522c` gives native drawing properties ownership of descriptions and dimensions. Private records retain resource provenance. Ninety integrated tests pass with 12,090 assertions. Coverage includes concurrent edits, reconnect, history, copying, deletion, protection, and save/reopen. App image controls remain under review.
- SVG conversion: app `168ffbae8` converts explicitly selected SVG files into bounded native PNG images. Three integrated browser cases pass. They cover both client orders, history, reopen, hostile content, and refusal before oversized allocations. Single-image controls remain open.
- Block selection: app `98301386d` keeps keyboard shortcuts with their editor and clears replaced or deleted owners. Four integrated browser cases pass. Ordinary paragraph selection, menus, and movement remain open.
- Code language: engine `30c57b692` preserves source paragraph identities when the language changes. Six integrated tests pass with 105 assertions. Both client orders retain independent peer history. Full formatting, lint, and private adapter types pass. Runtime F does not yet include this change.
- Contents: app `392a8b6d7` and `bbfb4bc20` use canonical outline styles and shared Plate row controls. Seven integrated browser cases pass. Coverage includes inherited styles, table headings, long titles, peer deletion, history, reopen, and remote cursor geometry. Coordinator keyboard and desktop/narrow visual review pass. Header content stays outside the contents list. The application TOC's Word fallback remains open under UX27.
- Insertion menus: app through `43917fed3` uses Plate's catalog and shared query controls. Forty-two integrated browser cases pass at `a8012cca9`. Coverage includes repeated document replacement, first-trigger behavior, keyboard navigation, filtering, composition, feature gates, narrow layouts, peer relayout, history, and reopen. Coordinator query editing and cancellation preserve the document and history. The mention history check first settles preceding native text input; its cancellation assertion remains exact.
- Drawing selection: engine `643caaeec` follows exact replica identifiers during publication. Thirty-eight focused tests pass with 292 assertions. Selection survives alignment and reconnect. Peer deletion never transfers selection to another image at the same position. Core build and all 15 API reports pass. Full format and lint pass. A new test used an incorrect command property; its correction passes private types, while all other workspace type checks already pass. Runtime F does not include this repair.
- Gallery controls: engine `98587b001` keeps opaque widgets connected during painting and scales controls with document zoom. App through `421dfd1ca` shares Plate gallery controls and preserves multiline caption focus. Four isolated browser cases pass with 61 assertions. App `cf61e4b96` keeps all five controls reachable on phones. Two integrated keyboard and touch cases pass with 19 assertions. Coordinator visual review passes. Single-image controls and SVG conversion remain open.
- Media: app `9ee3b3002` and `ec250fa4f` require explicit document-owned loading for gallery images, audio, video, and embeds. Imported sources remain inert until approval. Full media and gallery control comparison remains open.

The installed runtime F uses engine `fb6f6405b`, pinned by app `000d05a1d`. The private adapter and application resolve one physical core copy. The first combined browser run passes 44 cases with 391 assertions. The separate toolbar run passes 27 cases. Direct app and harness type checks pass, covering 8,484 and 3,226 files. The production build at `000d05a1d` exits successfully and includes all six pinned font files and their license notice. Full engine formatting and lint also pass.

The later menu, upload, and contents boundary at app `8f2fc8515` passes 56 cases and fails two. App `43917fed3` resolves both failures; its focused boundary passes 42 cases. App types pass 8,488 files before the shared image toolbar batch. App `cecdd8009` repairs a test-helper import, and harness types then pass 3,233 files. The SVG and block ownership boundary at `98301386d` passes seven cases with 68 assertions. Block drag changes at `c51d206a7` pass 16 cases with 75 assertions across new and original interaction tests. Activity card border overlap remains open under UX19.

The completed 47-file regression run at `a7cb34701` records 840 passes and 77 failures. Sixty-nine failures reproduce at the exact starting revision, `47b456e0d`. Four earlier baseline failures now pass. The eight remaining failures are timeouts; all pass unchanged within their original limits when rerun. This result does not establish a passing full suite. The original image creation-result test passes after the identity repair.

Workspace type checks, adapter parity, localization, lane checks, and the Office.js signature report pass at the integrated boundary. Fifteen fresh core API reports match after `104843ff1`. The collaboration metadata check still references a missing pre-existing paragraph-alignment test. The OpenSpec command is unavailable in this checkout.
