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
| UX01 | Tables | Insertion grid, contextual controls, cell and row selection, resizing, structure, appearance, keyboard navigation | `ui_tables`: implementing |
| UX02 | Math | Shared MathLive dialog, virtual keyboard, selected-text insertion, inline/display editing, draft ownership | `ui_math`: implementing |
| UX03 | SMILES | Shared chemistry editor, drawing and preview, validation, activation, block actions, failed-save recovery | `ui_smiles`: implementing |
| UX04 | Links | Insert/edit/remove dialogs, selected text, URL validation, opening, keyboard shortcuts | Pending; native toolbar lacks a usable editor |
| UX05 | Images and galleries | Upload, resize, captions, alignment, replacement, gallery actions, loading and errors | Pending |
| UX06 | Files, audio, video, and embeds | Insert/edit controls, previews, loading, playback, removal, accessible fallback | Pending |
| UX07 | Lists | Bullets, numbering, checklists, nesting, indentation, continuation and exit shortcuts | Pending |
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
| UX22 | Shared MDX and custom blocks | Actual registered renderers, inline/block geometry, controls, opaque content and fallbacks | Pending |
| UX23 | AI and material tools | Full material tools menu, selection quotes, prompts, progress, cancellation and result review | Pending; native toolbar opens a reduced menu |
| UX24 | Suggestions and review | Discovery, navigation, attribution, accept/reject controls, grouping and conflict feedback | Pending |
| UX25 | Core editing interactions | Caret, range selection, clipboard, shortcuts, undo/redo, text entry and focus boundaries | Pending |
| UX26 | Page view and search | Zoom, scrolling, find/replace, page navigation, resizing and selection visibility | Pending |
| UX27 | Document lifecycle | Loading, saving, recovery, import/export, document switches and error feedback | Pending |
| UX28 | Final comparison | Every production kit mapped; cross-feature flows, accessibility, responsive and visual review | Coordinator: pending |

## Acceptance records

No feature has passed coordinator review yet. Each record must identify its commits, actual Plate comparison, browser checks, and unresolved limitations.
