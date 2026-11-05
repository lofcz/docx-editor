# Native column controls

Application commit `3326720a299b4e110bfe200a9763e943a55005cd` installs runtime `7d5131205ac226a7e8db1f369615864e077d0c2b`.

Column groups use native borderless tables. Native paragraphs retain their identities, formatting, selection, and editing history.

- Width handles support pointer and keyboard input, cancellation, and zoom.
- Content movement retains native blocks and creates an editable paragraph when the source becomes empty.
- **Turn into → 3 columns** wraps the selected paragraph or changes its enclosing group's count.
- Count reduction gathers removed content into the final surviving column.
- Explicit entry creates a native paragraph in an empty column restored by undo. Peer deletions remain deleted.
- Legacy conversion requires a local document without an attached collaboration session.
- Structural column actions refuse suggestion mode. Native tracked text supports acceptance, rejection, and undo.

Shared wrapping requires `prepareNativeDocument` before room bootstrap. This keeps shared payload storage outside a participant's undo history. Concurrent wrapping preserves both groups and one original paragraph. The command refuses unprepared rooms without changing their contents.

## Verification

The final installed Chromium run passes 12 cases and 140 assertions across `column-counts.test.ts`, `columns-controls.test.ts`, `columns-insert.test.ts`, and `legacy-columns.test.ts`. It uses the actual application controls and two real collaboration sessions. Both receiver orders pass.

The cases cover native identities, pointer and keyboard controls, cancellation, zoom, count changes, tracked text, peer deletion, history, offline conversion, shared conversion refusal, and save/reopen. Local conversion also survives reload. Narrow viewport screenshots show restored text in both editors.

The count, geometry, and conversion checkpoint passes 49 engine cases with 670 assertions. The final wrapping guard passes nine cases with 240 assertions. These checks preserve exact native structure and saved semantic digests.

The application build, frontend compiler, harness compiler, scoped formatting, and lint pass. All five archive hashes match `vendor/docx-native/SOURCE.json`. The application and private adapter resolve one core module.

Logs:

- `/tmp/sciobot-column-final-browser.log`
- `/tmp/sciobot-column-final-build.log`
- `/tmp/sciobot-column-final-types.log`
- `/tmp/sciobot-column-final-harness-types.log`
- `/tmp/sciobot-column-final-format-check.log`
- `/tmp/sciobot-column-final-lint.log`
- `/tmp/sciobot-column-final-archives.log`
- `/tmp/docx-column-count-final.log`
- `/tmp/docx-column-bootstrap-tests.log`

Desktop Microsoft Word and operating-system input-method verification remain unverified. The user removed those checks from the required scope on November 5, 2026. Existing repository baseline failures remain recorded in [upstream verification](upstream-234-verification.json).
