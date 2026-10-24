# Native tracked control removal

Runtime: `5f5ea5fd`. D30 and R19 remain open.

`removeContentControl` now accepts native revision attribution with `keepContent: false` for inline controls.
The complete control stays inside one native `w:del` until a decision.
Control and descendant identities remain stable.
Text and field instructions use their native deleted forms.
Rejection restores the complete original Word fingerprint and identities.
Acceptance removes the control.

Tracked unwrapping, block controls, wrapper locks, pending native ownership, and OMML subtrees refuse.
Ordinary untracked removal retains its existing behavior.
The API snapshot and minor changeset describe this additive infrastructure option.
The application coordinator still needs tracked atom payload and resource ownership before D30 closure.

Evidence:

- `/tmp/d30-tracked-removal-batch.log`: 21 passed across four files, 438 assertions.
- `/tmp/d30-tracked-field.log`: the additional complex-field and unknown-metadata case passed, six assertions.
- `/tmp/d30-tracked-removal-compiled-final.log`: eight unchanged shared cases passed against built core and both built private entries, 336 assertions.
- Shared cases cover both client orders, plain and linked controls, concurrent typing, both actor histories, both decisions, decision history, caret geometry, and DOCX/room reopening.
- `/tmp/d30-tracked-removal-core-types-closed.log`: 1,606 files passed.
- Core build and API extraction passed. API extraction checked 15 entries with zero errors.
- Root formatting, root lint, and compatibility policy passed.

The first compiled-run command selected the wrong fixture directory and did not execute tests.
The runner now accepts explicit `review/name.test.ts` selections while retaining all fixture assertions.
Only the corrected compiled command supplies execution evidence.
No full suite was repeated after the focused field correction.
