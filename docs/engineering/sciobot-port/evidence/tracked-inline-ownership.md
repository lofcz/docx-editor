# Tracked inline payload ownership

Runtime checkpoint: `44f01237`, branch `codex/d30-tracked-inline-ownership`.
D30 and R19 remain open.

Application presence proposal schema 3 couples an inline control's complete payload and normalized native fingerprint to one exact native insertion or deletion wrapper.
The group decides native content and payload presence together.
Owned inverses preserve native control identities, source run boundaries, and source hyperlink parents.
Block-only presence retains schema 2.

Schema 3 is required for safety across application upgrades.
Older application readers must refuse inline ownership even after its native wrapper disappears through resolution or history.
The current reader reserves retained native addresses in pending, resolved, and tombstoned states.
The generic property store deliberately retains unsupported records as opaque data, including tombstoned records.
It must not hide unsupported ownership by interpreting newer application proposal schemas.
No native transport field or version changes.
Upgrade private application readers before using the new profile.

Evidence at this checkpoint:

- `/tmp/d30-inline-ownership-schema3-integration.log`: 52 passed and one new tombstone expectation failed across eight files.
- The failed expectation incorrectly assumed the generic store hides unsupported tombstones. The corrected expectation preserves opaque records.
- `/tmp/d30-inline-ownership-reservation-final.log`: the corrected reservation case passed, including actual native insertion without address reuse.
- The integration batch includes all eight inline shared cases, existing block presence, lifecycle, owned inverses, allocation, and C01 cases.
- `/tmp/d30-inline-schema3-compat.log`: a separate cross-artifact probe passed pending, resolved, and tombstoned states against built old and current private readers.
- Old artifact SHA256: `ecd9bbb3955546bbdbd5805278589e187e385fbead73b66444982064da82484d`.
- Current artifact SHA256: `fc43693a84718ca8444529721842945fb29fde24c7e346e0b30e2d5072c71f44`.
- Private types passed 2,234 files; the private build verified 274 production modules with core external.
- Core build, root formatting, root lint, and line caps passed. Later schema files passed scoped formatting and lint.

The full target coordinator still refuses inline births, image replacement/removal, and combined atom/text reconciliation at this checkpoint.
The cross-artifact probe verifies reader compatibility and retained reservations; it does not establish complete compiled authoring.
The installed application remains read-only and its production revision audit still fails pinned revision drift.
No baseline was recaptured, test timeout increased, or production integration claimed.

## AI coordinator removal checkpoint

Runtime: `3b388abc`.
Removal-only AI targets now use the real ordered transaction writer.
The source validator is shared with direct removal and performs no candidate operations.
The complete native control and payload remain pending until their exact group is decided.
Host content locks return to their original native properties subtree before ownership proof capture.
An absent-side decision releases only the exact owned host content lock before removing that control.
Ordinary native protection rules remain active.

This checkpoint admits removal-only plans; it does not admit dependent mixed coordinate plans.
Inline births, image resource ownership, and OMML still refuse before live publication.
The acceptance matrix currently proves mentions, including plain and linked controls with host content locks.
Other supported payload categories require their coordinator acceptance coverage before feature closure.

Verification used one affected source batch, focused corrections, and one built-SDK boundary check:

- `/tmp/d30-coordinator-removal-batch.log`: 53 passed and six new coordinator fixtures failed across five files.
- All existing 49 cases passed, including direct removals, local/shared inline ownership, and application fragment recomposition.
- Four failures came from a new fixture assuming all mentions project as MDX. Plain mentions project as mention links.
- Another fixture failed to alter its escaped source text. Both fixture assumptions were corrected without changing product oracles.
- The remaining failure exposed a real host-lock gap during acceptance; the exact owned decision now releases that lock.
- `/tmp/d30-coordinator-removal-fixed.log`: all 22 local/shared cases passed after those corrections.
- `/tmp/d30-coordinator-removal-waves-locks.log`: five focused cases passed, including three AI waves and locked shared removal.
- Shared coordinator cases cover both client orders, peer typing, both actor histories, decisions, geometry, and DOCX/room reopening.
- `/tmp/d30-coordinator-removal-compiled-final.log`: all 11 coordinator cases passed against the built SDK and built core, 194 assertions.
- The compiled fixtures now construct their editor through the product module directly, avoiding source-backed product helpers.
- An initial compiled command refused an unmapped public planner import before running tests. The harness now maps that established export.
- No internal staging helpers were exported for compiled checks.
- `/tmp/d30-coordinator-removal-final-types.log`: 2,236 files passed.
- `/tmp/d30-coordinator-removal-build.log`: 275 production modules passed the package boundary check with native core external.
- Changed-file formatting, lint, and line caps passed.
- `/tmp/d30-coordinator-removal-compatibility.log`: policy passed with one decision and two inherited format fields.

D30 and R19 remain open. No installed-application audit or complete final composer gate is claimed for this checkpoint.
