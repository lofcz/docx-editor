# Cross-feature evidence and cutover gates

These leaves collect required evidence. Passing one gate does not replace focused checks on each feature.

Read the [dispatch and completion rules](../../../SCIOBOT_PORT_TODO.md) before claiming a leaf.

Each named test or artifact is a deliverable. These specifications do not claim those checks already exist.

## V01

- [x] Generalize the two-editor collaboration fixture

**Depends:** None.

**Primary files:** `fork:e2e/support/sciobot-collaboration.ts`; `smoke fixture using examples/react-blocks/test/collaboration.test.ts`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Provide two real editor sessions, pause/reconnect, semantic digests, and presence assertions. Keep the evaluated transport in test-only dependencies.

**Close when:** The smoke case covers concurrent edits, undo/redo, deletion during peer editing, remote selections, reconnect, and save/reopen convergence.

**Check:** Run the two-real-session smoke scenario and save its semantic convergence assertions.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `collaboration-harness.test.ts` passes one real two-session scenario and 18 assertions: live synchronization, offline concurrent typing, peer-preserving undo/redo, remote selections, offline deletion during peer editing, reconnect, DOM repaint, and equal save/reopen digests. The reusable fixture owns test-only transport dependencies in `packages/sciobot/test/support`; e2e re-exports it. Package typecheck, scoped Prettier, and oxlint pass. Compatibility decision: tests only; no shared format or production runtime change.

## V02

- [ ] Compare both AI modes against the seven Plate fixtures

**Depends:** [C01](01-contracts.md#c01), [D15](04-edit-planner.md#d15), [A01](06-direct-ai.md#a01), [R17](05-suggestions.md#r17), [R18](05-suggestions.md#r18), [V01](11-verification.md#v01)

**Primary files:** `fork:packages/sciobot/src/__tests__/ai-parity.test.ts`

**Read:** [S03](sources.md#s03), [S04](sources.md#s04)

**Change:** Run every C01 pair through direct edits and accepted suggestions. Compare independently captured normalized semantics, not only the two implementations.

**Close when:** All seven pairs match the Plate oracle. Reject restores the baseline. Direct undo/redo and review decision undo/redo restore exact states.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## V03

- [ ] Exercise mixed review waves through the application

**Depends:** [I06](10-app-integration.md#i06), [U09](09-editor-ui.md#u09), [U10](09-editor-ui.md#u10), [U11](09-editor-ui.md#u11), [U12](09-editor-ui.md#u12), [R20](05-suggestions.md#r20), [V01](11-verification.md#v01)

**Primary files:** `app:workspace native editor end-to-end tests and fixtures`

**Read:** [S02](sources.md#s02), [S07](sources.md#s07), [S08](sources.md#s08)

**Change:** Run two suggestion waves with a manual edit, accepted and rejected groups, and a property change. Include a direct edit between waves.

**Close when:** Agent reads show the working view. Decisions preserve user text. Reload restores pending groups with stable identities.

**Check:** Add a focused application integration case. Use browser input for user-facing actions. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## V04

- [ ] Verify cross-feature collaboration convergence

**Depends:** [V02](11-verification.md#v02), [V03](11-verification.md#v03), [B10](02-react-blocks.md#b10), [B13](02-react-blocks.md#b13), [X05](08-domain-content.md#x05), [X10](08-domain-content.md#x10), [I08](10-app-integration.md#i08), [V01](11-verification.md#v01)

**Primary files:** `fork:e2e/sciobot-parity-collaboration.spec.ts`; `application integration fixture`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Exercise nested block moves, property proposals, text suggestions, and listening metadata through two real sessions.

**Close when:** Concurrent edits, peer deletion, undo/redo, presence, and reconnect converge. Both saved documents reopen with equal semantic digests.

**Check:** Add a focused native document regression. Include save/reopen and undo/redo. Add the focused two-session scenario described in the shared completion rules.

**Claim:** unclaimed. **Evidence:** pending.

## V05

- [ ] Measure large-document interaction and retention

**Depends:** [B16](02-react-blocks.md#b16), [I01](10-app-integration.md#i01), [U10](09-editor-ui.md#u10)

**Primary files:** `fork:e2e/sciobot-performance.spec.ts`; `docs/engineering/sciobot-port/evidence/performance.json`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Measure open, typing, review application, page movement, and retained roots on a fixed 100-page fixture. Record environment and raw samples.

**Close when:** Publish p50/p95 timings and memory observations against the same native baseline. File explicit leaves for regressions and keep readiness red until resolved.

**Check:** Save the repeatable command, environment, fixture hash, raw samples, and regression decision.

**Claim:** unclaimed. **Evidence:** pending.

## V06

- [ ] Verify payload and fallback behavior in desktop Word

**Depends:** [R22](05-suggestions.md#r22), [X09](08-domain-content.md#x09), [X19](08-domain-content.md#x19), [X23](08-domain-content.md#x23), [I14](10-app-integration.md#i14)

**Primary files:** `fork:docs/engineering/sciobot-port/evidence/word-roundtrip.md`; `sanitized resulting DOCX fixtures`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Open and save exported files in desktop Word. Exercise native review acceptance/rejection and ordinary text edits around custom blocks.

**Close when:** Reopen in the fork and verify payloads, review decisions, tables, columns, equations, and complete fallbacks. Record Word version and visual evidence.

**Check:** Use desktop Microsoft Word. ZIP inspection or a browser screenshot cannot substitute for this check.

**Claim:** unclaimed. **Evidence:** pending.

## V07

- [ ] Check native runtime package boundaries

**Depends:** [C03](01-contracts.md#c03), [I21](10-app-integration.md#i21)

**Primary files:** `fork:private adapter dependency check`; `app:production bundle report`

**Read:** [S24](sources.md#s24), [S25](sources.md#s25)

**Change:** Inspect source imports and built runtime graphs. Exclude Pro, editor-api, and Plate editor runtime from the replacement path.

**Close when:** The graph contains one core copy and application-owned adapters. Test-only compatibility dependencies do not appear in the shipped route.

**Check:** Run scoped build and typecheck checks. Save the inspected dependency or bundle evidence.

**Claim:** unclaimed. **Evidence:** pending.

## V08

- [ ] Close the production coverage ledger

**Depends:** Every other leaf in this backlog.

**Primary files:** `fork:docs/engineering/sciobot-port/coverage.md`; `evidence index`

**Read:** [S01](sources.md#s01), [S21](sources.md#s21), [S22](sources.md#s22), [S23](sources.md#s23)

**Change:** Match every production capability to completed leaves and reproducible evidence. Include migration and the reversible editor selection path.

**Close when:** No required coverage row remains pending or refused. Every failed or unverified requirement has an open leaf and prevents cutover.

**Check:** Validate every referenced file, task ID, and required coverage row.

**Claim:** unclaimed. **Evidence:** pending.
