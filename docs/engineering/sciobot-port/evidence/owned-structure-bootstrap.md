# Structural ownership bootstrap checkpoint

This checkpoint extends `3abfe153`. D30 remains open.

Complete review history configures native structural keys before bootstrap.
Creators publish the policy. Joiners validate it without writing competing metadata.
Incomplete deliveries wait for their dependencies. Conflicting policies remain refused after reopening.
The evaluated provider exports its keys before a session exists.
Production private modules retain their Apache core dependency and import no Pro runtime.

The planner also restores existing per-native-gap operations before its general mixed-atom fallback.
The original retained-atom expectations remain unchanged.

Verification of the final bootstrap implementation:

- Built private entries, built provider, and built core: 198 passed across seven files, 22,465 assertions.
- Ownership, flags, and history guards: 40 passed across four files, 193 assertions.
- Three source cases that previously exceeded their unchanged five-second limits: three passed, 315 assertions.
- Private type check: 2,220 files passed. Provider type check, builds, and API extraction passed.
- Root formatting and lint passed during this batch. Changed source files received subsequent scoped formatting and lint.

Earlier integration evidence is partial, not a clean full-suite result:

- The 623-case source batch passed 614 cases. Seven cases timed out; two retained-gap expectations failed.
- All seven timed-out cases passed a focused rerun without changed limits.
- After the gap fix, 70 affected cases passed, including the original retained-gap expectations.
- All 333 unchanged C01/composer cases passed before the final bootstrap changes.
- The final 198-case built batch includes both native image histories and replacement-plus-birth coverage.

Local execution logs reside under `/tmp/d30-*.log`; these are ephemeral diagnostic artifacts.
No production audit baseline was recaptured. This checkpoint closes no TODO leaf.
