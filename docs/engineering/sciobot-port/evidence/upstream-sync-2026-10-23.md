# Upstream reconciliation, 2026-10-23

Upstream `main` is `b9cbb36f`. The detour imports 11 commits after `241085fd`, across 103 files. Local `main` fast-forwards to upstream. Primary merge `1672c213` retains the fork changes. Active D30 merge `06d2f43d` also retains its unfinished work. These are local updates; this detour does not push branches.

The upstream changes cover CJK automatic spacing and punctuation fitting, picture clipping, column balance, canvas caching, and evaluation tooling. The merge preserves the fork's inline widgets, equations, native adapters, and collaboration work. No merge conflicts require manual resolution.

The merged table layout exceeds its existing line cap by five lines. A dedicated pagination-error module retains the same constructor and codes. The original module re-exports them. Direct imports remove two unnecessary dependencies on the large layout module. Core API extraction checks all 15 entries without snapshot changes.

The frozen layout run reports 7,315 passes and two timeouts across 581 files. The list resolver retry passes with its original 5-second limit. The table-width side-rule case still times out alone: 5.30 seconds after merging, and 5.58 seconds on unchanged `75930cb4`. Both reach 81 assertions before the timeout. This verifies a pre-existing gate failure, not completion of that test. No timeout or acceptance assertion changes.

Selected adapter and font checks pass 132 cases with 2,276 assertions. Upstream two-editor checks pass 13 cases with 477 assertions. The PDF evaluation loader passes 11 cases with 98 assertions. Error extraction checks pass 28 cases with 366 assertions, plus five exact-row cases with 92 assertions. Active D30 checks pass 39 cases with 707 assertions. They cover code controls, linked controls, linked shared histories, inline code, insertion marks, and C01.

Core types check 1,595 files after extraction. Active private adapter types check 2,178 files. Core and private package builds pass; the private build keeps core external across 256 private modules. Root formatting and lint pass. Compatibility policy passes with 176 primary-branch decisions and no changed format fields.

This detour closes no port leaf. The ledger remains at 179 closed leaves and 97 open leaves. D30 still has the recorded prefix-plus-code peer-placement failure. Its exact native link oracle remains unchanged. No application source or production artifacts change.
