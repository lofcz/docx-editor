# WinPC editor performance

Use `codex/react-block-proof` for progressive integration. Sciobot uses `claude/docx-native-cutover`.

## Measured results

| Installed native baseline operation         |   Native |    Plate | Sample                     |
| ------------------------------------------- | -------: | -------: | -------------------------- |
| Paste 431 characters                        | 35.27 ms | 24.44 ms | Four paired trials, median |
| Undo that paste                             | 18.20 ms |  8.02 ms | Four paired trials, median |
| Redo that paste                             | 17.01 ms |  6.32 ms | Four paired trials, median |
| Paste 158,633 characters through two frames | 223.4 ms | 121.5 ms | One trial per editor       |

The HarfBuzz change saved 29.0 ms of layout and 28.2 ms of renderer CPU on the largest paste. Both values are median paired differences across four WinPC trials against the native baseline. A direct Plate comparison for the consolidated commit is pending.

[Detailed Sciobot ledger and test conditions](https://github.com/sciocz/sciobot-next/blob/f16827cf8/docs/engineering/sciobot-port/evidence/winpc-editor-performance.md)

## Source history

The active source at consolidation includes upstream, shared history, caret preparation, and checked glyph widths. The archive merge preserves the other experimental commits with their original trees. An archived tip can contain an unsuccessful experiment, a superseded implementation, or an equivalent patch. Use `git show <commit>` to inspect a tip. Use `git diff a13a086e1...<commit>` to review its original divergence.

| Local branch at consolidation | Commit | Included in development history before archival |
| --- | --- | --- |
| `codex/browser-layout-performance` | `0601ad0fd73d` | Already integrated |
| `codex/react-block-proof` | `a13a086e1011` | Already integrated |
| `fix/experimental-peer-insert-suffix` | `0776de2d1208` | Already integrated |
| `fix/peer-insert-suffix` | `e13c3e68e1d5` | Archived |
| `fix/recorded-shared-redo` | `defa518ac3bc` | Already integrated |
| `fix/retained-review-joins` | `58ab9a16b051` | Archived |
| `integration/shared-history-upstream` | `18d4048fdc87` | Already integrated |
| `integration/upstream-hold-repair20` | `3b3361106b3f` | Already integrated |
| `main` | `f1175b8184b7` | Already integrated |
| `perf/bounded-prepared-opening` | `ef45a6578c70` | Archived |
| `perf/browser-caret-collapse` | `463354c869e5` | Archived |
| `perf/caret-after-selection` | `109380ff8128` | Archived |
| `perf/caret-endpoint-lookup` | `b145c4a4b7bd` | Archived |
| `perf/checked-shape-width` | `7768fb96c2d5` | Already integrated |
| `perf/combined-peer-picture-open` | `37bc9f95dafc` | Already integrated |
| `perf/control-boundary-delta` | `3977666583d0` | Archived |
| `perf/control-chrome-geometry` | `0fc52688658d` | Archived |
| `perf/deferred-projection-caret` | `4ed9cafe66f3` | Already integrated |
| `perf/first-disclosure-return` | `415ad276316b` | Archived |
| `perf/initial-caret-owner` | `2cd948ce5ce9` | Already integrated |
| `perf/initial-widget-reservations` | `0ca9035be542` | Already integrated |
| `perf/isolated-content-control-layers` | `80547b8e1e40` | Archived |
| `perf/latin-fitting-words` | `ba4635278309` | Archived |
| `perf/lazy-review-payloads` | `55725ba2d31b` | Archived |
| `perf/lesson-layout-resolution` | `180665d24cb2` | Archived |
| `perf/lesson-live-caret` | `8ca103c6489b` | Archived |
| `perf/lesson-native-word-ranges` | `5b3c58a91c36` | Archived |
| `perf/lesson-paced-publication` | `09acec39130e` | Archived |
| `perf/lesson-word-style-owner` | `59627138e44d` | Archived |
| `perf/paragraph-paint-publication` | `4ed9cafe66f3` | Already integrated |
| `perf/pin6-combined-caret` | `61dad648ab8e` | Archived |
| `perf/pin6-first-caret` | `373b4a8b7d46` | Archived |
| `perf/pin8-cell-widget-cumulative` | `d3b6e56611b8` | Archived |
| `perf/pin8-equation-layout` | `f543bc948ccb` | Archived |
| `perf/plain-probe-shaping` | `27982e6e0dc9` | Archived |
| `perf/plain-typing-publication` | `e8ad6c699c7c` | Already integrated |
| `perf/retain-text-tail` | `fcce7756d6f7` | Archived |
| `perf/reuse-shared-character-keys` | `d391678aba62` | Archived |
| `perf/reuse-xml-validation` | `cbf547bbb220` | Archived |
| `perf/same-page-translation-reviewed` | `26e804385718` | Archived |
| `perf/same-sheet-paragraph-translation` | `487ff6bff1a2` | Archived |
| `perf/scoped-inline-integration` | `23557b799855` | Archived |
| `perf/scoped-inline-widgets` | `177e23c93e0e` | Archived |
| `perf/scroll-wheel-ahead` | `b2e98fbd9c6d` | Archived |
| `perf/shared-control-properties` | `e770c078caa1` | Archived |
| `perf/shared-history-combined` | `7b49c3d380d8` | Already integrated |
| `perf/shared-paste-lookups` | `e7ceab007d14` | Already integrated |
| `perf/stable-text-lines` | `187cbaa527d8` | Archived |
| `perf/stable-virtual-paper` | `7c0afb787acf` | Archived |
| `perf/sync-projection-runtime` | `0abd4f3a9e55` | Archived |
| `perf/table-furniture-size-owner` | `1d497de458a3` | Archived |
| `perf/verified-rapid-typing` | `2043769497f5` | Already integrated |
| `perf/viewport-distance-evaluation` | `e911e292b874` | Archived |
| `perf/viewport-distance-overscan` | `f5c7c5f39fa7` | Archived |
| `perf/word-scoped-opening` | `4c31ccab340a` | Already integrated |
| `perf/word-widget-height-owner` | `515f0973deaa` | Already integrated |
| `winpc/combined-layout` | `421383d0956a` | Already integrated |
| `winpc/combined-layout-final` | `59c4188a4914` | Already integrated |
| `winpc/fitting-layout` | `d4a1c0a302c6` | Archived |

## HarfBuzz cold opening and typing

Four paired WinPC trials compared the installed native baseline with the packed width measurement change.
Positive values mean the candidate took longer. All 16 exact text and caret checks passed.

| Document | Cold opening elapsed change | Cold opening CPU change | Typing CPU change |
| --- | ---: | ---: | ---: |
| Ordinary | +9.65 ms | +17.27 ms | −5.95 ms |
| Stored lesson | +15.80 ms | +18.51 ms | −1.60 ms |

Every stored lesson opening slowed. The experimental application keeps its previous package while we investigate.

WinPC ran 101 focused shaping and retained-selection tests against the consolidated source. All 565 assertions passed.
